using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;

namespace JwlMediaWin.Core
{
    internal sealed class FrameMotionDetector : IDisposable
    {
        // ===== P/Invoke =====
        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("gdi32.dll")]
        private static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest,
            int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, int dwRop);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        private const int SRCCOPY = 0x00CC0020;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        public enum ContentType
        {
            None,
            Image,
            Video
        }

        private readonly IntPtr _coreHwnd;
        private readonly int _sampleMs;
        private readonly int _windowedMinW;
        private readonly int _windowedMinH;
        private readonly int _motionHammingThreshold;
        private readonly int _consecutiveToDeclareVideo;
        private readonly int _hashHistory;

        private readonly Queue<ulong> _lastHashes = new Queue<ulong>();
        private readonly object _sync = new object();

        private Thread _thr;
        private CancellationTokenSource _cts;

        public ContentType CurrentType { get; private set; } = ContentType.None;

        /// <summary>Assine isto para espelhar no console/logger.</summary>
        public event Action<ContentType, string> OnLog;

        public FrameMotionDetector(
            IntPtr coreHwnd,
            int sampleMs = 400,
            int hashHistory = 8,
            int motionHammingThreshold = 10,
            int consecutiveToDeclareVideo = 3,
            int windowedMinW = 320,
            int windowedMinH = 240)
        {
            _coreHwnd = coreHwnd;
            _sampleMs = sampleMs;
            _hashHistory = Math.Max(2, hashHistory);
            _motionHammingThreshold = motionHammingThreshold;
            _consecutiveToDeclareVideo = Math.Max(2, consecutiveToDeclareVideo);
            _windowedMinW = windowedMinW;
            _windowedMinH = windowedMinH;
        }

        public void Start()
        {
            Stop();
            _cts = new CancellationTokenSource();
            _thr = new Thread(new ParameterizedThreadStart(ThreadLoop))
                { IsBackground = true, Name = "FrameMotionDetector" };
            _thr.Start(_cts.Token);
            Log("loop=started");
        }

        public void Stop()
        {
            try
            {
                if (_cts != null) _cts.Cancel();
                if (_thr != null) _thr.Join(1500);
            }
            catch
            {
                /* ignore */
            }
            finally
            {
                _thr = null;
                if (_cts != null) _cts.Dispose();
                _cts = null;
                lock (_sync)
                {
                    _lastHashes.Clear();
                    CurrentType = ContentType.None;
                }

                Log("loop=stopped");
            }
        }

        private void ThreadLoop(object state)
        {
            var ct = (CancellationToken)state;
            int consecutiveMotion = 0;

            while (!ct.IsCancellationRequested)
            {
                try
                {
                    string rectInfo;
                    ulong hash;
                    if (!TryCaptureAndHash(out hash, out rectInfo))
                    {
                        SetType(ContentType.None, "probe=None " + rectInfo, true);
                        Thread.Sleep(_sampleMs);
                        continue;
                    }

                    bool motion = false;
                    ulong prev = 0UL;

                    lock (_sync)
                    {
                        if (_lastHashes.Count > 0) prev = _lastHashes.Last();
                        _lastHashes.Enqueue(hash);
                        while (_lastHashes.Count > _hashHistory) _lastHashes.Dequeue();
                    }

                    if (prev != 0UL)
                    {
                        int hd = Hamming(prev, hash);
                        motion = hd >= _motionHammingThreshold;
                        Log(string.Format("tick hash=0x{0:X16} hamming={1} motion={2} {3}", hash, hd, motion,
                            rectInfo));
                    }
                    else
                    {
                        Log(string.Format("tick hash=0x{0:X16} first {1}", hash, rectInfo));
                    }

                    if (motion)
                    {
                        consecutiveMotion++;
                        if (consecutiveMotion >= _consecutiveToDeclareVideo)
                            SetType(ContentType.Video, "reason=consecutive-motion", true);
                        else
                            SetType(ContentType.Image, "reason=transient-motion", true);
                    }
                    else
                    {
                        consecutiveMotion = 0;
                        SetType(ContentType.Image, "reason=stable-frame", true);
                    }
                }
                catch (Exception ex)
                {
                    Log("error=" + ex.GetType().Name + " msg=" + ex.Message);
                }
                finally
                {
                    Thread.Sleep(_sampleMs);
                }
            }
        }

        private void SetType(ContentType t, string why, bool alwaysLog)
        {
            lock (_sync)
            {
                if (t != CurrentType)
                {
                    CurrentType = t;
                    Log("type_change=" + t + " " + why);
                }
                else if (alwaysLog)
                {
                    Log("type=" + t + " " + why);
                }
            }
        }

        private void Log(string msg)
        {
            var handler = OnLog;
            if (handler != null) handler(CurrentType, msg);
        }

        private bool TryCaptureAndHash(out ulong ahash, out string rectInfo)
        {
            ahash = 0UL;
            rectInfo = "";
            if (_coreHwnd == IntPtr.Zero)
            {
                rectInfo = "hwnd=0x0";
                return false;
            }

            RECT rc;
            if (!GetClientRect(_coreHwnd, out rc))
            {
                rectInfo = "getclientrect=false";
                return false;
            }

            int w = rc.Right - rc.Left;
            int h = rc.Bottom - rc.Top;
            rectInfo = "rect=" + w + "x" + h;

            if (w < _windowedMinW || h < _windowedMinH)
            {
                rectInfo += " small=true";
                return false;
            }

            POINT pt = new POINT { X = 0, Y = 0 };
            if (!ClientToScreen(_coreHwnd, ref pt))
            {
                rectInfo += " c2s=false";
                return false;
            }

            // Captura do desktop DC
            using (Bitmap bmp = new Bitmap(w, h, PixelFormat.Format24bppRgb))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    IntPtr hdcDest = g.GetHdc();
                    IntPtr hdcSrc = GetDC(IntPtr.Zero); // desktop DC
                    try
                    {
                        if (!BitBlt(hdcDest, 0, 0, w, h, hdcSrc, pt.X, pt.Y, SRCCOPY))
                        {
                            rectInfo += " bitblt=false";
                            return false;
                        }
                    }
                    finally
                    {
                        ReleaseDC(IntPtr.Zero, hdcSrc);
                        g.ReleaseHdc(hdcDest);
                    }
                }

                // Reduz pra 8x8 e gera aHash
                using (Bitmap small = new Bitmap(8, 8, PixelFormat.Format24bppRgb))
                {
                    using (Graphics g2 = Graphics.FromImage(small))
                    {
                        g2.DrawImage(bmp, 0, 0, 8, 8);
                    }

                    ulong sum = 0;
                    ulong bits = 0;
                    byte[] px = new byte[64];
                    for (int y = 0; y < 8; y++)
                    {
                        for (int x = 0; x < 8; x++)
                        {
                            Color c = small.GetPixel(x, y);
                            byte l = (byte)((c.R * 299 + c.G * 587 + c.B * 114) / 1000);
                            px[y * 8 + x] = l;
                            sum += l;
                        }
                    }

                    byte avg = (byte)(sum / 64);
                    for (int i = 0; i < 64; i++)
                        if (px[i] >= avg)
                            bits |= (1UL << i);

                    ahash = bits;
                }
            }

            return true;
        }

        private static int Hamming(ulong a, ulong b)
        {
            ulong v = a ^ b;
            int c = 0;
            while (v != 0)
            {
                v &= v - 1;
                c++;
            }

            return c;
        }

        public void Dispose()
        {
            Stop();
        }
    }
}