namespace JwlMediaWin.Core
{
    using System;
    using System.Diagnostics;
    using System.Threading;
    using System.Windows.Automation;
    using Models;
    using WindowsInput;
    using WindowsInput.Native;

    // ReSharper disable once UnusedMember.Global
    internal sealed class Fixer
    {
        private const string JwLibProcessName = "JWLibrary";
        private const string JwLibCaption = "JW Library";

        private const string JwLibSignLanguageProcessName = "JWLibrary.Forms.UWP";
        private const string JwLibSignLanguageCaption = "JW Library Sign Language";

        private AutomationElement _cachedDesktopElement;
        private MediaAndCoreWindows _cachedWindowElements;

        // Estado simples para detectar mudança de conteúdo entre ciclos
        private string _lastContentType = "Unknown";
        private bool _lastHasContent = false;

        // ===== Monitoramento contínuo (STA) =====
        private Thread _watchThread;
        private volatile bool _watchRunning;
        private IntPtr _watchCoreHwnd = IntPtr.Zero;
        private IntPtr _watchHostHwnd = IntPtr.Zero;

        // >>> Detector de movimento (hash de frames)
        private FrameMotionDetector _motion;

        private void MotionLog(FrameMotionDetector.ContentType type, string msg)
        {
            try
            {
                var line = "[MOTION] type=" + type + " " + msg;
                Debug.WriteLine(line);
                if (VERBOSE_LOG) Log(line);
            }
            catch
            {
                /* ignore */
            }
        }

        // opcional: se quiser jogar as mensagens no seu Status atual (INF)
        public static Action<string> StatusSink;

        // ===== Verbose logging infra =====
        private const bool VERBOSE_LOG = true;
        private static readonly object _logLock = new object();
        private static System.IO.StreamWriter _logWriter;

        static Fixer()
        {
            try
            {
                var path = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    "JwlMediaWin.uia.log");
                _logWriter = new System.IO.StreamWriter(path, true) { AutoFlush = true };
                _ = DateTime.Now; // só pra breakpoint fácil
                Log("==== UIA LOG START ====");
            }
            catch
            {
                /* best-effort */
            }
        }

        private static void Log(string msg)
        {
            try
            {
                var line = $"[{DateTime.Now:HH:mm:ss.fff}] {msg}";
                Debug.WriteLine(line);
                lock (_logLock)
                {
                    _logWriter?.WriteLine(line);
                }
            }
            catch
            {
                /* ignore */
            }
        }

        private static string Ct(AutomationElement el)
        {
            try
            {
                return el.Current.ControlType?.ProgrammaticName ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static string Cls(AutomationElement el)
        {
            try
            {
                return el.Current.ClassName ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static string Aid(AutomationElement el)
        {
            try
            {
                return el.Current.AutomationId ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static string Nm(AutomationElement el)
        {
            try
            {
                return el.Current.Name ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static int Pid(AutomationElement el)
        {
            try
            {
                return el.Current.ProcessId;
            }
            catch
            {
                return -1;
            }
        }

        private static bool Off(AutomationElement el)
        {
            try
            {
                return (bool)el.GetCurrentPropertyValue(AutomationElement.IsOffscreenProperty, true);
            }
            catch
            {
                return false;
            }
        }

        private static System.Windows.Rect Bnd(AutomationElement el, out bool ok)
        {
            try
            {
                ok = true;
                return (System.Windows.Rect)el.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty,
                    true);
            }
            catch
            {
                ok = false;
                return new System.Windows.Rect(0, 0, 0, 0);
            }
        }

        private static string RectStr(AutomationElement el)
        {
            var r = Bnd(el, out var ok);
            return ok ? $"{(int)r.X},{(int)r.Y},{(int)r.Width}x{(int)r.Height}" : "<n/a>";
        }

        private sealed class Counter<T>
        {
            public readonly System.Collections.Generic.Dictionary<T, int> Map =
                new System.Collections.Generic.Dictionary<T, int>();

            public void Add(T key)
            {
                if (Map.TryGetValue(key, out var n)) Map[key] = n + 1;
                else Map[key] = 1;
            }

            public System.Collections.Generic.IEnumerable<System.Collections.Generic.KeyValuePair<T, int>> Top(int k)
            {
                foreach (var kv in System.Linq.Enumerable.Take(
                             System.Linq.Enumerable.OrderByDescending(Map, x => x.Value), k))
                    yield return kv;
            }
        }

        private sealed class Biggest
        {
            public struct Item
            {
                public AutomationElement El;
                public double Area;
                public System.Windows.Rect R;
            }

            private readonly System.Collections.Generic.List<Item> _it = new System.Collections.Generic.List<Item>();
            private readonly int _cap;

            public Biggest(int cap)
            {
                _cap = cap;
            }

            public void Consider(AutomationElement el)
            {
                var r = Bnd(el, out var ok);
                if (!ok) return;
                var area = r.Width * r.Height;
                if (area <= 1) return;
                _it.Add(new Item { El = el, Area = area, R = r });
            }

            public System.Collections.Generic.IEnumerable<Item> Top()
            {
                return System.Linq.Enumerable.Take(System.Linq.Enumerable.OrderByDescending(_it, x => x.Area), _cap);
            }
        }

        private void DumpRawSnapshot(AutomationElement root, int maxNodes = 5000, int maxDepth = 14)
        {
            if (!VERBOSE_LOG || root == null) return;
            try
            {
                Log($"[SNAP] RawView dump start pid={Pid(root)} name='{Nm(root)}' cls='{Cls(root)}' ct='{Ct(root)}'");
                var walker = TreeWalker.RawViewWalker;
                var q = new System.Collections.Generic.Queue<(AutomationElement el, int depth)>();
                q.Enqueue((root, 0));
                int visited = 0;

                var byClass = new Counter<string>();
                var byCt = new Counter<string>();
                var byAid = new Counter<string>();

                var biggestImages = new Biggest(10);
                var biggestAnything = new Biggest(10);

                while (q.Count > 0 && visited < maxNodes)
                {
                    var (cur, d) = q.Dequeue();
                    if (cur == null) continue;
                    visited++;

                    var cls = Cls(cur);
                    var ct = Ct(cur);
                    var aid = Aid(cur);

                    byClass.Add(cls);
                    byCt.Add(ct);
                    if (!string.IsNullOrEmpty(aid)) byAid.Add(aid);

                    if (cur.Current.ControlType == ControlType.Image) biggestImages.Consider(cur);
                    biggestAnything.Consider(cur);

                    // marcos importantes
                    if (string.Equals(aid, "ImageRenderer", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(aid, "PlaybackButtonImage", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(aid, "PlayPauseButton", StringComparison.OrdinalIgnoreCase))
                    {
                        Log(
                            $"[SNAP] HIT aid='{aid}' cls='{cls}' ct='{ct}' off={Off(cur)} rect={RectStr(cur)} pid={Pid(cur)} name='{Nm(cur)}'");
                    }

                    if (d < maxDepth)
                    {
                        try
                        {
                            var ch = walker.GetFirstChild(cur);
                            while (ch != null)
                            {
                                q.Enqueue((ch, d + 1));
                                ch = walker.GetNextSibling(ch);
                            }
                        }
                        catch
                        {
                            /* ignore */
                        }
                    }
                }

                Log($"[SNAP] visited={visited} depth<= {maxDepth}");
                Log("[SNAP] Top ClassName:");
                foreach (var kv in byClass.Top(15)) Log($"  {kv.Key}: {kv.Value}");
                Log("[SNAP] Top ControlType:");
                foreach (var kv in byCt.Top(15)) Log($"  {kv.Key}: {kv.Value}");
                Log("[SNAP] Top AutomationId:");
                foreach (var kv in byAid.Top(15)) Log($"  {kv.Key}: {kv.Value}");

                Log("[SNAP] Biggest Images:");
                foreach (var it in biggestImages.Top())
                    Log(
                        $"  IMG area={(int)it.Area} rect={(int)it.R.X},{(int)it.R.Y},{(int)it.R.Width}x{(int)it.R.Height} cls='{Cls(it.El)}' aid='{Aid(it.El)}' name='{Nm(it.El)}' off={Off(it.El)}");

                Log("[SNAP] Biggest Elements (any):");
                foreach (var it in biggestAnything.Top())
                    Log(
                        $"  ANY area={(int)it.Area} rect={(int)it.R.X},{(int)it.R.Y},{(int)it.R.Width}x{(int)it.R.Height} ct='{Ct(it.El)}' cls='{Cls(it.El)}' aid='{Aid(it.El)}' name='{Nm(it.El)}' off={Off(it.El)}");

                Log("[SNAP] RawView dump end");
            }
            catch (Exception ex)
            {
                Log("[SNAP] dump error: " + ex.GetType().Name + " | " + ex.Message);
            }
        }


        private static AutomationElement FromHwndSafe(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return null;
            try
            {
                return AutomationElement.FromHandle(hwnd);
            }
            catch
            {
                return null;
            }
        }

        private static void SafeNotify(string msg)
        {
            try
            {
                if (StatusSink != null) StatusSink(msg);
                else Debug.WriteLine("[JWL] " + msg);
            }
            catch
            {
            }
        }

        private void StartMediaWatcher(IntPtr coreHwnd, IntPtr hostHwnd)
        {
            if (VERBOSE_LOG)
            {
                Log(string.Format("[WATCH] start coreHwnd=0x{0:X} hostHwnd=0x{1:X}",
                    coreHwnd.ToInt64(), hostHwnd.ToInt64()));
            }

            StopMediaWatcher(); // garante restart limpo

            _watchCoreHwnd = coreHwnd;
            _watchHostHwnd = hostHwnd;

            // === INTEGRAÇÃO: iniciar detector de movimento, se tivermos coreHwnd válido ===
            if (_watchCoreHwnd != IntPtr.Zero)
            {
                try
                {
                    _motion = new FrameMotionDetector(_watchCoreHwnd, /* sampleMs */ 500, /* hashHistory */ 8,
                        /* hammingThreshold */ 10, /* consecutiveVideo */ 3,
                        /* minW */ 320, /* minH */ 240);
                    _motion.OnLog += MotionLog;
                    _motion.Start();
                    if (VERBOSE_LOG) Log("[WATCH] motion detector started");
                }
                catch (Exception ex)
                {
                    Log("[WATCH] motion start failed: " + ex.GetType().Name + " | " + ex.Message);
                }
            }

            _watchRunning = true;
            _watchThread = new Thread(MediaWatchLoop);
            try
            {
                _watchThread.SetApartmentState(ApartmentState.STA);
            }
            catch
            {
                /* ignore */
            }

            _watchThread.IsBackground = true;
            _watchThread.Start();
        }

        private void StopMediaWatcher()
        {
            _watchRunning = false;
            var t = _watchThread;
            _watchThread = null;
            if (t != null)
            {
                try { t.Join(1500); } catch { }
            }

            // === INTEGRAÇÃO: parar e descartar detector de movimento ===
            try
            {
                if (_motion != null)
                {
                    _motion.Stop();
                    _motion.OnLog -= MotionLog;
                    _motion.Dispose();
                    _motion = null;
                    if (VERBOSE_LOG) Log("[WATCH] motion detector stopped");
                }
            }
            catch { /* ignore */ }

            _watchCoreHwnd = IntPtr.Zero;
            _watchHostHwnd = IntPtr.Zero;
        }

        /// <summary>
        /// Tenta escolher a melhor raiz para sondagem:
        /// 1) Se houver core/host, usa o que existir.
        /// 2) Senão, tenta achar a CoreWindow do mesmo PID na RawView a partir do RootElement.
        /// 3) Se nada der, usa o RootElement (o filtro por PID dentro do Probe tenta limitar).
        /// </summary>
        private static AutomationElement PickProbeRoot(AutomationElement coreEl, AutomationElement hostEl)
        {
            var seed = coreEl ?? hostEl;
            if (seed != null) return seed;

            // Se não temos seed, tentamos deduzir PID pelo processo padrão (JWLibrary / ApplicationFrameHost).
            // PRIORIDADE: JWLibrary; FALLBACK: ApplicationFrameHost contendo caption.
            int pid = -1;
            try
            {
                var procs = Process.GetProcessesByName("JWLibrary");
                if (procs.Length > 0) pid = procs[0].Id;
                else
                {
                    var af = Process.GetProcessesByName("ApplicationFrameHost");
                    if (af.Length > 0) pid = af[0].Id; // melhor do que nada
                }
            }
            catch
            {
                /* ignore */
            }

            var root = AutomationElement.RootElement;
            if (root == null) return null;

            // Se soubermos o PID, tentamos achar a CoreWindow com o mesmo PID
            if (pid > 0)
            {
                var hit = FindRawFirst(root, el =>
                {
                    try
                    {
                        if (SafePid(el) != pid) return false;
                        var cls = el.Current.ClassName ?? "";
                        return string.Equals(cls, "Windows.UI.Core.CoreWindow", StringComparison.OrdinalIgnoreCase);
                    }
                    catch
                    {
                        return false;
                    }
                }, maxNodes: 8000, maxDepth: 16);

                if (hit != null) return hit;
            }

            var chosen = root /* o que você vai retornar */;
            if (VERBOSE_LOG && chosen != null)
            {
                Log(
                    $"[ROOT] PickProbeRoot -> pid={Pid(chosen)} name='{Nm(chosen)}' cls='{Cls(chosen)}' ct='{Ct(chosen)}'");
            }

            return chosen;
        }


        private void MediaWatchLoop()
        {
            // roda em STA
            int nullStreak = 0;

            while (_watchRunning)
            {
                try
                {
                    var coreEl = FromHwndSafe(_watchCoreHwnd);
                    var hostEl = FromHwndSafe(_watchHostHwnd);

                    if (coreEl == null && hostEl == null)
                    {
                        // não derruba o watcher de cara — dá algumas tentativas
                        nullStreak++;
                    }
                    else
                    {
                        nullStreak = 0;
                    }

                    if (nullStreak >= 5)
                    {
                        // só para depois de 5 checks vazios seguidos
                        _watchRunning = false;
                        break;
                    }

                    // >>> NOVO: escolhe uma “raiz” boa para sondar conteúdo
                    var root = PickProbeRoot(coreEl, hostEl);
                    if (VERBOSE_LOG)
                    {
                        var info = root != null
                            ? $"pid={Pid(root)} name='{Nm(root)}' cls='{Cls(root)}' ct='{Ct(root)}' off={Off(root)} rect={RectStr(root)}"
                            : "<null>";
                        Log("[WATCH] probe root: " + info);
                    }

                    var type = ProbeContentTypeLabel(root);
                    var has = string.Equals(type, "Video", StringComparison.OrdinalIgnoreCase)
                              || string.Equals(type, "Image", StringComparison.OrdinalIgnoreCase);
                    if (VERBOSE_LOG)
                    {
                        Log($"[WATCH] detected type='{type}' has={has}");
                    }

                    if (has != _lastHasContent ||
                        !string.Equals(type, _lastContentType, StringComparison.OrdinalIgnoreCase))
                    {
                        _lastHasContent = has;
                        _lastContentType = type;

                        var msg = has
                            ? $"JW Library: exibindo {type.ToLowerInvariant()}."
                            : "JW Library: no media.";
                        SafeNotify(msg);

                        // quando muda de estado, fotografa o RawView
                        if (has != _lastHasContent ||
                            !string.Equals(type, _lastContentType, StringComparison.OrdinalIgnoreCase))
                        {
                            DumpRawSnapshot(root, maxNodes: 6000, maxDepth: 16);
                        }
                    }
                }
                catch
                {
                    // UIA pode falhar durante transições – ignora este tick
                }

                try
                {
                    Thread.Sleep(1000);
                }
                catch
                {
                }
            }
        }

        /// <summary>
        /// Executa o "fixer": encontra a janela de mídia do JWL e a ajusta.
        /// </summary>
        public FixerStatus Execute(JwLibAppTypes appType, bool topMost)
        {
            try
            {
                string processName;
                string caption;

                switch (appType)
                {
                    case JwLibAppTypes.None:
                        throw new Exception("Expected app type!");

                    default:
                    case JwLibAppTypes.JwLibrary:
                        processName = JwLibProcessName;
                        caption = JwLibCaption;
                        break;

                    case JwLibAppTypes.JwLibrarySignLanguage:
                        processName = JwLibSignLanguageProcessName;
                        caption = JwLibSignLanguageCaption;
                        break;
                }

                return ExecuteInternal(appType, topMost, processName, caption);
            }
            catch (ElementNotAvailableException)
            {
                return new FixerStatus { ErrorIsTransitioning = true };
            }
            catch (Exception ex)
            {
                // Loga, mas não derruba totalmente. Devolve um status “neutro”.
                Debug.WriteLine("[JWL][Execute] erro inesperado: " + ex.GetType().Name + " | " + ex.Message);
                return new FixerStatus { ErrorUnknown = false }; // <- evita “Unknown error” no UI
            }
        }

        // ===========================
        // ==== Fluxo principal ======
        private FixerStatus ExecuteInternal(
            JwLibAppTypes appType,
            bool topMost,
            string processName,
            string caption)
        {
            var result = new FixerStatus { FindWindowResult = GetMediaAndCoreWindow(appType, processName, caption) };

            if (!result.FindWindowResult.FoundMediaWindow)
            {
                FillContentStatus(result);
                StopMediaWatcher(); // nada pra monitorar
                return result;
            }

            if (result.FindWindowResult.IsAlreadyFixed)
            {
                // já temos os handles: iniciar watcher mesmo assim
                var mh = (IntPtr)result.FindWindowResult.MainMediaWindow.Current.NativeWindowHandle;
                var ch = (IntPtr)result.FindWindowResult.CoreMediaWindow.Current.NativeWindowHandle;
                var coreAE = FromHwndSafe(ch); // ou AutomationElement.FromHandle(coreHandle)
                if (coreAE != null)
                {
                    Debug.WriteLine(
                        $"[JWL][Watch] Core root: pid={coreAE.Current.ProcessId} name='{coreAE.Current.Name}' class='{coreAE.Current.ClassName}'");
                }
                else
                {
                    Debug.WriteLine("[JWL][Watch] Core root: <null>");
                }

                StartMediaWatcher(ch, mh);

                FillContentStatus(result);
                return result;
            }

            var mainHandle = (IntPtr)result.FindWindowResult.MainMediaWindow.Current.NativeWindowHandle;
            var coreHandle = (IntPtr)result.FindWindowResult.CoreMediaWindow.Current.NativeWindowHandle;

            Thread.Sleep(1000); // margem de segurança

            NativeMethods.SetForegroundWindow(coreHandle);

            var rect = result.FindWindowResult.MainMediaWindow.Current.BoundingRectangle;

            result.FindWindowResult.CoreMediaWindow.SetFocus();

            if (!result.FindWindowResult.CoreMediaWindow.Current.HasKeyboardFocus)
            {
                FillContentStatus(result);
                return result;
            }

            result.CoreWindowFocused = true;

            // Converter a janela (Win+Shift+Enter)
            if (!ConvertMediaWindow(result.FindWindowResult.MainMediaWindow))
            {
                // <-- ADICIONE estas 2 linhas ANTES de retornar
                StartMediaWatcher(coreHandle, mainHandle); // mesmo sem “fix”, ainda podemos monitorar conteúdo
                FillContentStatus(result);
                return result;
            }

            var insertAfterValue = topMost
                ? new IntPtr(-1)
                : new IntPtr(-2);

            const uint ShowWindowFlag = 0x0040;
            const uint NoCopyBitsFlag = 0x0100;
            const uint NoSendChangingFlag = 0x0400;

            const int adjustment = 1; // titlebar transparente
            const int border = 8; // bordas

            Thread.Sleep(500); // sem isso o SetWindowPos às vezes não pega

            NativeMethods.SetWindowPos(
                mainHandle,
                insertAfterValue,
                (int)rect.Left - border,
                (int)rect.Top - adjustment,
                (int)rect.Width + (border * 2),
                (int)rect.Height + (adjustment + border),
                (int)(NoCopyBitsFlag | NoSendChangingFlag | ShowWindowFlag));

            result.IsFixed = true;

            EnsureWindowIsNonSizeable(mainHandle);

            // Desabilita a janela de mídia para "remover" os botões da titlebar transparente
            NativeMethods.SetForegroundWindow(coreHandle);
            NativeMethods.EnableWindow(mainHandle, false);

            // Preenche status de conteúdo (vídeo/imagem/none) e mudança
            FillContentStatus(result);

            // iniciar vigilância contínua (monitoramento)
            StartMediaWatcher(coreHandle, mainHandle);

            return result;
        }

        private FindWindowResult GetMediaAndCoreWindow(
            JwLibAppTypes appType, string processName, string caption)
        {
            var result = new FindWindowResult();

            if (Process.GetProcessesByName(processName).Length == 0)
            {
                return result;
            }

            result.JwlRunning = true;

            CacheDesktopElement();
            if (_cachedDesktopElement == null)
            {
                return result;
            }

            result.FoundDesktop = true;

            if (_cachedWindowElements == null)
            {
                _cachedWindowElements = GetMediaAndCoreWindowsInternal(appType, caption, processName);
            }

            if (_cachedWindowElements != null)
            {
                result.FoundMediaWindow = true;

                result.CoreMediaWindow = _cachedWindowElements.CoreWindow;
                result.MainMediaWindow = _cachedWindowElements.MediaWindow;

                try
                {
                    // TransformPattern só aparece após “fix”
                    if (HasTransformPattern(_cachedWindowElements.MediaWindow))
                    {
                        result.IsAlreadyFixed = true;
                    }
                }
                catch (ElementNotAvailableException)
                {
                    _cachedWindowElements = null;
                    _cachedDesktopElement = null;

                    return new FindWindowResult
                    {
                        FoundDesktop = true,
                        JwlRunning = true
                    };
                }
            }

            return result;
        }

        private MediaAndCoreWindows GetMediaAndCoreWindowsInternal(
            JwLibAppTypes appType, string caption, string processName)
        {
            var candidateMediaWindows = _cachedDesktopElement.FindAll(
                TreeScope.Children,
                new PropertyCondition(AutomationElement.IsEnabledProperty, true));

            if (candidateMediaWindows.Count == 0)
            {
                return null;
            }

            // Caminho principal: hosts pertencentes ao processo JWLibrary/JWLibrary.Forms.UWP ou ApplicationFrameHost com caption
            for (int i = 0; i < candidateMediaWindows.Count; i++)
            {
                var candidate = (AutomationElement)candidateMediaWindows[i];

                if (!IsFromProcessOrHost(candidate, processName, caption))
                {
                    continue;
                }

                var coreWindow = GetJwlCoreWindow(candidate, caption);
                if (coreWindow != null && IsCorrectCoreWindow(appType, coreWindow))
                {
                    Debug.WriteLine("[JWL][Find] host=" + candidate.Current.ClassName +
                                    " pid=" + candidate.Current.ProcessId);
                    return new MediaAndCoreWindows
                    {
                        CoreWindow = coreWindow,
                        MediaWindow = candidate
                    };
                }
            }

            // Fallback: varrer Subtree por qualquer elemento do processo alvo e achar CoreWindow
            try
            {
                var all = _cachedDesktopElement.FindAll(
                    TreeScope.Subtree,
                    new PropertyCondition(AutomationElement.IsEnabledProperty, true));

                for (int i = 0; i < all.Count; i++)
                {
                    var el = (AutomationElement)all[i];
                    if (!IsFromProcessOrHost(el, processName, caption))
                    {
                        continue;
                    }

                    var coreWindow = FindCoreWindowDescendant(el);
                    if (coreWindow != null && IsCorrectCoreWindow(appType, coreWindow))
                    {
                        Debug.WriteLine("[JWL][Find][fallback] host=" + el.Current.ClassName +
                                        " pid=" + el.Current.ProcessId);
                        return new MediaAndCoreWindows
                        {
                            CoreWindow = coreWindow,
                            MediaWindow = el
                        };
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        // ===========================
        // ====== Detecção UIA =======
        // ===========================

        private static bool ConvertMediaWindow(AutomationElement mainMediaWindow)
        {
            if (mainMediaWindow == null)
            {
                return false;
            }

            var inputSim = new InputSimulator();

            inputSim.Keyboard.ModifiedKeyStroke(
                new[]
                {
                    VirtualKeyCode.LWIN,
                    VirtualKeyCode.SHIFT
                },
                VirtualKeyCode.RETURN);

            var counter = 0;
            while (!HasTransformPattern(mainMediaWindow) && counter < 20)
            {
                Thread.Sleep(200);
                ++counter;
            }

            return HasTransformPattern(mainMediaWindow);
        }

        private static bool IsCorrectCoreWindow(JwLibAppTypes appType, AutomationElement coreMediaWindow)
        {
            if (coreMediaWindow == null)
            {
                return false;
            }

            switch (appType)
            {
                default:
                case JwLibAppTypes.None:
                    return false;

                case JwLibAppTypes.JwLibrary:
                    return GetWebView(coreMediaWindow) != null ||
                           HasNoChildren(coreMediaWindow);

                case JwLibAppTypes.JwLibrarySignLanguage:
                    return GetImageControl(coreMediaWindow) != null;
            }
        }

        private static bool HasNoChildren(AutomationElement coreMediaWindow)
        {
            if (coreMediaWindow == null)
            {
                return false;
            }

            return coreMediaWindow.FindFirst(
                TreeScope.Children,
                new PropertyCondition(AutomationElement.IsEnabledProperty, true)) == null;
        }

        private static AutomationElement GetJwlCoreWindow(AutomationElement mainJwlWindow, string caption)
        {
            if (mainJwlWindow == null)
            {
                return null;
            }

            // Procura CoreWindow por classe na subárvore (mais robusto)
            var core = FindCoreWindowDescendant(mainJwlWindow);
            if (core != null) return core;

            // Fallback antigo: Name + classe
            var condition = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, caption),
                new PropertyCondition(AutomationElement.ClassNameProperty, "Windows.UI.Core.CoreWindow"));

            return mainJwlWindow.FindFirst(TreeScope.Children, condition);
        }

        private static AutomationElement GetWebView(AutomationElement coreJwlWindow)
        {
            if (coreJwlWindow == null)
            {
                return null;
            }

            return coreJwlWindow.FindFirst(
                TreeScope.Children,
                new PropertyCondition(AutomationElement.ClassNameProperty, "Microsoft.UI.Xaml.Controls.WebView2"));
        }

        private static AutomationElement GetImageControl(AutomationElement coreJwlWindow)
        {
            if (coreJwlWindow == null)
            {
                return null;
            }

            return coreJwlWindow.FindFirst(
                TreeScope.Children,
                new PropertyCondition(AutomationElement.ClassNameProperty, "Image"));
        }

        private static bool HasTransformPattern(AutomationElement item)
        {
            if (item == null)
            {
                return false;
            }

            return (bool)item.GetCurrentPropertyValue(AutomationElement.IsTransformPatternAvailableProperty);
        }

        private static bool HasWindowPattern(AutomationElement item)
        {
            if (item == null)
            {
                return false;
            }

            return (bool)item.GetCurrentPropertyValue(AutomationElement.IsWindowPatternAvailableProperty);
        }

        private static void EnsureWindowIsNonSizeable(IntPtr mainHandle)
        {
            const int GWL_STYLE = -16;
            const int WS_SIZEBOX = 0x040000;

            var val = (int)NativeMethods.GetWindowLongPtr(mainHandle, GWL_STYLE) & ~WS_SIZEBOX;
            NativeMethods.SetWindowLongPtr(mainHandle, GWL_STYLE, (IntPtr)val);
        }

        private static bool IsFromProcessOrHost(AutomationElement item, string processName, string caption)
        {
            if (item == null) return false;
            try
            {
                var pid = item.Current.ProcessId;
                var p = Process.GetProcessById(pid);
                var proc = p.ProcessName;

                if (string.Equals(proc, processName, StringComparison.OrdinalIgnoreCase))
                    return true;

                // UWP hospedada
                if (string.Equals(proc, "ApplicationFrameHost", StringComparison.OrdinalIgnoreCase))
                {
                    var name = item.Current.Name ?? string.Empty;
                    if (!string.IsNullOrEmpty(caption) &&
                        name.IndexOf(caption, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private static AutomationElement FindCoreWindowDescendant(AutomationElement root)
        {
            if (root == null) return null;
            try
            {
                var cond = new PropertyCondition(
                    AutomationElement.ClassNameProperty,
                    "Windows.UI.Core.CoreWindow");
                return root.FindFirst(TreeScope.Descendants, cond);
            }
            catch
            {
                return null;
            }
        }

        private void CacheDesktopElement()
        {
            if (_cachedDesktopElement == null)
            {
                _cachedDesktopElement = AutomationElement.RootElement;
            }
        }

        // ===========================
        // ====== Conteúdo/Mídia =====
        // ===========================

        private static bool IsVisible(AutomationElement el)
        {
            if (el == null) return false;
            try
            {
                var off = (bool)el.GetCurrentPropertyValue(AutomationElement.IsOffscreenProperty, true);
                if (off) return false;

                var rectObj = el.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty, true);
                var r = (System.Windows.Rect)rectObj;
                return r.Width > 5 && r.Height > 5;
            }
            catch
            {
                return true;
            }
        }

        private static string NormalizeNoDiacriticsLower(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var formD = s.Normalize(System.Text.NormalizationForm.FormD);
            var sb = new System.Text.StringBuilder(formD.Length);
            for (int i = 0; i < formD.Length; i++)
            {
                var uc = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(formD[i]);
                if (uc != System.Globalization.UnicodeCategory.NonSpacingMark)
                    sb.Append(formD[i]);
            }

            return sb.ToString().Normalize(System.Text.NormalizationForm.FormC).ToLowerInvariant();
        }

        private static bool NameContainsAllTokens(AutomationElement el, string[] tokens)
        {
            if (el == null) return false;
            string name = "";
            try
            {
                name = el.Current.Name ?? "";
            }
            catch
            {
            }

            var norm = NormalizeNoDiacriticsLower(name);
            for (int i = 0; i < tokens.Length; i++)
            {
                if (norm.IndexOf(tokens[i], StringComparison.Ordinal) < 0) return false;
            }

            return true;
        }

        private static AutomationElement FindAddToPlaylist(AutomationElement root)
        {
            if (root == null) return null;
            try
            {
                var tokens = new[] { "adicionar", "playlist" };

                // 1) Buttons com Name que contenha os tokens
                var condButton = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                var buttons = root.FindAll(TreeScope.Descendants, condButton);
                for (int i = 0; i < buttons.Count; i++)
                {
                    var el = (AutomationElement)buttons[i];
                    if (NameContainsAllTokens(el, tokens))
                        return el;
                }

                // 2) Text com os tokens -> sobe ao ancestral Button/AppBarButton
                var condText = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                var texts = root.FindAll(TreeScope.Descendants, condText);
                for (int i = 0; i < texts.Count; i++)
                {
                    var txt = (AutomationElement)texts[i];
                    if (!NameContainsAllTokens(txt, tokens)) continue;

                    var btn = GetAncestorButton(txt);
                    if (btn != null)
                        return btn;
                }

                // 3) fallback suave: qualquer elemento cujo Name contenha tokens + ancestral Button
                var condAll = new PropertyCondition(AutomationElement.IsControlElementProperty, true);
                var any = root.FindAll(TreeScope.Descendants, condAll);
                for (int i = 0; i < any.Count; i++)
                {
                    var el = (AutomationElement)any[i];
                    if (!NameContainsAllTokens(el, tokens)) continue;
                    var btn = GetAncestorButton(el);
                    if (btn != null)
                        return btn;
                }
            }
            catch
            {
            }

            return null;
        }

        private static string ProbeContentTypeLabel(AutomationElement root)
        {
            if (root == null) return "Unknown";

            var pid = SafePid(root);

            if (VERBOSE_LOG)
            {
                Log($"[PROBE] start pid={Pid(root)} name='{Nm(root)}' cls='{Cls(root)}' ct='{Ct(root)}'");
            }

            bool IsSameProc(AutomationElement el)
            {
                if (pid <= 0) return true;
                var p = SafePid(el);
                return p == pid;
            }

            // Sinal forte de vídeo: botão Play/Pause padrão dos UWP MTC
            var playPauseBtn = FindRawFirst(root, el =>
            {
                try
                {
                    // precisa ser do mesmo processo (ou sem filtro se não há pid)
                    var pidRoot = SafePid(root);
                    if (pidRoot > 0 && SafePid(el) != pidRoot) return false;

                    if (el.Current.ControlType != ControlType.Button) return false;

                    var autoId = el.Current.AutomationId ?? "";
                    if (string.Equals(autoId, "PlayPauseButton", StringComparison.OrdinalIgnoreCase))
                    {
                        var off = (bool)el.GetCurrentPropertyValue(AutomationElement.IsOffscreenProperty, true);
                        if (off) return false;
                        var rect = (System.Windows.Rect)el.GetCurrentPropertyValue(
                            AutomationElement.BoundingRectangleProperty, true);
                        return rect.Width >= 10 && rect.Height >= 10;
                    }

                    // Nomes localizados (fallback)
                    var nm = (el.Current.Name ?? string.Empty).ToLowerInvariant();
                    if (nm.Contains("pausar") || nm.Contains("reproduzir") || nm.Contains("play") ||
                        nm.Contains("pause"))
                    {
                        var off = (bool)el.GetCurrentPropertyValue(AutomationElement.IsOffscreenProperty, true);
                        if (off) return false;
                        var rect = (System.Windows.Rect)el.GetCurrentPropertyValue(
                            AutomationElement.BoundingRectangleProperty, true);
                        return rect.Width >= 10 && rect.Height >= 10;
                    }

                    return false;
                }
                catch
                {
                    return false;
                }
            });

            if (playPauseBtn != null)
            {
                Log("[PROBE] Video via PlayPauseButton/Button name");
                Debug.WriteLine("[JWL][ProbeRaw] Video via PlayPauseButton/Button name");
                return "Video";
            }

            // 1) VÍDEO: procura na Raw View por Image com AutomationId "PlaybackButtonImage"
            var playbackImg = FindRawFirst(root, el =>
            {
                try
                {
                    if (!IsSameProc(el)) return false;
                    if (el.Current.ControlType != ControlType.Image) return false;
                    var id = el.Current.AutomationId ?? "";
                    var cls = el.Current.ClassName ?? "";
                    // id e classe vindos do seu dump
                    if (!string.Equals(id, "PlaybackButtonImage", StringComparison.OrdinalIgnoreCase)) return false;
                    if (!string.Equals(cls, "Image", StringComparison.OrdinalIgnoreCase)) return false;

                    // visibilidade básica
                    var off = (bool)el.GetCurrentPropertyValue(AutomationElement.IsOffscreenProperty, true);
                    if (off) return false;
                    var rect = (System.Windows.Rect)el.GetCurrentPropertyValue(
                        AutomationElement.BoundingRectangleProperty, true);
                    return rect.Width >= 10 && rect.Height >= 10;
                }
                catch
                {
                    return false;
                }
            });

            if (playbackImg != null)
            {
                Log("[PROBE] Video via PlaybackButtonImage");
                Debug.WriteLine("[JWL][ProbeRaw] Video via PlaybackButtonImage");
                return "Video";
            }

            // 2) IMAGEM: procura na Raw View por Image com AutomationId "ImageRenderer"
            var imageRenderer = FindRawFirst(root, el =>
            {
                try
                {
                    if (!IsSameProc(el)) return false;
                    if (el.Current.ControlType != ControlType.Image) return false;
                    var id = el.Current.AutomationId ?? "";
                    var cls = el.Current.ClassName ?? "";
                    if (!string.Equals(id, "ImageRenderer", StringComparison.OrdinalIgnoreCase)) return false;
                    if (!string.Equals(cls, "Image", StringComparison.OrdinalIgnoreCase)) return false;

                    var off = (bool)el.GetCurrentPropertyValue(AutomationElement.IsOffscreenProperty, true);
                    if (off) return false;
                    var rect = (System.Windows.Rect)el.GetCurrentPropertyValue(
                        AutomationElement.BoundingRectangleProperty, true);
                    // evita ícones pequenos
                    return rect.Width > 300 && rect.Height > 200;
                }
                catch
                {
                    return false;
                }
            });

            if (imageRenderer != null)
            {
                Log("[PROBE] Image via ImageRenderer");
                Debug.WriteLine("[JWL][ProbeRaw] Image via ImageRenderer");
                return "Image";
            }

            // 3) Fallbacks leves (ainda em Raw): classes típicas de vídeo
            string[] videoClasses =
                { "MediaPlayerPresenter", "MediaElement", "SwapChainPanel", "SwapChainPanelNativeWindowClass" };
            foreach (var clsName in videoClasses)
            {
                var hit = FindRawFirst(root, el =>
                {
                    try
                    {
                        if (!IsSameProc(el)) return false;
                        var cls = el.Current.ClassName ?? "";
                        if (!string.Equals(cls, clsName, StringComparison.OrdinalIgnoreCase)) return false;
                        var off = (bool)el.GetCurrentPropertyValue(AutomationElement.IsOffscreenProperty, true);
                        if (off) return false;
                        var rect = (System.Windows.Rect)el.GetCurrentPropertyValue(
                            AutomationElement.BoundingRectangleProperty, true);
                        return rect.Width > 50 && rect.Height > 50;
                    }
                    catch
                    {
                        return false;
                    }
                });

                if (hit != null)
                {
                    Log("[PROBE] Video via class " + clsName);
                    Debug.WriteLine("[JWL][ProbeRaw] Video via class " + clsName);
                    return "Video";
                }
            }

            // 4) Fallback imagem grande qualquer (Raw)
            var bigImg = FindRawFirst(root, el =>
            {
                try
                {
                    if (!IsSameProc(el)) return false;
                    if (el.Current.ControlType != ControlType.Image) return false;
                    var off = (bool)el.GetCurrentPropertyValue(AutomationElement.IsOffscreenProperty, true);
                    if (off) return false;
                    var rect = (System.Windows.Rect)el.GetCurrentPropertyValue(
                        AutomationElement.BoundingRectangleProperty, true);
                    return rect.Width > 400 && rect.Height > 300;
                }
                catch
                {
                    return false;
                }
            });

            if (bigImg != null)
            {
                Log("[PROBE] Image via large Image (fallback)");
                Debug.WriteLine("[JWL][ProbeRaw] Image via large Image (fallback)");
                return "Image";
            }

            // ===== Plano B: ControlView heuristics (fallback rápido) =====

            // (vídeo) botão/ícone de play/pause/volume/transport
            var condBtn = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
            var btns = root.FindAll(TreeScope.Descendants, condBtn);
            for (int i = 0; i < btns.Count; i++)
            {
                var b = (AutomationElement)btns[i];
                if (NameContainsAllTokens(b, new[] { "pausar" }) ||
                    NameContainsAllTokens(b, new[] { "reproduzir" }) ||
                    NameContainsAllTokens(b, new[] { "play" }) ||
                    NameContainsAllTokens(b, new[] { "pause" }))
                {
                    if (IsVisible(b))
                    {
                        Debug.WriteLine("[JWL][ProbeCtl] video via play/pause button name");
                        return "Video";
                    }
                }
            }

            // (vídeo) classes típicas
            string[] classes =
            {
                "MediaTransportControls", "MediaPlayerPresenter", "MediaElement", "SwapChainPanel",
                "SwapChainPanelNativeWindowClass"
            };
            for (int i = 0; i < classes.Length; i++)
            {
                var cond = new PropertyCondition(AutomationElement.ClassNameProperty, classes[i]);
                var el = root.FindFirst(TreeScope.Descendants, cond);
                if (IsVisible(el))
                {
                    Log("[PROBE] video via class " + classes[i]);
                    Debug.WriteLine("[JWL][ProbeCtl] video via class " + classes[i]);
                    return "Video";
                }
            }

            // (imagem) ImageRenderer ou imagem grande
            var condImageRendererCtl = new AndCondition(
                new PropertyCondition(AutomationElement.AutomationIdProperty, "ImageRenderer"),
                new PropertyCondition(AutomationElement.ClassNameProperty, "Image"),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Image)
            );
            var imgR = root.FindFirst(TreeScope.Descendants, condImageRendererCtl);
            if (IsVisible(imgR))
            {
                try
                {
                    var rr = (System.Windows.Rect)imgR.GetCurrentPropertyValue(
                        AutomationElement.BoundingRectangleProperty, true);
                    if (rr.Width > 300 && rr.Height > 200)
                    {
                        Log("[PROBE] image via ImageRenderer");
                        Debug.WriteLine("[JWL][ProbeCtl] image via ImageRenderer");
                        return "Image";
                    }
                }
                catch
                {
                    Log("[PROBE] image via ImageRenderer (no rect)");
                    Debug.WriteLine("[JWL][ProbeCtl] image via ImageRenderer (no rect)");
                    return "Image";
                }
            }

            var condAnyImgCtl = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Image);
            var anyImgCtl = root.FindFirst(TreeScope.Descendants, condAnyImgCtl);
            if (IsVisible(anyImgCtl))
            {
                try
                {
                    var r2 = (System.Windows.Rect)anyImgCtl.GetCurrentPropertyValue(
                        AutomationElement.BoundingRectangleProperty, true);
                    if (r2.Width > 400 && r2.Height > 300)
                    {
                        Debug.WriteLine("[JWL][ProbeCtl] image via large Image (fallback)");
                        return "Image";
                    }
                }
                catch
                {
                    Log("[PROBE] some Image visible, no rect");
                    Debug.WriteLine("[JWL][ProbeCtl] some Image visible, no rect");
                    return "Image";
                }
            }

            Log("[PROBE] None");
            Debug.WriteLine("[JWL][ProbeRaw] None");
            return "None";
        }

        private static int SafePid(AutomationElement el)
        {
            try
            {
                return el.Current.ProcessId;
            }
            catch
            {
                return -1;
            }
        }

        private static AutomationElement FindRawFirst(
            AutomationElement start,
            Predicate<AutomationElement> match,
            int maxNodes = 8000,
            int maxDepth = 16)
        {
            if (start == null || match == null) return null;

            var walker = TreeWalker.RawViewWalker;

            // BFS com (nó, profundidade)
            var q = new System.Collections.Generic.Queue<(AutomationElement el, int depth)>();
            q.Enqueue((start, 0));

            int visited = 0;

            while (q.Count > 0 && visited < maxNodes)
            {
                var (cur, depth) = q.Dequeue();
                if (cur == null) continue;
                visited++;

                try
                {
                    if (match(cur)) return cur;
                }
                catch
                {
                    // ignora exceção do predicado e continua
                }

                if (depth >= maxDepth) continue;

                try
                {
                    var child = walker.GetFirstChild(cur);
                    while (child != null)
                    {
                        q.Enqueue((child, depth + 1));
                        child = walker.GetNextSibling(child);
                    }
                }
                catch
                {
                    // alguns providers falham — segue o baile
                }
            }

            return null;
        }

        private static AutomationElement GetAncestorButton(AutomationElement el)
        {
            if (el == null) return null;
            try
            {
                var walker = TreeWalker.ControlViewWalker;
                var cur = el;
                // limita a subida para evitar loops
                for (int i = 0; i < 8 && cur != null; i++)
                {
                    cur = walker.GetParent(cur);
                    if (cur == null) break;

                    try
                    {
                        var ct = cur.Current.ControlType;
                        if (ct == ControlType.Button) return cur;

                        var cls = cur.Current.ClassName ?? "";
                        // alguns XAML controls comuns
                        if (cls == "AppBarButton" || cls == "AppBarToggleButton" || cls == "MenuFlyoutItem")
                            return cur;
                    }
                    catch
                    {
                        /* continua subindo */
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private static bool TryRunOnSta<T>(Func<T> func, out T result, int timeoutMs = 8000)
        {
            // NÃO capture 'result' na lambda. Use um local.
            T local = default(T);
            Exception error = null;
            var done = new System.Threading.ManualResetEvent(false);

            var th = new Thread(() =>
            {
                try
                {
                    local = func();
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                finally
                {
                    try
                    {
                        done.Set();
                    }
                    catch
                    {
                    }
                }
            });

            th.IsBackground = true;
            try
            {
                th.SetApartmentState(ApartmentState.STA);
            }
            catch
            {
                /* best-effort */
            }

            th.Start();

            if (!done.WaitOne(timeoutMs))
            {
                try
                {
                    th.Abort();
                }
                catch
                {
                }

                Debug.WriteLine("[JWL][STA] Timeout após " + timeoutMs + "ms");
                result = default(T);
                return false;
            }

            if (error != null)
            {
                Debug.WriteLine("[JWL][STA] Exceção: " + error.GetType().Name + " | " + error.Message);
                result = default(T);
                return false;
            }

            // Copia para o 'out' só aqui, fora da lambda
            result = local;
            return true;
        }

        private static T RunOnSta<T>(Func<T> func, int timeoutMs = 4000)
        {
            T result = default(T);
            Exception error = null;
            var done = new System.Threading.ManualResetEvent(false);

            var th = new Thread(() =>
            {
                try
                {
                    result = func();
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                finally
                {
                    done.Set();
                }
            });
            th.IsBackground = true;
            th.SetApartmentState(ApartmentState.STA);
            th.Start();

            if (!done.WaitOne(timeoutMs))
                throw new TimeoutException("UIA STA call timed out");

            if (error != null) throw error;
            return result;
        }

        private void FillContentStatus(FixerStatus result)
        {
            try
            {
                var contentRoot = result.FindWindowResult != null
                    ? (result.FindWindowResult.CoreMediaWindow ?? result.FindWindowResult.MainMediaWindow)
                    : null;

                string type;
                if (!TryRunOnSta(() => ProbeContentTypeLabel(contentRoot), out type, timeoutMs: 8000))
                {
                    // se der timeout/erro, não derruba o ciclo
                    type = "Unknown";
                    Debug.WriteLine("[JWL][Content] sondagem falhou (timeout/erro). Mantendo Unknown.");
                }

                var has = string.Equals(type, "Video", StringComparison.OrdinalIgnoreCase)
                          || string.Equals(type, "Image", StringComparison.OrdinalIgnoreCase);

                result.MediaContentType = type;
                result.MediaHasContent = has;
                result.MediaContentChanged = (has != _lastHasContent)
                                             || !string.Equals(type, _lastContentType,
                                                 StringComparison.OrdinalIgnoreCase);

                _lastHasContent = has;
                _lastContentType = type;

                Debug.WriteLine("[JWL][Content] now=" + type + " changed=" + result.MediaContentChanged);
            }
            catch (Exception ex)
            {
                // último guarda-chuva: nunca permita que uma exceção aqui vire "Unknown error"
                result.MediaContentType = "Unknown";
                result.MediaHasContent = false;
                result.MediaContentChanged = false;
                Debug.WriteLine("[JWL][Content] erro inesperado em FillContentStatus: " + ex.GetType().Name + " | " +
                                ex.Message);
            }
        }
    }
}