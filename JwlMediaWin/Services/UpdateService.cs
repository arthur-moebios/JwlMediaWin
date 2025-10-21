namespace JwlMediaWin.Services
{
    using System;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.IO;
    using System.Net.Http;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows;
    using Serilog;

    internal sealed class UpdateService : IUpdateService
    {
        public async Task CheckAndOfferAsync(CancellationToken ct = default)
        {
            try
            {
                var current = VersionDetection.GetCurrentVersion();
                var latest = await VersionDetection.GetLatestAsync(ct).ConfigureAwait(false);

                if (latest == null)
                {
                    Log.Warning("[Update] Could not retrieve latest release info.");
                    return;
                }

                if (latest.Version > current)
                {
                    Log.Information("[Update] Update available: {Latest} (local {Current})",
                        latest.Version, current);

                    // Ensure UI thread for MessageBox
                    var answer = MessageBoxResult.No;
                    if (Application.Current?.Dispatcher != null)
                    {
                        answer = Application.Current.Dispatcher.Invoke(() =>
                            MessageBox.Show(
                                $"A new version ({latest.Version}) is available. You have {current}.\n\nInstall now?",
                                "JwlMediaWin - Update available",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Information));
                    }
                    else
                    {
                        // Fallback if no WPF dispatcher (rare in this app)
                        answer = MessageBox.Show(
                            $"A new version ({latest.Version}) is available. You have {current}.\n\nInstall now?",
                            "JwlMediaWin - Update available",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Information);
                    }

                    if (answer == MessageBoxResult.Yes)
                    {
                        await InstallLatestAsync(latest, ct).ConfigureAwait(false);
                    }
                }
                else
                {
                    Log.Information("[Update] Application is up to date ({Version}).", current);
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "[Update] Failed to check for updates.");
            }
        }

        public async Task CheckAndInstallSilentAsync(CancellationToken ct = default)
        {
            try
            {
                var current = VersionDetection.GetCurrentVersion();
                var latest = await VersionDetection.GetLatestAsync(ct).ConfigureAwait(false);

                if (latest == null)
                {
                    // Keep some UX feedback in silent path
                    if (Application.Current?.Dispatcher != null)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                            MessageBox.Show("Unable to retrieve the latest release information at this time.",
                                "JwlMediaWin",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information));
                    }
                    else
                    {
                        MessageBox.Show("Unable to retrieve the latest release information at this time.",
                            "JwlMediaWin", MessageBoxButton.OK, MessageBoxImage.Information);
                    }

                    return;
                }

                if (latest.Version > current)
                {
                    await InstallLatestAsync(latest, ct).ConfigureAwait(false);
                }
                else
                {
                    if (Application.Current?.Dispatcher != null)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                            MessageBox.Show($"You already have the latest version ({current}).",
                                "JwlMediaWin",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information));
                    }
                    else
                    {
                        MessageBox.Show($"You already have the latest version ({current}).",
                            "JwlMediaWin", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "[Update] Silent update flow failed.");
            }
        }

        private static async Task InstallLatestAsync(VersionDetection.LatestInfo info, CancellationToken ct)
        {
            if (info.AssetUri == null)
            {
                Log.Warning("[Update] No downloadable asset found. Opening Releases page instead.");
                OpenInBrowser(info.HtmlPage.AbsoluteUri);
                return;
            }

            var fileName = string.IsNullOrWhiteSpace(info.AssetName) ? "JwlMediaWinSetup.exe" : info.AssetName;
            var tempPath = Path.Combine(Path.GetTempPath(), fileName);

            try
            {
                // C# 7.3 compatible: no 'using var', no 'await using'
                using (var http = new HttpClient())
                using (var resp = await http.GetAsync(info.AssetUri, HttpCompletionOption.ResponseHeadersRead, ct)
                           .ConfigureAwait(false))
                {
                    resp.EnsureSuccessStatusCode();

                    // ReadAsStreamAsync() without CancellationToken in older frameworks
                    using (var s = await resp.Content.ReadAsStreamAsync().ConfigureAwait(false))
                    using (var fs = File.Create(tempPath))
                    {
                        // Stream.CopyToAsync has a CancellationToken overload we can use
                        await s.CopyToAsync(fs, 81920, ct).ConfigureAwait(false);
                    }
                }

                Log.Information("[Update] Installer downloaded to {Path}", tempPath);

                // Decide EXE vs MSI
                bool isMsi = string.Equals(Path.GetExtension(tempPath), ".msi", StringComparison.OrdinalIgnoreCase);
                var logPath = Path.Combine(Path.GetTempPath(), "JwlMediaWinSetup.log");

                var psi = new ProcessStartInfo
                {
                    UseShellExecute = true, // required for Verb=runas
                    WorkingDirectory = Path.GetDirectoryName(tempPath) ?? Environment.CurrentDirectory,
                    Verb = "runas" // UAC elevation
                };

                if (isMsi)
                {
                    psi.FileName = "msiexec.exe";
                    // Show UI + log
                    psi.Arguments = string.Format("/i \"{0}\" /norestart /l*vx \"{1}\"", tempPath, logPath);
                }
                else
                {
                    psi.FileName = tempPath;
                    // Show UI + log (Inno Setup). No /SILENT so user sees UI.
                    psi.Arguments = string.Format("/NORESTART /LOG=\"{0}\"", logPath);
                }

                // C# 7.3: reference types are nullable by default; don't use 'Process?'
                Process proc = null;

                try
                {
                    // Single start attempt with elevation
                    proc = Process.Start(psi);
                    Log.Information("[Update] Installer process started. PID={Pid}, IsMsi={IsMsi}, Log={Log}",
                        proc != null ? proc.Id : 0, isMsi, logPath);
                }
                catch (System.ComponentModel.Win32Exception w32ex)
                {
                    // 1223 = ERROR_CANCELLED (user canceled the UAC prompt)
                    // 740  = ERROR_ELEVATION_REQUIRED (try without elevation as fallback)
                    Log.Warning(w32ex, "[Update] Failed to start installer with elevation. NativeErrorCode={Code}",
                        w32ex.NativeErrorCode);

                    if (w32ex.NativeErrorCode == 740)
                    {
                        // Try without elevation (may work if the user has perms)
                        var noUac = new ProcessStartInfo
                        {
                            UseShellExecute = true,
                            WorkingDirectory = Path.GetDirectoryName(tempPath) ?? Environment.CurrentDirectory,
                            FileName = isMsi ? "msiexec.exe" : tempPath,
                            Arguments = isMsi
                                ? string.Format("/i \"{0}\" /passive /norestart", tempPath)
                                : "/VERYSILENT /NORESTART"
                            // No Verb => no elevation
                        };

                        try
                        {
                            proc = Process.Start(noUac);
                            Log.Information("[Update] Installer started without elevation. PID={Pid}",
                                proc != null ? proc.Id : 0);
                        }
                        catch (Exception ex2)
                        {
                            Log.Error(ex2, "[Update] Fallback start without elevation failed. Opening Releases page.");
                            OpenInBrowser(info.HtmlPage.AbsoluteUri);
                            return;
                        }
                    }
                    else if (w32ex.NativeErrorCode == 1223)
                    {
                        // User canceled UAC; do not close the app
                        Log.Information("[Update] Installation canceled by user.");
                        return;
                    }
                    else
                    {
                        // Unknown start error -> open releases page as fallback
                        OpenInBrowser(info.HtmlPage.AbsoluteUri);
                        return;
                    }
                }

                // Small delay to let UAC/installer UI front before shutting down this app
                await Task.Delay(1200, ct).ConfigureAwait(false);

                // Close the app only once to free binaries for the installer
                if (Application.Current != null && Application.Current.Dispatcher != null)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        Log.Information("[Update] Shutting down application to allow update.");
                        Application.Current.Shutdown();
                    });
                }
                else
                {
                    Log.Information("[Update] Shutting down application to allow update (no dispatcher).");
                    if (Application.Current != null)
                        Application.Current.Shutdown();
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex,
                    "[Update] Failed to download or execute the installer. Opening Releases page as fallback.");
                OpenInBrowser(info.HtmlPage.AbsoluteUri);
            }
        }

        private static void OpenInBrowser(string url)
        {
            try
            {
                // Primary: rely on the shell to open the default handler (browser/mail/etc.).
                var psi = new ProcessStartInfo
                {
                    FileName = url, // can be http/https/mailto/file
                    UseShellExecute = true,
                    Verb = "open"
                };
                Process.Start(psi);
                return;
            }
            catch (Exception ex1)
            {
                // Fallback: use 'cmd /c start "" "<url>"' which delegates to ShellExecute internally.
                try
                {
                    // Escape double quotes in URL to avoid breaking the cmd line
                    var safeUrl = url.Replace("\"", "\\\"");

                    var psiCmd = new ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        Arguments = "/c start \"\" \"" + safeUrl + "\"",
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden,
                        UseShellExecute = false
                    };
                    Process.Start(psiCmd);
                    return;
                }
                catch (Exception ex2)
                {
                    // Last resort: log both failures
                    Log.Warning(ex1, "[Update] Primary shell open failed.");
                    Log.Warning(ex2, "[Update] Fallback 'cmd /c start' also failed for URL: {Url}", url);
                }
            }
        }
    }
}