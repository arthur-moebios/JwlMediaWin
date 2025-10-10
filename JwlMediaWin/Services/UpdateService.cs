namespace JwlMediaWin.Services
{
    using System;
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

                    var res = MessageBox.Show(
                        $"A new version ({latest.Version}) is available. You have {current}.\n\nInstall now?",
                        "JwlMediaWin - Update available",
                        MessageBoxButton.YesNo, MessageBoxImage.Information);

                    if (res == MessageBoxResult.Yes)
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
                    MessageBox.Show("Unable to retrieve the latest release information at this time.",
                        "JwlMediaWin", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (latest.Version > current)
                {
                    await InstallLatestAsync(latest, ct).ConfigureAwait(false);
                }
                else
                {
                    MessageBox.Show($"You already have the latest version ({current}).",
                        "JwlMediaWin", MessageBoxButton.OK, MessageBoxImage.Information);
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
                Process.Start(info.HtmlPage.AbsoluteUri);
                return;
            }

            var fileName = string.IsNullOrWhiteSpace(info.AssetName) ? "JwlMediaWinSetup.exe" : info.AssetName;
            var tempPath = Path.Combine(Path.GetTempPath(), fileName);

            try
            {
                using (var http = new HttpClient())
                using (var s = await http.GetStreamAsync(info.AssetUri).ConfigureAwait(false))
                using (var fs = File.Create(tempPath))
                {
                    await s.CopyToAsync(fs, 81920, ct).ConfigureAwait(false);
                }

                Log.Information("[Update] Installer downloaded to {Path}", tempPath);

                // Decide EXE vs MSI
                bool isMsi = string.Equals(Path.GetExtension(tempPath), ".msi", StringComparison.OrdinalIgnoreCase);

                var logPath = Path.Combine(Path.GetTempPath(), "JwlMediaWinSetup.log");

                var psi = new ProcessStartInfo
                {
                    UseShellExecute  = true,
                    WorkingDirectory = Path.GetDirectoryName(tempPath) ?? Environment.CurrentDirectory,
                    Verb             = "runas" // UAC
                };

                if (isMsi)
                {
                    psi.FileName  = "msiexec.exe";
                    // MOSTRAR UI + LOG
                    psi.Arguments = string.Format("/i \"{0}\" /norestart /l*vx \"{1}\"", tempPath, logPath);
                }
                else
                {
                    psi.FileName  = tempPath;
                    // MOSTRAR UI + LOG (Inno Setup): /LOG grava em arquivo, sem /SILENT
                    psi.Arguments = string.Format("/NORESTART /LOG=\"{0}\"", logPath);
                    // Se quiser ainda mais verboso: remova /NORESTART para ver prompts de reboot, se houver.
                }

                Process proc = Process.Start(psi);
                Log.Information("[Update] Installer process started. PID={Pid}, Log={Log}", proc != null ? proc.Id : 0, logPath);

                Application.Current.Dispatcher.Invoke(() =>
                {
                    Log.Information("[Update] Shutting down application to allow update.");
                    Application.Current.Shutdown();
                });


                try
                {
                    proc = Process.Start(psi);
                    Log.Information("[Update] Installer process started. PID={Pid}, IsMsi={IsMsi}",
                        proc != null ? proc.Id : 0, isMsi);
                }
                catch (System.ComponentModel.Win32Exception w32ex)
                {
                    // 1223 = ERROR_CANCELLED (user canceled the UAC prompt)
                    // 740  = ERROR_ELEVATION_REQUIRED (try without elevation as fallback)
                    Log.Warning(w32ex, "[Update] Failed to start installer with elevation. NativeErrorCode={Code}",
                        w32ex.NativeErrorCode);

                    if (w32ex.NativeErrorCode == 740)
                    {
                        // Try without elevation (may still work if user has perms)
                        var noUac = new ProcessStartInfo
                        {
                            UseShellExecute = true,
                            WorkingDirectory = Path.GetDirectoryName(tempPath) ?? Environment.CurrentDirectory,
                            FileName = isMsi ? "msiexec.exe" : tempPath,
                            Arguments = isMsi
                                ? string.Format("/i \"{0}\" /passive /norestart", tempPath)
                                : "/VERYSILENT /NORESTART"
                            // Verb omitted → no elevation
                        };

                        try
                        {
                            proc = Process.Start(noUac);
                            Log.Information("[Update] Installer started without elevation. PID={Pid}",
                                proc != null ? proc.Id : 0);
                        }
                        catch (Exception ex2)
                        {
                            Log.Error(ex2, "[Update] Fallback start without elevation failed.");
                            Process.Start(info.HtmlPage.AbsoluteUri);
                            return;
                        }
                    }
                    else if (w32ex.NativeErrorCode == 1223)
                    {
                        // User canceled UAC; don't shutdown app
                        Log.Information("[Update] Installation canceled by user.");
                        return;
                    }
                    else
                    {
                        Process.Start(info.HtmlPage.AbsoluteUri);
                        return;
                    }
                }

                // Pequeno atraso para garantir que o Shell/UAC apareça antes de fechar o app
                await Task.Delay(1200).ConfigureAwait(false);

                // Fecha o app para liberar arquivos para o instalador
                Application.Current.Dispatcher.Invoke(() =>
                {
                    Log.Information("[Update] Shutting down application to allow update.");
                    Application.Current.Shutdown();
                });
            }
            catch (Exception ex)
            {
                Log.Error(ex,
                    "[Update] Failed to download or execute the installer. Opening Releases page as fallback.");
                Process.Start(info.HtmlPage.AbsoluteUri);
            }
        }
    }
}