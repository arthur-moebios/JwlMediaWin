namespace JwlMediaWin
{
    using System;
    using System.Linq;
    using System.Net.Http;
    using System.Reflection;
    using System.Threading;
    using System.Threading.Tasks;
    using Newtonsoft.Json.Linq;
    using Serilog;

    /// <summary>
    /// Gets the local assembly version and the latest release information from GitHub.
    /// </summary>
    internal static class VersionDetection
    {
        private const string ApiLatestRelease =
            "https://api.github.com/repos/AntonyCorbett/JwlMediaWin/releases/latest";

        public static string LatestReleaseHtml => "https://github.com/AntonyCorbett/JwlMediaWin/releases/latest";

        public static Version GetCurrentVersion()
        {
            return Assembly.GetExecutingAssembly().GetName().Version;
        }

        public static string GetCurrentVersionString()
        {
            var v = GetCurrentVersion();
            return string.Format("{0}.{1}.{2}.{3}", v.Major, v.Minor, v.Build, v.Revision);
        }

        public sealed class LatestInfo
        {
            public Version Version { get; set; }
            public Uri AssetUri { get; set; } // direct link to installer (preferred)
            public string AssetName { get; set; }
            public Uri HtmlPage { get; set; } // releases/latest (fallback)
        }

        public static async Task<LatestInfo> GetLatestAsync(CancellationToken ct = default(CancellationToken))
        {
            using (var http = new HttpClient())
            {
                // Reasonable timeout for update checks
                http.Timeout = TimeSpan.FromSeconds(10);

                // GitHub API requires a User-Agent
                http.DefaultRequestHeaders.UserAgent.ParseAdd("JwlMediaWin-Updater");
                http.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
                http.DefaultRequestHeaders.Add("X-GitHub-Api-Version", "2025-08-02");

                var resp = await http.GetAsync(ApiLatestRelease, ct).ConfigureAwait(false);
                resp.EnsureSuccessStatusCode();

                var json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                var jo = JObject.Parse(json);

                var tag = (string)jo["tag_name"] ?? (string)jo["name"] ?? string.Empty;
                var normalized = tag.Trim().TrimStart('v', 'V');

                Version latest;
                if (!Version.TryParse(NormalizeToSystemVersion(normalized), out latest))
                {
                    Log.Warning("[Update] Unable to parse 'tag_name' '{Tag}' into System.Version", tag);
                    return null;
                }

                // Choose an installer-like asset if available
                var assets = jo["assets"] as JArray ?? new JArray();

                JToken chosen = null;
                string[] prefs = { ".exe", ".msi" };
                foreach (var ext in prefs)
                {
                    chosen = assets.FirstOrDefault(a =>
                        ((string)a["browser_download_url"]) != null &&
                        ((string)a["browser_download_url"]).EndsWith(ext, StringComparison.OrdinalIgnoreCase));
                    if (chosen != null)
                        break;
                }

                // Fallback if nothing found
                if (chosen == null && assets.Count > 0)
                    chosen = assets.First();

                var assetUrl = chosen != null ? (string)chosen["browser_download_url"] : null;
                var assetName = chosen != null ? (string)chosen["name"] : null;

                Log.Information("[Update] Latest remote version: {Remote} | Asset: {Asset}",
                    latest, string.IsNullOrEmpty(assetName) ? "(none)" : assetName);

                return new LatestInfo
                {
                    Version = latest,
                    AssetUri = string.IsNullOrWhiteSpace(assetUrl) ? null : new Uri(assetUrl),
                    AssetName = assetName,
                    HtmlPage = new Uri(LatestReleaseHtml)
                };
            }
        }

        /// <summary>
        /// Normalizes SemVer-ish strings (e.g., "1.2.3", "1.2.3-beta") to something System.Version can parse.
        /// </summary>
        private static string NormalizeToSystemVersion(string v)
        {
            // Remove pre-release or build metadata suffixes like "-beta", "+build.5"
            int dashIdx = v.IndexOf('-');
            int plusIdx = v.IndexOf('+');
            int cutIdx = -1;

            if (dashIdx >= 0 && plusIdx >= 0)
                cutIdx = Math.Min(dashIdx, plusIdx);
            else if (dashIdx >= 0)
                cutIdx = dashIdx;
            else if (plusIdx >= 0)
                cutIdx = plusIdx;

            if (cutIdx >= 0)
                v = v.Substring(0, cutIdx);

            var parts = v.Split('.');
            if (parts.Length == 2)
                return v + ".0";
            if (parts.Length == 3)
                return v + ".0"; // "1.2.3" -> "1.2.3.0"
            return v; // already 4 fields or invalid (TryParse will decide)
        }
    }
}