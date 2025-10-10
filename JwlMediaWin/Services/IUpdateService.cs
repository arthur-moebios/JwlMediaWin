namespace JwlMediaWin.Services
{
    using System.Threading;
    using System.Threading.Tasks;

    internal interface IUpdateService
    {
        /// <summary>
        /// Checks for updates and prompts the user if a newer version is available.
        /// </summary>
        Task CheckAndOfferAsync(CancellationToken ct = default);

        /// <summary>
        /// Checks for updates and, if available, downloads and runs the installer silently.
        /// Shows a small info message if already up-to-date.
        /// </summary>
        Task CheckAndInstallSilentAsync(CancellationToken ct = default);
    }
}