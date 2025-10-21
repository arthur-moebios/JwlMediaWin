namespace JwlMediaWin.Core.Models
{
    public class FixerStatus
    {
        /// <summary>
        /// Gets or sets a value indicating whether the core window is focused.
        /// </summary>
        /// <value>
        ///   <c>true</c> if the core window is focused; otherwise, <c>false</c>.
        /// </value>
        public bool CoreWindowFocused { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this the media window is "fixed".
        /// </summary>
        /// <value>
        ///   <c>true</c> if the media window is fixed; otherwise, <c>false</c>.
        /// </value>
        public bool IsFixed { get; set; }

        /// <summary>
        /// Gets or sets the "find window" result.
        /// </summary>
        /// <value>
        /// The "find window" result.
        /// </value>
        public FindWindowResult FindWindowResult { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether teh media window was transitioning.
        /// </summary>
        /// <value>
        ///   <c>true</c> if the media window was transitioning; otherwise, <c>false</c>.
        /// </value>
        public bool ErrorIsTransitioning { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether there was an unknown error.
        /// </summary>
        /// <value>
        ///   <c>true</c> if there was an unknown error unknown; otherwise, <c>false</c>.
        /// </value>
        public bool ErrorUnknown { get; set; }
        
        /// <summary>True se Core/Media window exibe algum conteúdo (imagem ou vídeo).</summary>
        public bool MediaHasContent { get; set; }

        /// <summary>"Video", "Image", "None" ou "Unknown".</summary>
        public string MediaContentType { get; set; }

        /// <summary>True se o estado de conteúdo mudou desde o ciclo anterior.</summary>
        public bool MediaContentChanged { get; set; }
    }
}
