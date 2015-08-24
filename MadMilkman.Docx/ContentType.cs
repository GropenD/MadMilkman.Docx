namespace MadMilkman.Docx
{
    /// <summary>
    /// Represents the type of the document's content.
    /// </summary>
    public enum ContentType
    {
        /// <summary>
        /// Defines HTML formatted content.
        /// </summary>
        [ContentInfo(".html", "text/html")]
        Html = 0,

        /// <summary>
        /// Defines RTF formatted content.
        /// </summary>
        [ContentInfo(".rtf", "application/rtf")]
        Rtf
    }
}
