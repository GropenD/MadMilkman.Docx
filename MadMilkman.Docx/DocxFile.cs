using System;
using System.Diagnostics;
using System.IO;

namespace MadMilkman.Docx
{
    /// <summary>
    /// <para>In-memory representation of a DOCX file.</para>
    /// <para>Use <see cref="Body"/>, <see cref="Header"/> and <see cref="Footer"/> properties to fill out the content of the required document's part.</para>
    /// </summary>
    public sealed class DocxFile
    {
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private DocxContentBuilder body;
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private DocxContentBuilder header;
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private DocxContentBuilder footer;

        /// <summary>
        /// Initializes a new instance of <see cref="DocxFile"/> class.
        /// </summary>
        public DocxFile() { this.body = new DocxContentBuilder(); }

        /// <summary>
        /// Gets document's main content builder.
        /// </summary>
        public DocxContentBuilder Body { get { return this.body; } }

        /// <summary>
        /// Gets document's header content builder.
        /// </summary>
        public DocxContentBuilder Header
        {
            get
            {
                if (this.header == null)
                    this.header = new DocxContentBuilder();
                return this.header;
            }
        }

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        internal bool HasHeader { get { return this.header != null && this.header.Count != 0; } }

        /// <summary>
        /// Gets document's footer content builder.
        /// </summary>
        public DocxContentBuilder Footer
        {
            get
            {
                if (this.footer == null)
                    this.footer = new DocxContentBuilder();
                return this.footer;
            }
        }

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        internal bool HasFooter { get { return this.footer != null && this.footer.Count != 0; } }

        /// <summary>
        /// Saves a file to a path.
        /// </summary>
        /// <param name="filePath">Path to which to save a file.</param>
        public void Save(string filePath)
        {
            if (filePath == null)
                throw new ArgumentNullException("filePath");

            using (Stream fileStream = File.Create(filePath))
                this.Save(fileStream);
        }

        /// <summary>
        /// Saves a file to a stream.
        /// </summary>
        /// <param name="fileStream">Stream to which to save a file.</param>
        public void Save(Stream fileStream)
        {
            if (fileStream == null)
                throw new ArgumentNullException("fileStream");

            using (var writer = new DocxPackageWriter(fileStream))
                writer.Write(this);

            if (fileStream.CanSeek)
                fileStream.Seek(0, SeekOrigin.Begin);
        }
    }
}
