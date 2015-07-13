using System;
using System.Collections.Generic;
using System.IO;

namespace MadMilkman.Docx
{
    /// <summary>
    /// Builds the content for a part of a document.
    /// </summary>
    public sealed class DocxContentBuilder
    {
        private List<DocxChunk> chunks;

        internal DocxContentBuilder() { this.chunks = new List<DocxChunk>(); }

        /// <summary>
        /// Adds text content to document.
        /// </summary>
        /// <param name="content">Text content to add to document.</param>
        /// <param name="contentType">Type of the text content.</param>
        /// <returns>A reference to this instance after the append operation has completed.</returns>
        public DocxContentBuilder AppendText(string content, DocxContentType contentType)
        {
            if (content == null)
                throw new ArgumentNullException("content");

            this.chunks.Add(new DocxChunk(content, contentType));
            return this;
        }

        /// <summary>
        /// Adds file's content to document.
        /// </summary>
        /// <param name="filePath">Path from which to read a file.</param>
        /// <param name="contentType">Type of the file's content.</param>
        /// <returns>A reference to this instance after the append operation has completed.</returns>
        public DocxContentBuilder AppendFile(string filePath, DocxContentType contentType)
        {
            if (filePath == null)
                throw new ArgumentNullException("filePath");

            using (Stream fileStream = File.OpenRead(filePath))
                return this.AppendFile(fileStream, contentType);
        }

        /// <summary>
        /// Adds file's content to document.
        /// </summary>
        /// <param name="fileStream">Stream with which to read a file.</param>
        /// <param name="contentType">Type of the file's content.</param>
        /// <returns>A reference to this instance after the append operation has completed.</returns>
        public DocxContentBuilder AppendFile(Stream fileStream, DocxContentType contentType)
        {
            if (fileStream == null)
                throw new ArgumentNullException("fileStream");

            using (StreamReader reader = new StreamReader(fileStream))
                return this.AppendText(reader.ReadToEnd(), contentType);
        }

        /// <summary>
        /// Removes all content from the current <see cref="DocxContentBuilder"/> instance.
        /// </summary>
        public void Clear() { this.chunks.Clear(); }

        internal int Count { get { return this.chunks.Count; } }

        internal DocxChunk this[int index] { get { return this.chunks[index]; } }
    }
}
