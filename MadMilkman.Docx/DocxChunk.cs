using System;
using MadMilkman.Docx.Properties;
using ChunkContentType = MadMilkman.Docx.ContentType;

namespace MadMilkman.Docx
{
    internal sealed class DocxChunk
    {
        private ChunkContentType contentType;

        public string Content { get; private set; }

        public DocxChunk(string content, ChunkContentType contentType)
        {
            this.Content = content;
            this.contentType = contentType;
        }

        public string ContentType
        {
            get
            {
                if (this.contentType == ChunkContentType.Html)
                    return Resources.HtmlContentType;

                if (this.contentType == ChunkContentType.Rtf)
                    return Resources.RtfContentType;

                else
                    throw new NotSupportedException();
            }
        }

        public string FileExtension
        {
            get
            {
                if (this.contentType == ChunkContentType.Html)
                    return ".html";

                if (this.contentType == ChunkContentType.Rtf)
                    return ".rtf";

                else
                    throw new NotSupportedException();
            }
        }
    }
}
