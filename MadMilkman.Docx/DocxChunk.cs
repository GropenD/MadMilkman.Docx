using System;
using MadMilkman.Docx.Properties;

namespace MadMilkman.Docx
{
    internal sealed class DocxChunk
    {
        private DocxContentType docxContentType;

        public string Content { get; private set; }

        public DocxChunk(string content, DocxContentType docxContentType)
        {
            this.Content = content;
            this.docxContentType = docxContentType;
        }

        public string ContentType
        {
            get
            {
                if (this.docxContentType == DocxContentType.Html)
                    return Resources.HtmlContentType;

                if (this.docxContentType == DocxContentType.Rtf)
                    return Resources.RtfContentType;

                else
                    throw new NotSupportedException();
            }
        }

        public string FileExtension
        {
            get
            {
                if (this.docxContentType == DocxContentType.Html)
                    return ".html";

                if (this.docxContentType == DocxContentType.Rtf)
                    return ".rtf";

                else
                    throw new NotSupportedException();
            }
        }
    }
}
