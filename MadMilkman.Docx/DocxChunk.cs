namespace MadMilkman.Docx
{
    internal sealed class DocxChunk
    {
        private readonly ContentInfoAttribute contentInfo;

        public string Content { get; private set; }
        public string Extension { get { return this.contentInfo.Extension; } }
        public string Type { get { return this.contentInfo.Type; } }

        public DocxChunk(string content, ContentType contentType)
        {
            this.Content = content;
            this.contentInfo = GetContentInfo(contentType);
        }

        private static ContentInfoAttribute GetContentInfo(ContentType contentType)
        {
            var memberInfo = typeof(ContentType).GetMember(contentType.ToString())[0];
            return (ContentInfoAttribute)memberInfo.GetCustomAttributes(typeof(ContentInfoAttribute), false)[0];
        }
    }
}
