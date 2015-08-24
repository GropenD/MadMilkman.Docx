using System;

namespace MadMilkman.Docx
{
    [AttributeUsage(AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    internal sealed class ContentInfoAttribute : Attribute
    {
        public string Extension { get; private set; }
        public string Type { get; private set; }

        public ContentInfoAttribute(string extension, string type)
        {
            this.Extension = extension;
            this.Type = type;
        }
    }
}
