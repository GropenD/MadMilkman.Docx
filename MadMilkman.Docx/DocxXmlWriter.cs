using System;
using System.IO;
using System.Text;
using System.Xml;
using MadMilkman.Docx.Properties;

namespace MadMilkman.Docx
{
    internal sealed class DocxXmlWriter : IDisposable
    {
        private XmlWriter writer;

        public DocxXmlWriter(Stream stream) { this.writer = new XmlTextWriter(stream, new UTF8Encoding(false)); }

        public void WriteStartDocument(string elementName)
        {
            this.writer.WriteStartDocument(true);
            this.writer.WriteStartElement("w", elementName, Resources.MainNamespace);
            this.writer.WriteAttributeString("xmlns", "w", null, Resources.MainNamespace);
            this.writer.WriteAttributeString("xmlns", "r", null, Resources.RelationshipsNamespace);
        }

        public void WriteEndDocument() { this.writer.WriteEndDocument(); }

        public void WriteStartElement(string elementName) { this.writer.WriteStartElement(elementName, Resources.MainNamespace); }
        
        public void WriteElement(string elementName, string relAttributeName, string relAttributeValue)
        {
            this.WriteElement(elementName, relAttributeName, relAttributeValue, null, null);
        }

        public void WriteElement(string elementName, string relAttributeName, string relAttributeValue, string mainAttributeName, string mainAttributeValue)
        {
            this.writer.WriteStartElement(elementName, Resources.MainNamespace);

            if (mainAttributeName != null)
                this.writer.WriteAttributeString(mainAttributeName, Resources.MainNamespace, mainAttributeValue);

            this.writer.WriteAttributeString(relAttributeName, Resources.RelationshipsNamespace, relAttributeValue);
            this.writer.WriteEndElement();
        }

        public void Dispose() { this.writer.Close(); }
    }
}
