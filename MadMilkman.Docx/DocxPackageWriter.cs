using System;
using System.IO;
using System.IO.Packaging;
using MadMilkman.Docx.Properties;

namespace MadMilkman.Docx
{
    internal sealed class DocxPackageWriter : IDisposable
    {
        private const string AltChunkPartialId = "altChunk";
        private const string AltChunkPartialPath = "/word/altChunk";

        private const string DocumentPath = "/word/document.xml";
        private const string HeaderPath = "/word/header1.xml";
        private const string FooterPath = "/word/footer1.xml";

        private Package package;
        private PackagePart documentPart;

        private int altChunkCounter;

        public DocxPackageWriter(Stream stream)
        {
            this.package = Package.Open(stream, FileMode.Create);

            Uri temp = new Uri("/temp.xml", UriKind.Relative);
            this.package.CreatePart(temp, "application/xml");
            this.package.DeletePart(temp);
        }

        public void Write(DocxFile docx)
        {
            this.documentPart = this.CreatePart(
                new Uri(DocxPackageWriter.DocumentPath, UriKind.Relative),
                Resources.DocumentContentType);

            using (DocxXmlWriter writer = new DocxXmlWriter(this.documentPart.GetStream()))
            {
                writer.WriteStartDocument("document");
                writer.WriteStartElement("body");

                this.WriteAltChunks(docx.Body, writer, this.documentPart);

                if (docx.HasHeader || docx.HasHeader)
                    this.WriteHeaderAndFooter(docx, writer);
                
                writer.WriteEndDocument();
            }

            this.package.CreateRelationship(
                new Uri(DocxPackageWriter.DocumentPath.Substring(1), UriKind.Relative),
                TargetMode.Internal,
                Resources.DocumentRelationshipType, "rId1");
        }

        private void WriteHeaderAndFooter(DocxFile docx, DocxXmlWriter writer)
        {
            writer.WriteStartElement("sectPr");
            int idCounter = 1;

            if (docx.HasHeader)
            {
                string id = "rId" + idCounter;
                writer.WriteElement("headerReference", "id", id, "type", "default");

                this.CreateHeaderOrFooterPart(
                    DocxPackageWriter.HeaderPath,
                    Resources.HeaderContentType,
                    Resources.HeaderRelationshipType,
                    id, "hdr", docx.Header);

                idCounter++;
            }

            if (docx.HasFooter)
            {
                string id = "rId" + idCounter;
                writer.WriteElement("footerReference", "id", id, "type", "default");

                this.CreateHeaderOrFooterPart(
                    DocxPackageWriter.FooterPath,
                    Resources.FooterContentType,
                    Resources.FooterRelationshipType,
                    id, "ftr", docx.Footer);
            }
        }

        private void CreateHeaderOrFooterPart(string partPath, string partContentType, string partRelationshipType, string id, string elementName, DocxContentBuilder chunks)
        {
            Uri partUri = new Uri(partPath, UriKind.Relative);
            PackagePart part = this.CreatePart(partUri, partContentType);

            using (DocxXmlWriter writer = new DocxXmlWriter(part.GetStream()))
            {
                writer.WriteStartDocument(elementName);
                this.WriteAltChunks(chunks, writer, part);
                writer.WriteEndDocument();
            }

            CreatePartRelationship(partUri, partRelationshipType, id, this.documentPart);
        }

        private void WriteAltChunks(DocxContentBuilder chunks, DocxXmlWriter writer, PackagePart parentPart)
        {
            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++)
            {
                DocxChunk chunk = chunks[chunkIndex];
                this.altChunkCounter++;

                string altChunkId = DocxPackageWriter.AltChunkPartialId + altChunkCounter;
                writer.WriteElement("altChunk", "id", altChunkId);

                string altChunkPath = DocxPackageWriter.AltChunkPartialPath + altChunkCounter + chunk.Extension;
                this.CreateAltChunkPart(chunk, altChunkPath, altChunkId, parentPart);
            }
        }

        private void CreateAltChunkPart(DocxChunk chunk, string altChunkPath, string altChunkId, PackagePart parentPart)
        {
            Uri altChunkUri = new Uri(altChunkPath, UriKind.Relative);
            PackagePart altChunkPart = this.CreatePart(altChunkUri, chunk.Type);

            using (var writer = new StreamWriter(altChunkPart.GetStream()))
                writer.Write(chunk.Content);

            CreatePartRelationship(altChunkUri, Resources.AltChunkRelationshipType, altChunkId, parentPart);
        }
        
        private PackagePart CreatePart(Uri partUri, string partContentType)
        {
            return this.package.CreatePart(partUri, partContentType, CompressionOption.Normal);
        }

        private static void CreatePartRelationship(Uri partUri, string partRelationshipType, string partId, PackagePart parentPart)
        {
            Uri relativeUri = PackUriHelper.GetRelativeUri(parentPart.Uri, partUri);
            parentPart.CreateRelationship(relativeUri, TargetMode.Internal, partRelationshipType, partId);
        }

        public void Dispose() { this.package.Close(); }
    }
}
