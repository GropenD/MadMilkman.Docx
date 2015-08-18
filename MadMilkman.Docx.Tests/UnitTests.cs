using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using NUnit.Framework;

namespace MadMilkman.Docx.Tests
{
    [TestFixture]
    public class UnitTests
    {
#if DEBUG
        [Test]
        public void ContentBuilderTest()
        {
            var docx = new DocxFile();
            AppendBody(docx);

            Assert.AreEqual(2, docx.Body.Count);
            Assert.AreEqual("text/html", docx.Body[0].ContentType);
            Assert.AreEqual(".html", docx.Body[0].FileExtension);
            Assert.AreEqual("application/rtf", docx.Body[1].ContentType);
            Assert.AreEqual(".rtf", docx.Body[1].FileExtension);

            Assert.IsFalse(docx.HasHeader);
            Assert.IsFalse(docx.HasFooter);

            AppendHeaderFooter(docx);

            Assert.AreEqual(2, docx.Header.Count);
            Assert.AreEqual("text/html", docx.Header[0].ContentType);
            Assert.AreEqual(".html", docx.Header[0].FileExtension);
            Assert.AreEqual("application/rtf", docx.Header[1].ContentType);
            Assert.AreEqual(".rtf", docx.Header[1].FileExtension);

            Assert.AreEqual(2, docx.Footer.Count);
            Assert.AreEqual("text/html", docx.Footer[0].ContentType);
            Assert.AreEqual(".html", docx.Footer[0].FileExtension);
            Assert.AreEqual("application/rtf", docx.Footer[1].ContentType);
            Assert.AreEqual(".rtf", docx.Footer[1].FileExtension);
        }
#endif

        [Test]
        public void PackagePartTest()
        {
            var docx = new DocxFile();
            AppendBody(docx);
            AppendHeaderFooter(docx);

            var stream = new MemoryStream();
            docx.Save(stream);
            var package = Package.Open(stream);

            var parts = package.GetParts();
            HashSet<string> expectedParts =
                new HashSet<string>(
                    new string[] { "/word/altChunk1.html", "/word/altChunk2.rtf", "/word/altChunk3.html", "/word/altChunk4.rtf", "/word/altChunk5.html", "/word/altChunk6.rtf", "/word/document.xml", "/word/footer1.xml", "/word/header1.xml", "/word/_rels/document.xml.rels", "/word/_rels/footer1.xml.rels", "/word/_rels/header1.xml.rels", "/_rels/.rels" });

            foreach (PackagePart part in parts)
                Assert.IsTrue(expectedParts.Remove(part.Uri.OriginalString));
            Assert.IsTrue(expectedParts.Count == 0);


            package.Close();
        }

        [Test]
        public void PackageRelationshipTest()
        {
            var docx = new DocxFile();
            AppendBody(docx);
            AppendHeaderFooter(docx);

            var stream = new MemoryStream();
            docx.Save(stream);
            var package = Package.Open(stream);

            CheckRelationships(
                new string[] { "altChunk1.html", "altChunk2.rtf", "header1.xml", "footer1.xml" },
                package.GetPart(new Uri("/word/document.xml", UriKind.Relative)).GetRelationships());

            CheckRelationships(
                new string[] { "altChunk3.html", "altChunk4.rtf" },
                package.GetPart(new Uri("/word/header1.xml", UriKind.Relative)).GetRelationships());

            CheckRelationships(
                new string[] { "altChunk5.html", "altChunk6.rtf" },
                package.GetPart(new Uri("/word/footer1.xml", UriKind.Relative)).GetRelationships());

            package.Close();
        }

        private static void CheckRelationships(string[] expected, PackageRelationshipCollection relationships)
        {
            HashSet<string> expectedRelationships = new HashSet<string>(expected);

            foreach (PackageRelationship relationship in relationships)
                Assert.IsTrue(expectedRelationships.Remove(relationship.TargetUri.OriginalString));
            Assert.IsTrue(expectedRelationships.Count == 0);
        }

        private static void AppendBody(DocxFile docx)
        {
            docx.Body.AppendText("<html><body><span style='color:blue;'>Body HTML content!</span></body></html>", ContentType.Html);
            docx.Body.AppendText(@"{\rtf1\ansi\deff0{\colortbl;\red255\green0\blue0;}\cf1 Body RTF content!}", ContentType.Rtf);
        }

        private static void AppendHeaderFooter(DocxFile docx)
        {
            docx.Header.AppendText("<html><body><span style='color:blue;'>Header HTML content!</span></body></html>", ContentType.Html);
            docx.Header.AppendText(@"{\rtf1\ansi\deff0{\colortbl;\red255\green0\blue0;}\cf1 Header RTF content!}", ContentType.Rtf);

            docx.Footer.AppendText("<html><body><span style='color:blue;'>Footer HTML content!</span></body></html>", ContentType.Html);
            docx.Footer.AppendText(@"{\rtf1\ansi\deff0{\colortbl;\red255\green0\blue0;}\cf1 Footer RTF content!}", ContentType.Rtf);
        }
    }
}
