using System.Diagnostics;
using NUnit.Framework;

namespace MadMilkman.Docx.Tests
{
    [TestFixture]
    public class RegressionTests
    {
        [Test]
        public void DocxRtfResultTest()
        {
            var docx = new DocxFile();

            docx.Body.AppendFile(@"..\..\TestFiles\Font.rtf", DocxContentType.Rtf);
            docx.Body.AppendFile(@"..\..\TestFiles\Image.rtf", DocxContentType.Rtf);
            docx.Body.AppendFile(@"..\..\TestFiles\Table.rtf", DocxContentType.Rtf);

            docx.Save("DocxRtfResultTest.docx");
            Process.Start("DocxRtfResultTest.docx");
        }

        [Test]
        public void DocxHtmlResultTest()
        {
            var docx = new DocxFile();

            docx.Body.AppendFile(@"..\..\TestFiles\Font.html", DocxContentType.Html);
            docx.Body.AppendFile(@"..\..\TestFiles\Image.html", DocxContentType.Html);
            docx.Body.AppendFile(@"..\..\TestFiles\Table.html", DocxContentType.Html);

            docx.Save("DocxHtmlResultTest.docx");
            Process.Start("DocxHtmlResultTest.docx");
        }
    }
}
