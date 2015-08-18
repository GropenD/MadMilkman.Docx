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

            docx.Body.AppendFile(@"..\..\TestFiles\Font.rtf", ContentType.Rtf)
                     .AppendFile(@"..\..\TestFiles\Image.rtf", ContentType.Rtf)
                     .AppendFile(@"..\..\TestFiles\Table.rtf", ContentType.Rtf);

            docx.Save("DocxRtfResultTest.docx");
            Process.Start("DocxRtfResultTest.docx");
        }

        [Test]
        public void DocxHtmlResultTest()
        {
            var docx = new DocxFile();

            docx.Body.AppendFile(@"..\..\TestFiles\Font.html", ContentType.Html)
                     .AppendFile(@"..\..\TestFiles\Image.html", ContentType.Html)
                     .AppendFile(@"..\..\TestFiles\Table.html", ContentType.Html);

            docx.Save("DocxHtmlResultTest.docx");
            Process.Start("DocxHtmlResultTest.docx");
        }
    }
}
