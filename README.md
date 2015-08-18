# ![MadMilkman.Docx component's icon](../master/MadMilkman.Docx/Properties/MadMilkman.Docx.png) MadMilkman.Docx
**MadMilkman.Docx** is a .NET library which can enable you to create a new DOCX file by appending HTML and RTF content to the DOCX document's body, header and footer.

```csharp
// Create new file.
DocxFile document = new DocxFile();

// Add body content.
document.Body.AppendFile("Body.html", ContentType.Html)
             .AppendFile("Body.rtf", ContentType.Rtf);

// Add header content.
document.Header.AppendFile("Header.html", ContentType.Html)
               .AppendText("<html><body><p>Sample header content!</p></html></body>", ContentType.Html);

// Add footer content.
document.Footer.AppendFile("Footer.Rtf", ContentType.Rtf)
               .AppendText(@"{\rtf1\ansi\cf1 Sample footer content!}", ContentType.Rtf);

// Save file.
document.Save("Sample.docx");
```

## Overview:
The component uses an _alternative format import part_ mechanism which enables importing of external content, in an alternate format, directly into a DOCX file. In other words it achieves combining of multiple documents into a single Word document.

The component currently supports importing of only HTML and RTF content, however the following is a list of additional formats that are supported in importing functionality:
* TXT
* XHTML
* MHTML
* XML
* DOCX
* DOCM
* DOTX
* DOTM

## Feedback & Support:
Feel free to contact me with any questions or issues regarding the MadMilkman.Docx component. Also if you require a support for any additional format(s) or you have some other suggestion for component's improvement, don't hesitate to ask!

## Installation:
You can download the library from [here](https://github.com/MarioZ/MadMilkman.Docx/raw/master/MadMilkman.Docx.zip) or from [NuGet](http://www.nuget.org/packages/MadMilkman.Docx).

## Limitations
The _alternative format import part_ technique does not convert the imported content from an alternative format into a WordprocessingML format but rather relies on the consuming application to merge the documents. However by DOCX specification consuming application is free to ignore the imported external content. In other words some Word processing applications (like MS Word) are able to render the external content(s) while others (like Open Office Writer) will ignore them and render a blank document instead.
