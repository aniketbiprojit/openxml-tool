using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

string baseDir = "/Users/peppercontent/Work/openxml-tool/data/";

string filepath = baseDir + "test-doc.docx";

//int x = 20;
//int y = 60;

using (WordprocessingDocument document =
       WordprocessingDocument.Open(filepath, true))
{
    if (document.MainDocumentPart != null) {
        MainDocumentPart mainDocumentPart = document.MainDocumentPart;
        Console.WriteLine("File read.");
        Console.WriteLine(mainDocumentPart.Document.Descendants().Count());
    }
}
