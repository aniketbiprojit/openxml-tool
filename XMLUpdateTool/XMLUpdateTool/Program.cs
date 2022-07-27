using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

string baseDir = "/Users/peppercontent/Work/openxml-tool/data/";

string filepath = baseDir + "test-doc-no-comment.docx";

int x = 7;
int y = 59;

string textToSearch = "overview of portal_backend repository code.\nGet clari";
textToSearch = Regex.Replace(textToSearch, @"\t|\n|\r", "");


string textFromDoc = "";

Console.WriteLine(x);
Console.WriteLine(y);
Console.WriteLine(textToSearch);

Run? forX = null;
Run? forY = null;

using (WordprocessingDocument document =
       WordprocessingDocument.Open(filepath, true))
{
    if (document.MainDocumentPart != null) {
        MainDocumentPart mainDocumentPart = document.MainDocumentPart;
        Console.WriteLine("File read.");
        Console.WriteLine(mainDocumentPart.Document.Descendants().Count());
        foreach(Run data in mainDocumentPart.Document.Descendants<Run>())
        {
            Console.WriteLine(data.InnerText);
            textFromDoc += data.InnerText;
            if (textFromDoc.Length - 1 >= x && forX==null) {
                forX = data;
                Console.WriteLine("x = " + x + " " + textFromDoc[x]);
            }
            if (textFromDoc.Length - 1 >= y && forY == null)
            {
                forY = data;
                Console.WriteLine("y = " + y + " " + textFromDoc[y]);
            }
        }
    }
    if(forX!=null && forY != null)
    {
        if (textFromDoc[x..y] == textToSearch) {
            Console.WriteLine("Matched");
        }
        else
        {
            Console.WriteLine(textFromDoc[x..y]);
            Console.WriteLine(textToSearch);
            Console.WriteLine(textFromDoc);
            throw new Exception("text does not match");
        }
    }
    else
    {
        throw new Exception("forX or forY null");
    }

    
}
