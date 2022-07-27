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

        if (forX != null && forY != null)
        {
            if (textFromDoc[x..y] == textToSearch)
            {
                Console.WriteLine("Matched");
                Comments comments = null;

                string initialCommentId = "0";

                if (mainDocumentPart.GetPartsCountOfType<WordprocessingCommentsPart>() > 0)
                {
                    comments = mainDocumentPart.WordprocessingCommentsPart.Comments;
                    if (comments.HasChildren)
                    {
                        // Obtain an unused ID.
                        initialCommentId = (comments.Descendants<Comment>().Select(e => int.Parse(e.Id.Value)).Max() + 1).ToString();
                    }
                }
                else
                {
                    // No WordprocessingCommentsPart part exists, so add one to the package.
                    WordprocessingCommentsPart commentPart = mainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                    commentPart.Comments = new Comments();
                    comments = commentPart.Comments;
                }


                Paragraph p = new Paragraph(new Run(new Text("Added Comment 1")));
                Comment cmt =
                    new Comment()
                    {
                        Id = initialCommentId,
                        Author = "Aniket Biprojit Chowdhury",
                        Initials = "AC",
                        Date = DateTime.Now
                    };
                cmt.AppendChild(p);
                comments.AppendChild(cmt);
                comments.Save();

                // add a commentRangeStart before forX
                forX.InsertBeforeSelf(new CommentRangeStart(){ Id = initialCommentId });

                // add a commentRangeEnd after forY
                var cmtEnd = forY.InsertAfterSelf(new CommentRangeEnd() { Id = initialCommentId });


                // add a commentReference after commentRangeEnd
                cmtEnd.InsertAfterSelf(new Run(new CommentReference() { Id = initialCommentId }));

                document.Save();
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
}
