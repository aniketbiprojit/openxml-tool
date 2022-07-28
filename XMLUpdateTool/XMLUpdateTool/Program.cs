using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;

string baseDir = "/Users/peppercontent/Work/openxml-tool/data/";

string filepath = baseDir + "document.docx";

int x = 140;
int y = 244;


string textToSearch = "dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make ";
textToSearch = Regex.Replace(textToSearch, @"\t|\n|\r", "");


string textFromDoc = "";

Console.WriteLine(x);
Console.WriteLine(y);
Console.WriteLine(textToSearch);

Run? forX = null;
Run? forY = null;

int lengthBeforeY = 0;
int lengthBeforeX = 0;

bool addComment = false;

using (WordprocessingDocument document =
       WordprocessingDocument.Open(filepath, true))
{
    if (document.MainDocumentPart != null) {
        MainDocumentPart mainDocumentPart = document.MainDocumentPart;
        
        foreach(Run data in mainDocumentPart.Document.Descendants<Run>())
        {
          
            textFromDoc += data.InnerText;
            
            if (textFromDoc.Length - 1 >= x && forX == null) {
               forX = data;
                Console.WriteLine(lengthBeforeY + ", x = " + x + " " + textFromDoc[x]);
            }

            if (textFromDoc.Length - 1 >= y && forY == null)
            {
                forY = data;
                Console.WriteLine(lengthBeforeY + ", y = " + y + " " + data.InnerText[..(y-lengthBeforeY)]);
            }
            if (forX == null)
            {
                lengthBeforeX += data.InnerText.Length;
                Console.WriteLine(lengthBeforeX);
            }

            if (forY==null)
            {
                lengthBeforeY += data.InnerText.Length;
                Console.WriteLine(lengthBeforeY);
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
                // split forX into two
                // handle forX==forY
                if (forX == forY)
                {
                    Console.WriteLine("Equal");
                    // split forX or forY into three elements.
                    // [..lengthBeforeX], [lengthBeforeX..lengthBeforeY], [lengthBeforeY..]
                    string text1 = forX.InnerText[..(x - lengthBeforeX)]; // text1
                    string text2 = textFromDoc[x..y];                     // text2
                    string text3 = forY.InnerText[(y - lengthBeforeY)..]; // text3

                    Run addedBefore = (Run)forY.Clone();
                    Text addedRunText = addedBefore.GetFirstChild<Text>();
                    //addedRunText.Space = SpaceProcessingModeValues.Preserve;
                    addedRunText.Text = text1;

                    Run addedAfter = (Run)forY.Clone();
                    Text addedAfterText = addedAfter.GetFirstChild<Text>();
                    //addedRunText.Space = SpaceProcessingModeValues.Preserve;
                    addedAfterText.Text = text3;

                    Text forYText = forY.GetFirstChild<Text>();
                    forYText.Text = text2;

                    forY.InsertBeforeSelf(addedBefore);
                    forY.InsertAfterSelf(addedAfter);

                    forY.InsertBeforeSelf(new CommentRangeStart() { Id = initialCommentId });
                    var cmtEnd = forY.InsertAfterSelf(new CommentRangeEnd() { Id = initialCommentId });

                    // add a commentReference after commentRangeEnd
                    cmtEnd.InsertAfterSelf(new Run(new CommentReference() { Id = initialCommentId }));

                    document.Save();

                    Console.WriteLine("Finished, Equal");
                }
                else
                {
                    // split forX into before([0..lengthBeforeX], [lengthBeforeX..X]) and forX
                    forX.InsertBeforeSelf(new CommentRangeStart() { Id = initialCommentId });

                    // add a commentRangeEnd after forY
                    Console.WriteLine("forY: " + forY.InnerText);

                    // split forY into two runs
                    string text1 = forY.InnerText[..(y - lengthBeforeY)];
                    string text2 = forY.InnerText[(y - lengthBeforeY)..];

                    Console.WriteLine(text1 + " 1");
                    Console.WriteLine(text2 + " 2");

                    Text forYText = forY.GetFirstChild<Text>();
                    forYText.Text = text1;

                    Run addedRun = (Run)forY.Clone();
                    Text addedRunText = addedRun.GetFirstChild<Text>();
                    //addedRunText.Space = SpaceProcessingModeValues.Preserve;
                    addedRunText.Text = text2;

                    forY.InsertAfterSelf(addedRun);


                    var cmtEnd = forY.InsertAfterSelf(new CommentRangeEnd() { Id = initialCommentId });


                    // add a commentReference after commentRangeEnd
                    cmtEnd.InsertAfterSelf(new Run(new CommentReference() { Id = initialCommentId }));

                    document.Save();

                    Console.WriteLine("Finished");
                }
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
