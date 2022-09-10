using System.Xml;
using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;

public class DocxRepository
{
    public DocxRepository()
    {
    }


    public void extractDocxArchive()
    {
        var extractPath = @".\assets\unzipped";
        var zipPath = @".\assets\unzipped\Project proposal.zip";

        Console.WriteLine($"Extracting Zip Archive {zipPath} to Extract Folder {extractPath}");
        ZipFile.ExtractToDirectory(zipPath, extractPath);
        Console.WriteLine($"Successfully Extracted Zip Archive {zipPath} to Extract Folder {extractPath}");
    }

    public void createNewWordDocument()
    {
        var fileName = @".\outputs\testdocument.docx";

        //Create a document
        using (WordprocessingDocument myDocument =
        WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
        {
            // Add a main body
            MainDocumentPart mainPart = myDocument.AddMainDocumentPart();

            // Create the document structure.
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());

            // Add some text to the document.
            run.AppendChild(new Text("Hello, World!"));
        }

        Console.WriteLine($"Word Processing Document has been created at {fileName}");
    }

    public string extractDocxDocumentAsJson()
    {

        var fileName = @".\assets\Project proposal.docx";

        using (WordprocessingDocument myDocument = WordprocessingDocument.Open(fileName, true))
        {

            IEnumerable<Paragraph> paragraphList = myDocument.MainDocumentPart.Document.Body.Elements().OfType<Paragraph>();
            var proj = paragraphList.Select(p => p.ChildElements.OfType<Run>().Select(r => new
            {
                r.InnerText,
                runProperties = r.ChildElements.OfType<RunProperties>().FirstOrDefault()?.Select(rp => new { rp.GetType().Name, Val = rp.GetAttributes().FirstOrDefault().Value })
            }));
            var json = JsonConvert.SerializeObject(proj, Newtonsoft.Json.Formatting.Indented);


            return json;
        }
    }


}