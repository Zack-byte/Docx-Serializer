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
            SectionProperties sectionProperties = myDocument.MainDocumentPart.Document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
            PageMargin pageMargin = sectionProperties.ChildElements.OfType<PageMargin>().FirstOrDefault();
            PageSize pageSize = sectionProperties.ChildElements.OfType<PageSize>().FirstOrDefault();
            DocumentPageMargin margin = new DocumentPageMargin();
            DocumentPageSize size = new DocumentPageSize();
            margin.Bottom = pageMargin.Bottom;
            margin.Left = pageMargin.Left;
            margin.Right = pageMargin.Right;
            margin.Top = pageMargin.Top;
            margin.Footer = pageMargin.Footer;
            margin.Header = pageMargin.Header;
            size.Height = pageSize.Height;
            size.Width = pageSize.Width;

            
            ShadowBody body = new ShadowBody(paragraphList.Select(p => new ShadowParagraph(p.ChildElements.OfType<Run>().Select(r => new
            TextRun(
            
                r.InnerText,
                r.ChildElements.OfType<RunProperties>().FirstOrDefault()?.Select(rp => new TextRunAttribute( rp.GetType().Name, rp.GetAttributes().FirstOrDefault().Value ))
            )))), new DocumentProperties(new DocumentHeaderReference(), new DocumentFooterReference(), size, margin, new DocumentPageNumberType(), new DocumentTitlePage()));

            ShadowDocument document = new ShadowDocument(body);
            var json = JsonConvert.SerializeObject(document, Newtonsoft.Json.Formatting.Indented);


            return json;
        }
    }


}