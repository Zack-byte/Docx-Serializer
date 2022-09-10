using Microsoft.AspNetCore.Mvc;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.Json;

namespace Docx_Serializer.Controllers;

[ApiController]
[Route("[controller]")]
public class DocxController : ControllerBase
{
    private readonly ILogger<DocxController> _logger;

    public DocxController(ILogger<DocxController> logger)
    {
        _logger = logger;
    }

    [HttpGet(Name = "GetDocx")]
    public Document Get()
    {
        var repo = new DocxRepository();
        var document = repo.extractDocxDocument();
        var options = new JsonSerializerOptions { WriteIndented = true };
        var json = System.Text.Json.JsonSerializer.Serialize<Document>(document, options);
        Console.WriteLine(document.ToString());
        return document;
    }
}
