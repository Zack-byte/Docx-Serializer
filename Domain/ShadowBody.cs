public class ShadowBody
{
    public IEnumerable<ShadowParagraph> Paragraphs { get; set; }
    public DocumentProperties BodyProperties { get; set; }

    public ShadowBody(IEnumerable<ShadowParagraph> paragraphs, DocumentProperties bodyProperties) {
        Paragraphs = paragraphs;
        BodyProperties = bodyProperties;
    }
}