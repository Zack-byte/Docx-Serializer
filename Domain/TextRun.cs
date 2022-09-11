public class TextRun {
    public string Text { get; set; }
    public IEnumerable<TextRunAttribute> Attributes { get; set; }

    public TextRun(string text, IEnumerable<TextRunAttribute> attributes) {
        Text = text;
        Attributes = attributes;
    }

}