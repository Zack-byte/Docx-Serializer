public class ShadowParagraph
{
    public IEnumerable<TextRun> Runs { get; set; }


    public ShadowParagraph(IEnumerable<TextRun> runs) {
        Runs = runs;
    }
}