public class TextRunAttribute {
    public string Name { get; set; }
    public string Value { get; set; }

    public TextRunAttribute(string name, string value) {
        Name = name;
        Value = value;
    }
}