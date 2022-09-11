public class DocumentProperties
{
    public DocumentHeaderReference HeaderReference { get; set; }
    public DocumentFooterReference FooterReference { get; set; }
    public DocumentPageSize PageSize { get; set; }
    public DocumentPageMargin PageMargin { get; set; }
    public DocumentPageNumberType PageNumberType { get; set; }
    public DocumentTitlePage TitlePage { get; set; }

    public DocumentProperties(
        DocumentHeaderReference headerReference, 
        DocumentFooterReference footerReference, 
        DocumentPageSize pageSize, 
        DocumentPageMargin pageMargin, 
        DocumentPageNumberType pageNumberType, 
        DocumentTitlePage titlePage
    )
    {
        HeaderReference = headerReference;
        FooterReference = footerReference;
        PageSize = pageSize;
        PageMargin = pageMargin;
        PageNumberType = pageNumberType;
        TitlePage = titlePage;
    }
}