namespace OAMPS.Office.BusinessLogic.Interfaces.Wizards
{
    public interface IThumbnail
    {
        byte[] ImageStream { get; set; }
        string ImageTitle { get; set; }
        int Order { get; set; }
        string FullImageUrl { get; set; }
        string HeaderType { get; set; }
        string Abn { get; set; }
        string Afsl { get; set; }
        string WebSite { get; set; }
        string ShortDescription { get; set; }
        string LongDescription { get; set; }
        string RelativeUrl { get; set; }
    }
}