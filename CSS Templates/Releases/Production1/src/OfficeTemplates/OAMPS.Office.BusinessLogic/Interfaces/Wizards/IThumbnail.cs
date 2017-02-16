
using System.IO;

namespace OAMPS.Office.BusinessLogic.Interfaces.Wizards
{
 public  interface IThumbnail
    {
        Stream ImageStream { get; set; }
        string ImageTitle { get; set; }
        int Order { get; set; }
        string FullImageUrl { get; set; }
        string HeaderType { get; set; }
        string ABN { get; set; }
        string AFSL { get; set; }
        string WebSite { get; set; }
        string ShortDescription { get; set; }
        string LongDescription { get; set; }
    }
}
