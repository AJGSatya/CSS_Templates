
using System.IO;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Models.Wizards
{
    public class Thumbnail : IThumbnail
    {
        public byte[] ImageStream { get; set; }
        public string ImageTitle { get; set; }
        public int Order { get; set; }
        public string FullImageUrl { get; set; }
        public string HeaderType { get; set; }
        public string ABN { get; set; }
        public string AFSL { get; set; }
        public string WebSite { get; set; }
        public string ShortDescription { get; set; }
        public string LongDescription { get; set; }
    }
}
