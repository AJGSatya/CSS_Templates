using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Models.Wizards
{
    public class Insurer : IInsurer
    {
        public string Title { get; set; }
        public string Id { get; set; }
        public string Category { get; set; }
        public bool Selected { get; set; }
        public decimal Percent { get; set; }
    }
}