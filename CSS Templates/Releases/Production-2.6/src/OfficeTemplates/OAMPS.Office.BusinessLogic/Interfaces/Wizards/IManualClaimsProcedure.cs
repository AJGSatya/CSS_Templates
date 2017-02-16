namespace OAMPS.Office.BusinessLogic.Interfaces.Wizards
{
    public interface IManualClaimsProcedure
    {
        string Title { get; set; }
        string PolicyClass { get; set; }
        string Url { get; set; }
        string Id { get; set; }
    }
}