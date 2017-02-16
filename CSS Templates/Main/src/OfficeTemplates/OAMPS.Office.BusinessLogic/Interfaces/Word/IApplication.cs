namespace OAMPS.Office.BusinessLogic.Interfaces.Word
{
    public interface IApplication : IBaseView
    {
        int WindowCount { get; }
        string Name { get; }
        string Version { get; }
        void CloseActiveWindow(bool saveChanges);
    }
}