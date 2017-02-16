using System;

namespace OAMPS.Office.BusinessLogic.Interfaces.Word
{
    public interface IApplication : IBaseView
    {
        Int32 WindowCount { get; }
        string Name { get; }
        string Version { get; }
        void CloseActiveWindow(bool saveChanges);
    }
}