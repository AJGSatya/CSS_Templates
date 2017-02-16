namespace OAMPS.Office.BusinessLogic.Interfaces.Logging
{
    public enum Type
    {
        Error,
        Information,
        Warning,
        Debug
    }

    public interface IBaseLogger
    {
        void Log(string message, Type type);
    }
}