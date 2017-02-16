using OAMPS.Office.BusinessLogic.Interfaces.Word;

namespace OAMPS.Office.BusinessLogic.Presenters.Word
{
    public class ApplicatonPresenter
    {
        private readonly IApplication _application;

        public ApplicatonPresenter(IApplication application)
        {
            _application = application;
        }

        public void CloseActiveWindow(bool saveChanges)
        {
            _application.CloseActiveWindow(saveChanges);
        }
    }
}