

using OAMPS.Office.BusinessLogic.Interfaces.Word;

namespace OAMPS.Office.BusinessLogic.Presenters
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
