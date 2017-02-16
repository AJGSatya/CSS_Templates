using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;

namespace OAMPS.Office.BusinessLogic.Presenters
{
    public class AgendaWizardPresenter : BaseWizardPresenter
    {
        public AgendaWizardPresenter(IDocument document, IBaseView view)
            : base(document, view)
        {
        }

        public void InsertMinutesFragement(string path)
        {
            if (Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.AgendaMinutes))
            {
                Document.InsertFile(path);
            }
        }
    }
}