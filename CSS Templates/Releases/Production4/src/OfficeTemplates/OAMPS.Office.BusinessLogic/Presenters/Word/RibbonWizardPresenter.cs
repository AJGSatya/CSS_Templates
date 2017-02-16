using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;

namespace OAMPS.Office.BusinessLogic.Presenters.Word
{
    public class RibbonWizardPresenter : BaseWizardPresenter
    {
        public RibbonWizardPresenter(IDocument document, IBaseView view)
            : base(document, view)
        {
        }


        private void InsertLocalFactFinder()
        {
        }
    }
}