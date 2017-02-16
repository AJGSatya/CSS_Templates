using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;

namespace OAMPS.Office.BusinessLogic.Presenters.Word
{
    public class HelpContentWizardPresenter : BaseWizardPresenter
    {
        public HelpContentWizardPresenter(IDocument document, IBaseView view)
            : base(document, view)
        {
        }

        public string FindHeadingTextForCurrentDocument()
        {
            return Document.FindTextByStyleForCurrentDocument(Constants.WordStyles.Heading1);
        }
    }
}