using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Word;

namespace OAMPS.Office.BusinessLogic.Presenters.Wizards
{
    public class SummaryOfDiscussionWizardPresenter : BaseWizardPresenter
    {
        public SummaryOfDiscussionWizardPresenter(IDocument document, IBaseView view)
            : base(document, view)
        {
        }

        public void CreatePropertiesForRadioButtons(bool isClient, bool isCaller, bool isUnerwriter, bool isByPhone,
            bool isInPerson, bool isOther)
        {
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.FileNoteIsClient, isByPhone.ToString());
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.FileNoteIsCaller, isCaller.ToString());
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.FileNoteIsUnderwriter,
                isUnerwriter.ToString());
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.FileNoteByPhone, isByPhone.ToString());
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.FileNoteInPerson,
                isInPerson.ToString());
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.FileNoteOther, isOther.ToString());
        }
    }
}