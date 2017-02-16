using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
namespace OAMPS.Office.BusinessLogic.Presenters
{
    public class ClientDiscoveryPresenter : BasePresenter
    {

        public ClientDiscoveryPresenter(IDocument document, IBaseView view)
            : base(document, view)
        {                        
        }

        public string ReadSelectedQuestionsFromDocument()
        {
            return Document.GetPropertyValue(Helpers.Constants.WordDocumentProperties.DiscoveryQuestions);
        }

        public void PopulateClientQuestions(List<IQuestionClass> questions)
        {
            string storedKey = null;

            Dictionary<string, List<IQuestionClass>> fragements;
            fragements = questions.GroupBy(x => x.TopCategory).ToDictionary(x => x.Key, x => x.ToList());
            if (Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.ClientDiscoveryQuestions))
            {
                fragements.ToList().ForEach((x) =>
                {
                    //TODO FIX THE STYLE
                    Document.TypeText(x.Key, Constants.WordStyles.Heading3Black);
                    Document.InsertParagraphBreak();

                    var sorted = x.Value.GroupBy(y => y.SubCategory).ToDictionary(y => y.Key, y => y.ToList());

                    sorted.ToList().ForEach((keyValuePair) =>
                    {
                        if (!string.IsNullOrEmpty(keyValuePair.Key))
                        {
                            Document.TypeText(keyValuePair.Key, "Sub Category");
                            Document.InsertParagraphBreak();
                        }
                        keyValuePair.Value.ForEach((questionClass) =>
                        {
                            storedKey += questionClass.Id + ";";

                            //  Document.ResetListNumbering();
                            Document.TypeText(questionClass.Title, "Number List");

                            Document.InsertParagraphBreak(); //these three para breaks control the business rule to have 3 breaks between each question                            Document.InsertParagraphBreak();
                            // Document.TypeText(" ", "Normal");
                            Document.TypeText("\t", "Normal");
                            //Document.InsertBackspace(); //clear the numbering
                            Document.InsertParagraphBreak();
                            Document.InsertParagraphBreak();
                            Document.InsertParagraphBreak();
                            Document.InsertParagraphBreak();
                            Document.InsertParagraphBreak();
                        });
                    });
                });
                Document.UpdateOrCreatePropertyValue(Helpers.Constants.WordDocumentProperties.DiscoveryQuestions, storedKey);
            }                
        }
    }
}
