using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Presenters;

namespace OAMPS.Office.Tests.UnitTest
{
    /// <summary>
    ///     This is a test class for DocumentPresenterTest and is intended
    ///     to contain all DocumentPresenterTest Unit Tests
    /// </summary>
    [TestClass]
    public class InsuranceRenewealReportPresenterTest
    {
        #region UNIT TEST MOCK OBJECTS

        public class UnitTestDocument : IDocument
        {
            public IList<string> ImageAltText;
            public IList<string> ImageUrl;
            public List<string> InsertedFiles = new List<String>();
            public List<string> InsertedPageBreaks = new List<String>();
            public int MockPageCount { get; set; }
            public List<string> MockBookmarks { get; set; }
            public List<string> MockDocumentProperties { get; set; }

            public string Name
            {
                get { return "Unit Test"; }
            }

            public string DocumentPath
            {
                get { return "c:\fakepath"; }
            }

            public int PageCount
            {
                get { return MockPageCount; }
            }


            public bool HasPassword
            {
                get { throw new NotImplementedException(); }
            }

            public void CloseMe(bool saveChanges)
            {
                throw new NotImplementedException();
            }

            public void CloseMe(bool saveChanges, bool delayClose)
            {
                throw new NotImplementedException();
            }

            public void DeleteAllComments()
            {
                throw new NotImplementedException();
            }

            public void DeletePage(int pageNumber)
            {
                MockPageCount = PageCount - 1;
            }

            public void DeleteRange(int start, int end)
            {
                throw new NotImplementedException();
            }

            public void DeleteCharacter(int count)
            {
                throw new NotImplementedException();
            }

            public void DeleteBookmark(string name, bool deleteContent)
            {
                MockBookmarks.Remove(name);
            }

            public void ChangeDocumentImages(IList<string> imageAltText, IList<string> imageUrl)
            {
                ImageAltText = imageAltText;
                ImageUrl = imageUrl;
            }

            public void ChangeDocumentImagesBody(IList<string> imageAltText, IList<string> imageUrl)
            {
                throw new NotImplementedException();
            }

            public void PopulateControl(string tag, string value)
            {
                throw new NotImplementedException();
            }

            public void DeleteControl(string tag)
            {
                throw new NotImplementedException();
            }

            public void CloseInformationPanel(bool delayClose = false)
            {
                throw new NotImplementedException();
            }

            public string ReadContentControlValue(string tag)
            {
                throw new NotImplementedException();
            }


            public void UpdateToc()
            {
                throw new NotImplementedException();
            }

            public void UpdateOrCreatePropertyValue(string propertyName, string value)
            {
            }


            public string GetPropertyValue(string propertyName)
            {
                foreach (string p in  MockDocumentProperties)
                {
                    if (String.Equals(p, propertyName, StringComparison.OrdinalIgnoreCase))
                        return p;
                }
                return string.Empty;
            }


            public void SetCurrentRangeText(string text)
            {
                throw new NotImplementedException();
            }

            public void RenameControl(string tag, string newTag)
            {
                throw new NotImplementedException();
            }


            public void RenameControlInBookmark(string tag, string newTag, string bookmarkName)
            {
                throw new NotImplementedException();
            }

            public void CloseInformationPanelOnAllDocuments()
            {


            }

            public void AddBookmarkToBookmark(string name, string newName = "")
            {
                throw new NotImplementedException();
            }

            public void CloneBookmarkContent(string name, string newName = "")
            {
                MockBookmarks.Add(name);
            }


            public void AddBookmarkToCurrentLocation(string bookmarkNamE)
            {
                throw new NotImplementedException();
            }


            public void PopulateControlInBookmark(string bookmarkName, string tag, string value)
            {
                throw new NotImplementedException();
            }


            public List<string> GetBookmarksByPartialName(string name)
            {
                throw new NotImplementedException();
            }


            public void UpdateFields()
            {
                throw new NotImplementedException();
            }

            public void InsertFile(string path)
            {
                InsertedFiles.Add(path);
            }

            public void InsertLocalFile(string serverRelativeUrl)
            {
                throw new NotImplementedException();
            }

            public bool MoveCursorToStartOfBookmark(string name)
            {
                return MockBookmarks.Contains(name);

                //  throw new NotImplementedException();
            }


            public void MoveCursorPastStartOfBookmark(string name, int moveCharacterCount)
            {
                throw new NotImplementedException();
            }


            public void RenameBookmark(string name, string newName)
            {
                throw new NotImplementedException();
            }


            public IDocument OpenFile(string path)
            {
                throw new NotImplementedException();
            }

            public IDocument OpenFile(string path, string cacheName)
            {
                throw new NotImplementedException();
            }

            public bool FileExistsInSharePoint(string serverRelativeUrl)
            {
                throw new NotImplementedException();
            }

            public void TypeText(string text, string style = "")
            {
                throw new NotImplementedException();
            }


            public bool HasBookmark(string name)
            {
                throw new NotImplementedException();
            }

            public void InsertPageAbove()
            {
                throw new NotImplementedException();
            }


            public void MoveCursorUp(int count)
            {
                throw new NotImplementedException();
            }


            public void InsertPageBreak()
            {
                InsertedPageBreaks.Add(DateTime.Now.Second.ToString(CultureInfo.InvariantCulture));
            }


            public void ChangePageOrientToPortrait()
            {
                throw new NotImplementedException();
            }

            public void ChangePageOrientToLandscape()
            {
                throw new NotImplementedException();
            }


            public void MoveToStartOfDocument()
            {
                throw new NotImplementedException();
            }

            public void MoveToEndOfDocument()
            {
                throw new NotImplementedException();
            }

            public void SendEmail(string body)
            {
                throw new NotImplementedException();
            }

            public void SendEmail(string subject, string body, string recipients)
            {
                throw new NotImplementedException();
            }

            public void ImportPDF(string path)
            {
                throw new NotImplementedException();
            }


            public void InsertSectionBreak()
            {
                throw new NotImplementedException();
            }

            public void SetMargins(float leftMargin, float rightMargin)
            {
                throw new NotImplementedException();
            }

            public void SetTopMargins(float margin)
            {
                throw new NotImplementedException();
            }

            public void UnlinkDocumentFooterAndHeader()
            {
                throw new NotImplementedException();
            }


            public void SelectTable(string name)
            {
                throw new NotImplementedException();
            }

            public int TableRowOrColumnCount(bool isColumnCount, string name = "")
            {
                throw new NotImplementedException();
            }

            public void InsertTableRow(int insertRowPosition, string name = "")
            {
                throw new NotImplementedException();
            }

            public void PopulateTableCell(int row, int column, string value, string name = "")
            {
                throw new NotImplementedException();
            }

            public string ReadTableCell(int row, int column, string tableName = "")
            {
                throw new NotImplementedException();
            }

            public bool IsRangeReadOnly()
            {
                throw new NotImplementedException();
            }

            public int TurnOffProtection(string password)
            {
                throw new NotImplementedException();
            }

            public void TurnOnProtection(int type, string password)
            {
                throw new NotImplementedException();
            }

            public void InsertParagraphBreak()
            {
                throw new NotImplementedException();
            }

            public void InsertStyleSeperator()
            {
                throw new NotImplementedException();
            }

            public void ResetListNumbering()
            {
                throw new NotImplementedException();
            }

            public void InsertBackspace()
            {
                throw new NotImplementedException();
            }

            public void InsertRealPageBreak()
            {
                throw new NotImplementedException();
            }

            public string FindTextByStyleForCurrentDocument(string style)
            {
                throw new NotImplementedException();
            }

            public void RemoveHeader()
            {
                throw new NotImplementedException();
            }

            public void RemoveFooter()
            {
                throw new NotImplementedException();
            }

            public void SwitchScreenUpdating(bool show)
            {
                throw new NotImplementedException();
            }

            public void DeleteImage(string imageAltText)
            {
                throw new NotImplementedException();
            }

            public void Activate()
            {
                throw new NotImplementedException();
            }

            public void InsertTableCell(int insertBeforeColumnIndex, string cellText, string tableName = "")
            {
                throw new NotImplementedException();
            }

            public void RenameTable(string tableName, string newName)
            {
                throw new NotImplementedException();
            }

            public void MoveCursorRight(int count)
            {
                throw new NotImplementedException();
            }

            public void MoveCursorLeft(int count)
            {
                throw new NotImplementedException();
            }

            public void PasteClipboard()
            {
                throw new NotImplementedException();
            }

            public void CopyRange(int startRange, int endRange)
            {
                throw new NotImplementedException();
            }

            public int GetBookmarkStartRange(string bookmarkName)
            {
                throw new NotImplementedException();
            }

            public int GetBookmarkEndRange(string bookmarkName)
            {
                throw new NotImplementedException();
            }

            public void InsertTableRowNonCellLoop(int insertRowPosition, string name = "")
            {
                throw new NotImplementedException();
            }

            public void OpenFileWithData(IBaseTemplate template)
            {
                throw new NotImplementedException();
            }

            public int TableRowCount(string name = "")
            {
                throw new NotImplementedException();
            }
        }

        public class UnitTestDocumentView : IBaseView
        {
            public string ReturnMessage { get; set; }

            public void DisplayMessage(string text, string caption)
            {
                ReturnMessage = text;
            }
        }

        #endregion

        /// <summary>
        ///     Gets or sets the test context which provides
        ///     information about and functionality for the current test run.
        /// </summary>
        public TestContext TestContext { get; set; }

        #region Additional test attributes

        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //

        #endregion

        /// <summary>
        ///     A test for LoadIncludedPolicyClasses
        /// </summary>
        [TestMethod]
        public void LoadIncludedPolicyClassesTest()
        {
            var document = new UnitTestDocument();
            IBaseView view = null;
            document.MockDocumentProperties = new List<string>();

            var target = new InsuranceRenewealReportWizardPresenter(document, view);
            var template = new InsuranceRenewalReport
                {
                    ClientName = "client name test",
                    ClientCommonName = "client common name test"
                };

            var expected = new InsuranceRenewalReport
                {
                    ClientName = "client name test",
                    ClientCommonName = "client common name test"
                };
            IInsuranceRenewalReport actual = target.LoadIncludedPolicyClasses(template);
            Assert.AreEqual(expected.ClientCommonName, actual.ClientCommonName);
            Assert.AreEqual(expected.ClientName, actual.ClientName);
            Assert.AreEqual(expected.CoverPageTitle, actual.CoverPageTitle);
        }

        /// <summary>
        ///     A test for CheckName
        /// </summary>
        [TestMethod]
        public void CheckNameTest()
        {
            var document = new UnitTestDocument();
            IBaseView view = null;
            var target = new InsuranceRenewealReportWizardPresenter(document, view);
            const string name = "Unit Test";
            bool actual = target.CheckName(name);
            Assert.AreEqual(true, actual);
        }

        /// <summary>
        ///     A test for DeletePage
        /// </summary>
        [TestMethod]
        public void DeletePageTest()
        {
            var document = new UnitTestDocument {MockPageCount = 3};
            var view = new UnitTestDocumentView();
            var target = new InsuranceRenewealReportWizardPresenter(document, view);
            int pageNumber = 2;
            target.DeletePage(pageNumber);
            Assert.AreSame(null, view.ReturnMessage);
            Assert.AreEqual(document.MockPageCount, 2);

            document.MockPageCount = 1;
            pageNumber = 1;
            target.DeletePage(pageNumber);
            Assert.AreNotSame(null, view.ReturnMessage);
        }

        /// <summary>
        ///     A test for PopulateGraphics
        /// </summary>
        [TestMethod]
        public void PopulateGraphicsTest()
        {
            var document = new UnitTestDocument {MockPageCount = 5};
            var template = new BaseTemplate {CoverPageTitle = "a cover page", CoverPageImageUrl = "http://mockUrl", LogoTitle = "a logo", LogoImageUrl = "http://mockUrl"};

            var view = new UnitTestDocumentView();
            var target = new InsuranceRenewealReportWizardPresenter(document, view);

            target.PopulateGraphics(template, String.Empty, String.Empty);
            Assert.IsNotNull(document.ImageAltText[0]);
            Assert.IsNotNull(document.ImageAltText[1]);
            Assert.IsNotNull(document.ImageUrl[0]);
            Assert.IsNotNull(document.ImageUrl[1]);

            target.PopulateGraphics(template, "previousCoverPage", String.Empty);
            Assert.IsNotNull(document.ImageAltText[0]);
            Assert.IsNotNull(document.ImageAltText[1]);
            Assert.IsNotNull(document.ImageUrl[0]);
            Assert.IsNotNull(document.ImageUrl[1]);

            target.PopulateGraphics(template, String.Empty, "previous Logo");
            Assert.IsNotNull(document.ImageAltText[0]);
            Assert.IsNotNull(document.ImageAltText[1]);
            Assert.IsNotNull(document.ImageUrl[0]);
            Assert.IsNotNull(document.ImageUrl[1]);

            target.PopulateGraphics(template, "Previous Cover Page", "previous Logo");
            Assert.IsNotNull(document.ImageAltText[0]);
            Assert.IsNotNull(document.ImageAltText[1]);
            Assert.IsNotNull(document.ImageUrl[0]);
            Assert.IsNotNull(document.ImageUrl[1]);

            target.PopulateGraphics(template, "a cover page", "a logo");
            Assert.IsNull(document.ImageAltText[0]);
            Assert.IsNull(document.ImageAltText[1]);
            Assert.IsNull(document.ImageUrl[0]);
            Assert.IsNull(document.ImageUrl[1]);


            target.PopulateGraphics(template, "a cover page", "something here");
            Assert.IsNull(document.ImageAltText[0]);
            Assert.IsNotNull(document.ImageAltText[1]);
            Assert.IsNull(document.ImageUrl[0]);
            Assert.IsNotNull(document.ImageUrl[1]);

            target.PopulateGraphics(template, "somerthing here", "a logo");
            Assert.IsNotNull(document.ImageAltText[0]);
            Assert.IsNull(document.ImageAltText[1]);
            Assert.IsNotNull(document.ImageUrl[0]);
            Assert.IsNull(document.ImageUrl[1]);
        }

        /// <summary>
        ///     A test for PopulatePolicy
        /// </summary>
        [TestMethod]
        public void PopulateBasisOfCoverTest()
        {
            var document = new UnitTestDocument {MockBookmarks = new List<string> {"aBookMark", "bBookMark", "cBookMark"}};

            var view = new UnitTestDocumentView();
            var target = new InsuranceRenewealReportWizardPresenter(document, view);
            var frags = new List<IPolicyClass>();
            
            
            for (int i = 0; i < 2; i++)
            {
                var newFrag = new PolicyClass {MajorClass = "A", Title = "A" + i, Url = "http//A" + i};
                frags.Add(newFrag);
            }

            for (int i = 0; i < 3; i++)
            {
                var newFrag = new PolicyClass {MajorClass = "B", Title = "B" + i, Url = "http//B" + i};
                frags.Add(newFrag);
            }
            target.PopulateBasisOfCover(frags, "http://templates.oamps.com.au/Fragments/Class%20of%20Insurance.docx");
        }


        /// <summary>
        ///     A test for PopulateImportantNotices
        /// </summary>
        [TestMethod]
        public void PopulateImportantNoticesTest()
        {
            var document = new UnitTestDocument();
            var view = new UnitTestDocumentView();
            var target = new InsuranceRenewealReportWizardPresenter(document, view);
            const string im = "http://mockImportantNotes.doc";
            const string priv = "http://mockPrivacy.doc";
            const string fsg = "http://mockFSG.doc";
            const string toe = "http://mockTOE.doc";
            document.MockBookmarks = new List<string> {Constants.WordBookmarks.ImportantNotes};

            target.PopulateImportantNotices(Enums.Statutory.Retail, im, priv, fsg, toe);
            Assert.AreEqual(3, document.InsertedFiles.Count);

            document.InsertedFiles.Clear();
            target.PopulateImportantNotices(Enums.Statutory.Wholesale, im, priv, fsg, toe);
            Assert.AreEqual(3, document.InsertedFiles.Count);

            document.InsertedFiles.Clear();
            target.PopulateImportantNotices(Enums.Statutory.WholesaleWithRetail, im, priv, fsg, toe);
            Assert.AreEqual(4, document.InsertedFiles.Count);
        }
    }
}