using System;
using System.Collections.Generic;
using OAMPS.Office.BusinessLogic.Interfaces.Template;

namespace OAMPS.Office.BusinessLogic.Interfaces.Word
{
    public interface IDocument
    {
        string Name { get; }
        string DocumentPath { get; }
        int PageCount { get; }
        Boolean HasPassword { get; }
        void CloseMe(bool saveChanges);
        void DeleteAllComments();
        void DeletePage(int pageNumber = -1);
        void ChangeDocumentImages(IList<string> imageAltText, IList<string> imageUrl);
        void PopulateControl(string tag, string value);
        void DeleteControl(string tag);
        string ReadContentControlValue(string tag);
        void CloseInformationPanel();
        void UpdateOrCreatePropertyValue(string propertyName, string value);
        string GetPropertyValue(string propertyName);
        void SetCurrentRangeText(string text);
        void CloneBookmarkContent(string name, string newName = "");
        void RenameControl(string tag, string newTag);
        void RenameControlInBookmark(string tag, string newTag, string bookmarkName);
        void AddBookmarkToBookmark(string name, string newName);
        void AddBookmarkToCurrentLocation(string bookmarkName);
        void PopulateControlInBookmark(string bookmarkName, string tag, string value);
        void DeleteBookmark(string name, bool deleteContent);
        List<string> GetBookmarksByPartialName(string name);
        void UpdateFields();
        void InsertFile(string serverRelativeUrl);
        void MoveCursorToStartOfBookmark(string name);
        void MoveCursorPastStartOfBookmark(string name, int moveCharacterCount);
        bool HasBookmark(string name);
        void RenameBookmark(string name, string newName);
        void OpenFile(string path);
        void OpenFile(string path, string cacheName);
        void TypeText(string text, string style = "");
        void InsertPageAbove();
        void MoveCursorUp(int count);
        void InsertPageBreak();
        void ChangePageOrientToPortrait();
        void ChangePageOrientToLandscape();
        void MoveToStartOfDocument();
        void SendEmail();
        void ImportPDF(string path);
        void InsertSectionBreak();
        void SetMargins(float margin);
        void SelectTable(string name);
        int TableRowCount(string name = "");
        void InsertTableRow(int insertRowPosition, string name = "");
        void PopulateTableCell(int row, int column, string value, string name = "");
        bool IsRangeReadOnly();
        int TurnOffProtection(string password);
        void TurnOnProtection(int type, string password);
        void InsertParagraphBreak();
        void InsertStyleSeperator();
        void ResetListNumbering();
        void InsertBackspace();

        void RemoveHeader();
        void RemoveFooter();

    }
}
