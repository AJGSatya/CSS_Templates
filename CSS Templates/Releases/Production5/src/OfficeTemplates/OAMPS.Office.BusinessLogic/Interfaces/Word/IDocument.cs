using System;
using System.Collections.Generic;

namespace OAMPS.Office.BusinessLogic.Interfaces.Word
{
    public interface IDocument
    {
        #region Properties
        string Name { get; }
        string DocumentPath { get; }
        int PageCount { get; }
        Boolean HasPassword { get; }
        #endregion

        #region Document
        void CloseMe(bool saveChanges);
        void CloseMe(bool saveChanges, bool delayClose);
        void CloseInformationPanel(bool delayClose = false);
        void CloseInformationPanelOnAllDocuments();
        int TurnOffProtection(string password);
        void TurnOnProtection(int type, string password);
        void ChangeDocumentImages(IList<string> imageAltText, IList<string> imageUrl);
        void ChangeDocumentImagesBody(IList<string> imageAltText, IList<string> imageUrl);
        void RemoveHeader();
        void RemoveFooter();
        void SwitchScreenUpdating(bool show);
        string FindTextByStyleForCurrentDocument(string style);
        void SendEmail(string subject, string body, string recipients);
        void ImportPDF(string path);
        void DeletePage(int pageNumber = -1);
        void DeleteRange(int start, int end);
        void DeleteCharacter(int count);
        void Activate();
        void InsertLandscapePage(string markWithbookmark = "");
        void InsertPortraitPage(string markWithbookmark = "");
        #endregion

        #region Content Controls
        void PopulateControl(string tag, string value);
        void DeleteControl(string tag);
        string ReadContentControlValue(string tag);
        void RenameControl(string tag, string newTag);
        void RenameControlInBookmark(string tag, string newTag, string bookmarkName);
        void PopulateControlInBookmark(string bookmarkName, string tag, string value);
        List<string> GetContentControlsByName(string prefix);
        #endregion

        #region Bookmarks
        void AddBookmarkToBookmark(string name, string newName);
        void AddBookmarkToCurrentLocation(string bookmarkName);
        void CloneBookmarkContent(string name, string newName = "");
        void DeleteBookmark(string name, bool deleteContent);
        List<string> GetBookmarksByPartialName(string name);
        bool HasBookmark(string name);
        void RenameBookmark(string name, string newName);
        Int32 GetBookmarkStartRange(string bookmarkName);
        Int32 GetBookmarkEndRange(string bookmarkName);
        #endregion

        #region Comments
        void DeleteAllComments();
        #endregion

        #region Fields
        void UpdateFields();
        void UpdateToc();
        #endregion

        #region Properties
        void UpdateOrCreatePropertyValue(string propertyName, string value);
        string GetPropertyValue(string propertyName);
        #endregion

        #region  Cursor
        bool MoveCursorToStartOfBookmark(string name);
        void MoveCursorPastStartOfBookmark(string name, int moveCharacterCount);
        void MoveToStartOfDocument();
        void MoveToEndOfDocument();
        void InsertSectionBreak();
        void TypeText(string text, string style = "");
        void InsertPageAbove();
        void MoveCursorUp(int count);
        void InsertPageBreak();
        void MoveCursorRight(int count);
        void MoveCursorLeft(int count);
        void InsertParagraphBreak();
        void InsertStyleSeperator();
        void ResetListNumbering();
        void InsertBackspace();
        void InsertRealPageBreak();
        #endregion

        #region Section
        void ChangePageOrientToPortrait();
        void ChangePageOrientToLandscape();
        void SetMargins(float leftMargin, float rightMargin);
        void SetTopMargins(float margin);
        void UnlinkDocumentFooterAndHeader();
        #endregion

        #region Files
        void InsertFile(string serverRelativeUrl);
        void InsertLocalFile(string serverRelativeUrl);
        IDocument OpenFile(string path);
        IDocument OpenFile(string path, string cacheName);
        bool FileExistsInSharePoint(string serverRelativeUrl);
        void LogInsertFile(string serverRelativeUrl);
        #endregion

        #region Tables
        void SelectTable(string name);
        int TableRowOrColumnCount(bool isColumnCount, string name = "");
        void InsertTableRow(int insertRowPosition, string name = "");
        void PopulateTableCell(int row, int column, string value, string name = "");
        string ReadTableCell(int row, int column, string tableName = "");
        void InsertTableCell(int insertBeforeColumnIndex, string cellText, string tableName = "");
        void RenameTable(string tableName, string newName);
        void InsertTableRowNonCellLoop(int insertRowPosition, string name = "");
        void CopyTable(string name);
        void DeleteTable(string name);
        List<string> GetTablesByNamePrefix(string prefix);
        bool TableExists(string name);
        string GetTableCellText(string tableName, int row, int column);
        void TypeTableCellText(string tableName, int row, int column, string text);
        #endregion

        #region Images
        void DeleteImage(string imageAltText);
        #endregion

        #region Range
        bool IsRangeReadOnly();
        void SetCurrentRangeText(string text);
        void PasteClipboard();
        void PasteClipboardOriginalFormatting();
        void CopyRange(int startRange, int endRange);
        void DeleteRangeWithBackspace(int start, int end);

        #endregion

    }
}
