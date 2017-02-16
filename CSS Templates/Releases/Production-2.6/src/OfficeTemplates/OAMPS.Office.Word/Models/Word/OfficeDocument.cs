using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Caching;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Properties;
using File = Microsoft.SharePoint.Client.File;
using WordOM = Microsoft.Office.Interop.Word;
using OutlookOM = Microsoft.Office.Interop.Outlook;


namespace OAMPS.Office.Word.Models.Word
{
    public class OfficeDocument : IDocument
    {
        private readonly WordOM.Document _document;

        public OfficeDocument(WordOM.Document document)
        {
            _document = document;
        }

        public int PageCount
        {
            get { return (int) _document.Application.Selection.Information[WordOM.WdInformation.wdNumberOfPagesInDocument]; }
        }

        public string DocumentPath
        {
            get { return _document.Path; }
        }

        public string Name
        {
            get { return _document.Name; }
        }

        public bool HasPassword
        {
            get { return _document.HasPassword; }
        }

        public void CloseMe(bool saveChanges)
        {
            _document.Close(saveChanges);
        }

        public void CloseMe(bool saveChanges, bool delayClose )
        {
            try
            {
                _document.Activate();
                if (delayClose)
                {
                    Task.Factory.StartNew(() => Thread.Sleep(2000)).ContinueWith(task => CloseMe(saveChanges, false));
                }
                else
                {
                    _document.Close(saveChanges);
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void DeleteAllComments()
        {
            _document.DeleteAllComments();
        }

        public void InsertRealPageBreak()
        {
            _document.Application.Selection.InsertBreak(WordOM.WdBreakType.wdPageBreak);
        }

        public string FindTextByStyleForCurrentDocument(string style)
        {
            WordOM.Find find = _document.Application.Selection.Find;

            int r = _document.Application.Selection.Move();


            //error handling needed.
            // ReSharper disable UseIndexedProperty
            find.set_Style(_document.Styles[style]);
            // ReSharper restore UseIndexedProperty
            find.Text = String.Empty;
            find.Forward = false;
            find.MatchWildcards = true;
            find.Execute();
            return _document.Application.Selection.Text;
        }

        public void PopulateControl(string tag, string value)
        {
            WordOM.ContentControls contentControls = _document.SelectContentControlsByTag(tag);
            if (contentControls == null) return;
            foreach (WordOM.ContentControl contentControl in contentControls)
            {
                if (contentControl.Type == WordOM.WdContentControlType.wdContentControlCheckBox)
                {
                    if (!String.IsNullOrEmpty(value)) contentControl.Checked = Boolean.Parse(value);
                }
                else
                {
                    if (value != null && !String.IsNullOrEmpty(value)) value = value.Replace("\r\n", ((char)11).ToString(CultureInfo.InvariantCulture));

                    contentControl.Range.Text = value;
                }

                if (value != null)
                {
                    contentControl.SetPlaceholderText(null, null, value);
                }
            }
        }

        public Int32 GetBookmarkStartRange(string bookmarkName)
        {
            if (!_document.Bookmarks.Exists(bookmarkName)) return -1;

            return _document.Bookmarks[bookmarkName].Range.Start;
        }

        public Int32 GetBookmarkEndRange(string bookmarkName)
        {
            if (!_document.Bookmarks.Exists(bookmarkName)) return -1;

            return _document.Bookmarks[bookmarkName].Range.End;
        }

        public void PasteClipboard()
        {
            _document.Activate();
            _document.Application.Selection.Paste();
        }

        public void CopyRange(int startRange, int endRange)
        {
            if(startRange < 0 || endRange < 0)
                return;

            _document.Range(startRange, endRange).Select();
            _document.Application.Selection.Copy();
        }

        public void PopulateControlInBookmark(string bookmarkName, string tag, string value)
        {
            if (!_document.Bookmarks.Exists(bookmarkName)) return;

            WordOM.ContentControls contentControls = _document.Bookmarks[bookmarkName].Range.ContentControls;
            if (contentControls == null) return;
            foreach (WordOM.ContentControl contentControl in contentControls)
            {
                contentControl.Range.Text = value;
            }
        }

        public void DeleteControl(string tag)
        {
            WordOM.ContentControls contentControls = _document.SelectContentControlsByTag(tag);
            if (contentControls == null) return;
            foreach (WordOM.ContentControl contentControl in contentControls)
            {
                contentControl.Delete(true);
            }
        }

        public string ReadContentControlValue(string tag)
        {
            WordOM.ContentControls contentControls = _document.SelectContentControlsByTag(tag);
            if (contentControls.Count < 1)
                return string.Empty;


            switch (contentControls[1].Type)
            {
                case WordOM.WdContentControlType.wdContentControlCheckBox:
                    return (contentControls.Count < 1) ? String.Empty : contentControls[1].Checked.ToString();
                default:
                    return (contentControls.Count < 1) ? String.Empty : contentControls[1].Range.Text;
            }
        }

        public void RenameControl(string tag, string newTag)
        {
            WordOM.ContentControls contentControls = _document.SelectContentControlsByTag(tag);
            if (contentControls == null) return;
            foreach (WordOM.ContentControl contentControl in contentControls)
            {
                contentControl.Tag = newTag;
            }
        }

        public void RenameControlInBookmark(string tag, string newTag, string bookmarkName)
        {
            if (!_document.Bookmarks.Exists(bookmarkName)) return;

            WordOM.ContentControls contentControls = _document.Bookmarks[bookmarkName].Range.ContentControls;
            if (contentControls == null) return;
            foreach (WordOM.ContentControl contentControl in contentControls)
            {
                contentControl.Tag = newTag;
            }
        }

        public void DeleteRange(int start, int end)
        {
            _document.Range(start, end).Delete();
        }

        public void DeleteCharacter(int count)
        {
            _document.Application.Selection.Delete(WordOM.WdUnits.wdCharacter, count);
             _document.Application.Selection.TypeBackspace();
        }


        public void DeletePage(int pageNumber = -1)
        {
            if (pageNumber > 0)
                _document.Application.Selection.GoTo(WordOM.WdGoToItem.wdGoToPage, WordOM.WdGoToDirection.wdGoToAbsolute, pageNumber);

            _document.Bookmarks[@"\Page"].Range.Delete();
        }

        public void CloseInformationPanelOnAllDocuments()
        {
            try 
            {
                for (var i = 1; i <= Globals.ThisAddIn.Application.Documents.Count; i++ ) //vsto arrays start at 1
                {
                    Globals.ThisAddIn.Application.Documents[i].Application.DisplayDocumentInformationPanel = false;
                }
            }
            catch (Exception)
            {

            }
        }

        public void CloseInformationPanel(bool delayClose = false)
        {
            try
            {
                _document.Activate();
                if (delayClose)
                {
                   // Task.Factory.StartNew(() => Thread.Sleep(2000)).Wait();
                    //CloseInformationPanel(false);
                    Task.Factory.StartNew(() => Thread.Sleep(2000)).ContinueWith(task => CloseInformationPanel(false));
                }
                else
                {
                   // _document.Application.DisplayDocumentInformationPanel = true;
                    _document.Application.DisplayDocumentInformationPanel = false;
                    _document.ActiveWindow.View.Zoom.Percentage = 100; //todo: THIS DOES NOT BELONG HERE, EXTRACT TO ITS OWN METHOD!
                }
            }
            catch (Exception ex)
            {
                try //i hate myself right now
                {
                    _document.Application.DisplayDocumentInformationPanel = false;
                }
                catch (Exception)
                {
                    
                }
            }
        }


        public void ChangeDocumentImagesBody(IList<string> imageAltText, IList<string> imageUrls)
        {
            var localImageUrls = new Dictionary<string, string>();
            var shapesCount = _document.Shapes.Count;
            var deleteableShapes = new List<WordOM.Shape>();
            
            for (var i = 1; i <= shapesCount; i++) //VSTO arrays start at 1
            {
                var j = i;
                var currentImage =_document.Shapes[j];
                
                if (!imageAltText.Contains(currentImage.AlternativeText)) continue;
                
                var imageIndex = imageAltText.IndexOf(currentImage.AlternativeText);
                if (String.IsNullOrEmpty(imageUrls[imageIndex])) continue;

                if (!localImageUrls.ContainsKey(imageAltText[imageIndex]))
                {
                    var localUrl = DownloadLocalImage(imageUrls[imageIndex]);
                    localImageUrls.Add(imageAltText[imageIndex], localUrl);
                }
                var key = imageAltText[imageIndex];
                var newImage = _document.Shapes.AddPicture(
                    localImageUrls[key], false,
                    true,
                    currentImage.Left,
                    currentImage.Top,
                    currentImage.Width,
                    currentImage.Height,
                    currentImage.Anchor);

                newImage.Apply();

                newImage.RelativeHorizontalPosition = currentImage.RelativeHorizontalPosition;
                newImage.RelativeVerticalPosition = currentImage.RelativeVerticalPosition;
                newImage.RelativeVerticalSize = currentImage.RelativeVerticalSize;
                newImage.RelativeHorizontalSize = currentImage.RelativeHorizontalSize;
                newImage.TopRelative = currentImage.TopRelative;
                newImage.LeftRelative = currentImage.LeftRelative;
                newImage.WidthRelative = currentImage.WidthRelative;
                newImage.HeightRelative = currentImage.HeightRelative;
                newImage.Width = currentImage.Width;
                newImage.Height = currentImage.Height;
                newImage.Left = currentImage.Left;
                newImage.Top = currentImage.Top;
                newImage.Title = imageAltText[imageIndex];
                newImage.AlternativeText = imageAltText[imageIndex];
                newImage.WrapFormat.Type = currentImage.WrapFormat.Type;
                newImage.LayoutInCell = currentImage.LayoutInCell;
                deleteableShapes.Add(currentImage);
            }

            foreach (var s in deleteableShapes)
            {
                s.Delete();
            }
        }

        //this method changes the cover page and company logo images that are selected on the wizard
        public void ChangeDocumentImages(IList<string> imageAltText, IList<string> imageUrls)
        {
            var localImageUrls = new Dictionary<string, string>();
            int shapesCount =
                _document.Sections[1].Headers[
                    WordOM.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count;
            var deleteableShapes = new List<WordOM.Shape>();

            for (int i = 1; i <= shapesCount; i++) //VSTO arrays start at 1
            {
                int j = i;
                WordOM.Shape currentImage =
                    _document.Sections[1].Headers[
                        WordOM.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[j];
                if (imageAltText.Contains(currentImage.AlternativeText))
                {
                    int imageIndex = imageAltText.IndexOf(currentImage.AlternativeText);


                    if (String.IsNullOrEmpty(imageUrls[imageIndex])) continue;

                    if (!localImageUrls.ContainsKey(imageAltText[imageIndex]))
                    {
                        string localUrl = DownloadLocalImage(imageUrls[imageIndex]);
                        localImageUrls.Add(imageAltText[imageIndex], localUrl);
                    }
                    string key = imageAltText[imageIndex];
                    WordOM.Shape newImage = _document.Sections[1].Footers[WordOM.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.AddPicture(
                        localImageUrls[key], false,
                        true,
                        currentImage.Left,
                        currentImage.Top,
                        currentImage.Width,
                        currentImage.Height,
                        currentImage.Anchor);

                    newImage.Apply();

                    newImage.RelativeHorizontalPosition = currentImage.RelativeHorizontalPosition;
                    newImage.RelativeVerticalPosition = currentImage.RelativeVerticalPosition;
                    newImage.RelativeVerticalSize = currentImage.RelativeVerticalSize;
                    newImage.RelativeHorizontalSize = currentImage.RelativeHorizontalSize;
                    newImage.TopRelative = currentImage.TopRelative;
                    newImage.LeftRelative = currentImage.LeftRelative;
                    newImage.WidthRelative = currentImage.WidthRelative;
                    newImage.HeightRelative = currentImage.HeightRelative;
                    newImage.Width = currentImage.Width;
                    newImage.Height = currentImage.Height;
                    newImage.Left = currentImage.Left;
                    newImage.Top = currentImage.Top;
                    newImage.Title = imageAltText[imageIndex];
                    newImage.AlternativeText = imageAltText[imageIndex];
                    newImage.WrapFormat.Type = currentImage.WrapFormat.Type;
                    newImage.LayoutInCell = currentImage.LayoutInCell;
                    deleteableShapes.Add(currentImage);
                }
            }

            foreach (WordOM.Shape s in deleteableShapes)
            {
                s.Delete();
            }
        }


        public void UpdateOrCreatePropertyValue(string propertyName, string value)
        {
            _document.UpdateOrCreatePropertyValue(propertyName, value);
        }

        public string GetPropertyValue(string propertyName)
        {
            return _document.GetPropertyValue(propertyName);
        }

        public void SetCurrentRangeText(string text)
        {
            _document.Application.Selection.InsertAfter(text);
        }

        public void AddBookmarkToBookmark(string name, string newName)
        {
            if (!_document.Bookmarks.Exists(name)) return;

            WordOM.Bookmark cBookmark = _document.Bookmarks[name];
            cBookmark.Copy(newName);
            //var nBookmark = _document.Application.Selection.Bookmarks.Add(newName);
        }

        public void AddBookmarkToCurrentLocation(string bookmarkName)
        {
            _document.Application.Selection.Bookmarks.Add(bookmarkName);
        }

        public bool HasBookmark(string name)
        {
            return _document.Bookmarks.Exists(name);
        }

        public void InsertPageAbove()
        {
            _document.Application.Selection.InsertBreak(WordOM.WdBreakType.wdPageBreak);
        }

        public void CloneBookmarkContent(string name, string newName = "")
        {
            if (!_document.Bookmarks.Exists(name)) return;

            _document.Bookmarks[name].Range.Copy();
            _document.Bookmarks[name].Range.Select();
            _document.Application.Selection.Collapse(WordOM.WdCollapseDirection.wdCollapseEnd);

            WordOM.Bookmark bs = _document.Application.Selection.Bookmarks.Add(newName + "start");
            bs.Range.InsertAfter(" ");
            _document.Application.Selection.Move(WordOM.WdUnits.wdCharacter, 1);
            WordOM.Bookmark be = _document.Application.Selection.Bookmarks.Add(newName + "end");
            _document.Range(bs.Start, bs.End).Paste();

            WordOM.Range contentRange = _document.Range(bs.Start, be.End);
            contentRange.Bookmarks.Add(newName);

            bs.Delete();
            be.Delete();
        }

        public void DeleteBookmark(string name, bool deleteContent)
        {
            if (!_document.Bookmarks.Exists(name)) return;

            if (deleteContent) _document.Bookmarks[name].Range.Delete();
            else _document.Bookmarks[name].Delete();
        }

        public void UpdateFields()
        {
            _document.Fields.Update();
        }


        public void UpdateToc()
        {
            if (_document.TablesOfContents.Count > 0)
                _document.TablesOfContents[1].Update();

        }

        public void InsertFile(string serverRelativeUrl)
        {
            if (string.IsNullOrEmpty(serverRelativeUrl)) return;

            //enhancement - we could not delete the file at the end of this method
            //enhancement - if we did not delete this file we could compqare its created date agasint the modified date in sharepoint and only take a copy if updates have occured.
            using (var clientContext = new ClientContext(Settings.Default.SharePointContextUrl))
            {
                var temp = Path.GetTempFileName();

                var fileInformation = File.OpenBinaryDirect(clientContext, serverRelativeUrl);
                var stream = fileInformation.Stream;

                using (var output = System.IO.File.OpenWrite(temp))
                {
                    stream.CopyTo(output);
                }

                _document.Application.Selection.InsertFile(temp, "", false, false);

                System.IO.File.Delete(temp);
            }
        }

        public void InsertLocalFile(string serverRelativeUrl)
        {
            if (string.IsNullOrEmpty(serverRelativeUrl)) return;
            
            _document.Application.Selection.InsertFile(serverRelativeUrl, "", false, false);
        }


      
        public bool MoveCursorToStartOfControl(string name)
        {
            _document.Activate();
            if (_document.SelectContentControlsByTag(name).Count > 0)
            {
                var controls = _document.SelectContentControlsByTag(name);
                if (controls.Count > 0)
                {
                    controls[1].Range.Select();
                    return true;
                }
            }
            return false;
        }

        public bool MoveCursorToStartOfBookmark(string name)
        {
            _document.Activate();
            var showHidden = _document.Bookmarks.ShowHidden;
            _document.Bookmarks.ShowHidden = true;
            if (_document.Bookmarks.Exists(name))
            {
                _document.Application.Selection.GoTo(What: WordOM.WdGoToItem.wdGoToBookmark, Name: name);
                _document.Bookmarks.ShowHidden = showHidden;
                return true;
            }
            return false;
        }

        public void MoveCursorUp(int count)
        {
            _document.Activate();
            _document.Application.Selection.MoveUp(WordOM.WdUnits.wdLine, count);
        }

        public void MoveCursorRight(int count)
        {
            _document.Activate();
            _document.Application.Selection.MoveRight(WordOM.WdUnits.wdCharacter, count);
        }

        public void MoveCursorLeft(int count)
        {
            _document.Activate();
            _document.Application.Selection.MoveLeft(WordOM.WdUnits.wdCharacter, count);
        }

        public void InsertPageBreak()
        {
            _document.Activate();
            _document.Application.Selection.InsertBreak(WordOM.WdBreakType.wdSectionBreakNextPage);
        }


        public void InsertSectionBreak()
        {
            _document.Activate();
            _document.Application.Selection.InsertBreak(WordOM.WdBreakType.wdSectionBreakContinuous);
        }


        public void ChangePageOrientToPortrait()
        {
            _document.Activate();
// ReSharper disable UseIndexedProperty
            if (_document.Application.Selection.get_Information(WordOM.WdInformation.wdWithInTable))
// ReSharper restore UseIndexedProperty
            {
                _document.Application.Selection.Tables[1].Select();
                _document.Application.Selection.Collapse(WordOM.WdCollapseDirection.wdCollapseEnd);
            }

            //if cursor didnt move out of table dont contintue else error will generate.
// ReSharper disable UseIndexedProperty
            if (_document.Application.Selection.get_Information(WordOM.WdInformation.wdWithInTable))
// ReSharper restore UseIndexedProperty
                return;

            _document.Application.Selection.PageSetup.Orientation = WordOM.WdOrientation.wdOrientPortrait;
        }

        public void ChangePageOrientToLandscape()
        {
            _document.Activate();
            _document.Application.Selection.PageSetup.Orientation = WordOM.WdOrientation.wdOrientLandscape;
        }

        public void MoveCursorPastStartOfBookmark(string name, int moveCharacterCount)
        {
            _document.Activate();
            MoveCursorToStartOfBookmark(name);
            _document.Application.Selection.MoveRight(Unit: WordOM.WdUnits.wdCharacter, Count: moveCharacterCount);
        }

        public List<string> GetBookmarksByPartialName(string name)
        {
            return (from WordOM.Bookmark b in _document.Bookmarks
                    where b.Name.StartsWith(name, true, CultureInfo.CurrentCulture)
                    select b.Name).ToList();
        }

        public void RenameBookmark(string name, string newName)
        {
            _document.Bookmarks[name].Range.Bookmarks.Add(newName);
            _document.Bookmarks[name].Delete();
        }

        public IDocument OpenFile(string serverRelativeUrl)
        {
            using (var clientContext = new ClientContext(Settings.Default.SharePointContextUrl))
            {
                string temp = Path.GetTempFileName();

                FileInformation fileInformation = File.OpenBinaryDirect(clientContext, serverRelativeUrl);
                Stream stream = fileInformation.Stream;

                using (FileStream output = System.IO.File.OpenWrite(temp))
                {
                    stream.CopyTo(output);
                }
                WordOM.Document d = _document.Application.Documents.Add(temp);
                var doc = new OfficeDocument(d);
                System.IO.File.Delete(temp);

                return doc;
            }
        }

        public IDocument OpenFile(string cacheName, string serverRelativeUrl)
        {
            if (serverRelativeUrl == null)
            {
                WordOM.Document nDoc = _document.Application.Documents.Add();
                var document = new OfficeDocument(nDoc);
                return document;
            }
            using (var clientContext = new ClientContext(Settings.Default.SharePointContextUrl))
            {
                string temp = Path.GetTempFileName();

                FileInformation fileInformation = File.OpenBinaryDirect(clientContext, serverRelativeUrl);
                Stream stream = fileInformation.Stream;

                using (FileStream output = System.IO.File.OpenWrite(temp))
                {
                    stream.CopyTo(output);
                }

                ObjectCache cache = MemoryCache.Default;
                cache.Add(cacheName, string.Empty, new CacheItemPolicy());

                WordOM.Document d = _document.Application.Documents.Add(temp);
                d.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.CacheName, cacheName);
                System.IO.File.Delete(temp);

                var document = new OfficeDocument(d);

                return document;
            }
        }

        public bool FileExistsInSharePoint(string serverRelativeUrl)
        {
            try
            {
                using (var clientContext = new ClientContext(Settings.Default.SharePointContextUrl))
                {
                    File.OpenBinaryDirect(clientContext, serverRelativeUrl);
                    return true;
                }
            }
            catch (Exception)
            {

                return false;
            }
        }

        public void TypeText(string text, string style = "")
        {
            _document.Activate();
            if (!String.IsNullOrEmpty(style))
            {
                WordOM.Style s = _document.Styles[style];
                //_document.Application.Selection.InsertStyleSeparator();
// ReSharper disable UseIndexedProperty
                _document.Application.Selection.set_Style(s);
// ReSharper restore UseIndexedProperty
            }

            _document.Application.Selection.TypeText(text);


            //if (!String.IsNullOrEmpty(style))
            //{
            //    //_document.Application.Selection.InsertStyleSeparator();

            //}
        }

        public void ResetListNumbering()
        {
            _document.Activate();
            if (_document.Application.Selection.Range.ListParagraphs.Count > 0)
            {
                WordOM.Paragraph listpara = _document.Application.Selection.Range.ListParagraphs[1];
                listpara.Range.ListFormat.ApplyListTemplate(listpara.Range.ListFormat.ListTemplate, false);
            }
        }

        public void InsertStyleSeperator()
        {
            _document.Activate();
            _document.Application.Selection.InsertStyleSeparator();
        }

        public void MoveToStartOfDocument()
        {
            _document.Activate();
            _document.Application.Selection.HomeKey(WordOM.WdUnits.wdStory);
        }

        public void MoveToEndOfDocument()
        {
            _document.Activate();
            _document.Application.Selection.EndKey(WordOM.WdUnits.wdStory);
        }

        public void ImportPDF(string path)
        {
            WordOM.Style s = _document.Styles["normal"];
// ReSharper disable UseIndexedProperty
            _document.Application.Selection.set_Style(s);
// ReSharper restore UseIndexedProperty
            _document.Application.Selection.InlineShapes.AddOLEObject("AcroExch.Document.7", path, false, false);
        }

        public void SelectTable(string name)
        {
            foreach (WordOM.Table table in _document.Tables)
            {
                if (String.Equals(name, table.Title, StringComparison.OrdinalIgnoreCase))
                {
                    table.Select();
                }
            }
        }

        public int TableRowOrColumnCount(bool isColumnCount, string name = "")
        {
            if (String.IsNullOrEmpty(name))
            {
                object t = _document.Application.Selection.Information[WordOM.WdInformation.wdWithInTable];
                var isT = false;
                Boolean.TryParse(t.ToString(), out isT);
                if (isT)
                {
                    return isColumnCount ? _document.Application.Selection.Tables[1].Columns.Count : _document.Application.Selection.Tables[1].Rows.Count;
                }
            }
            else
            {
                foreach (WordOM.Table table in _document.Tables)
                {
                    if (String.Equals(name, table.Title, StringComparison.OrdinalIgnoreCase))
                    {
                        return isColumnCount ? table.Columns.Count : table.Rows.Count;
                    }
                }
            }
            return -1;
        }

        public void InsertTableRowNonCellLoop(int insertRowPosition, string name = "")
        {
            if (String.IsNullOrEmpty(name))
            {
                var t = _document.Application.Selection.Information[WordOM.WdInformation.wdWithInTable];
                var isT = false;
                Boolean.TryParse(t.ToString(), out isT);
                if (isT)
                {
                    var table = _document.Application.Selection.Tables[1];
                    foreach (WordOM.Row r in table.Range.Rows)
                    {
                        if (r.Index == insertRowPosition)
                        {
                            r.Select();
                            _document.Application.Selection.InsertRowsBelow(1);
                            // _document.Application.Selection.Tables[1].Rows.Add(c.Row);
                            break;
                        }
                    }
                }
            }
            else
            {
                foreach (WordOM.Table table in _document.Tables)
                {
                    if (!String.Equals(name, table.Title, StringComparison.OrdinalIgnoreCase)) continue;
                    
                    foreach (var r in from WordOM.Row r in table.Range.Rows where r.Index == insertRowPosition select r)
                    {
                        r.Select();
                        _document.Application.Selection.InsertRowsBelow(1);
                        // _document.Application.Selection.Tables[1].Rows.Add(c.Row);
                        break;
                    }
                    break;
                }
            }
        }

        public void InsertTableRow(int insertRowPosition, string name = "")
        {
            if (String.IsNullOrEmpty(name))
            {
                object t = _document.Application.Selection.Information[WordOM.WdInformation.wdWithInTable];
                bool isT = false;
                Boolean.TryParse(t.ToString(), out isT);
                if (isT)
                {
                    WordOM.Table table = _document.Application.Selection.Tables[1];
                    foreach (WordOM.Cell c in table.Range.Cells)
                    {
                        if (c.RowIndex == insertRowPosition)
                        {
                            c.Select();
                            _document.Application.Selection.InsertRowsBelow(1);
                            // _document.Application.Selection.Tables[1].Rows.Add(c.Row);
                            break;
                        }
                    }
                }
            }
            else
            {
                foreach (WordOM.Table table in _document.Tables)
                {
                    if (String.Equals(name, table.Title, StringComparison.OrdinalIgnoreCase))
                    {
                        foreach (WordOM.Cell c in table.Range.Cells)
                        {
                            if (c.RowIndex == insertRowPosition)
                            {
                                //_document.Application.Selection.Tables[1].Rows.Add(c.Row);
                                c.Select();
                                _document.Application.Selection.InsertRowsBelow(1);
                                break;
                            }
                        }
                    }
                }
            }
        }

        public void InsertTableCell(int insertBeforeColumnIndex, string cellText, string tableName = "")
        {
            if (String.IsNullOrEmpty(tableName))
            {
                object t = _document.Application.Selection.Information[WordOM.WdInformation.wdWithInTable];
                var isT = false;
                Boolean.TryParse(t.ToString(), out isT);
                if (isT)
                {
                    WordOM.Table table = _document.Application.Selection.Tables[1];
                    table.Select();
                    _document.Application.Selection.InsertCells(WordOM.WdInsertCells.wdInsertCellsEntireColumn);
                }
            }
            else
            {
                foreach (WordOM.Table table in _document.Tables)
                {
                    if (String.Equals(tableName, table.Title, StringComparison.OrdinalIgnoreCase))
                    {
                        WordOM.Column col = table.Columns[insertBeforeColumnIndex];

                        col.Select();
                        _document.Application.Selection.InsertColumnsRight();
                        table.Columns[table.Columns.Count].Cells[1].Range.Text = cellText;

                        //table.Rows[1].Cells[insertBeforeColumnIndex].Select();
                        //_document.Application.Selection.InsertCells();
                        //table.Cell(1, table.Columns.Count).Range.Text = cellText;

                        // table.Columns.Add(table.Columns[insertBeforeColumnIndex]).Cells[insertBeforeColumnIndex + 1].Range.Text = cellText;
                        //    table.Rows[1].Cells[2].Select();
                        //    _document.Application.Selection.InsertColumnsRight();
                        ////    _document.Application.Selection.InsertCells();
                        //    table.Cell(1, table.Columns.Count).Range.Text = cellText;
                    }
                }
            }
        }

        public void RenameTable(string tableName, string newName)
        {
            foreach (WordOM.Table table in _document.Tables)
            {
                if (String.Equals(tableName, table.Title, StringComparison.OrdinalIgnoreCase))
                {
                    table.Title = newName;
                }
            }
        }

        public void PopulateTableCell(int row, int column, string value, string name = "")
        {
            if (String.IsNullOrEmpty(name))
            {
                object t = _document.Application.Selection.Information[WordOM.WdInformation.wdWithInTable];
                var isT = false;
                Boolean.TryParse(t.ToString(), out isT);
                if (isT)
                {
                    //var cell = _document.Application.Selection.Tables[1].Cell(row, column);
                    //cell.Range.Text = value;

                    //_document.Application.Selection.Tables[1].Columns[1].Cells

                    WordOM.Table table = _document.Application.Selection.Tables[1];
                    foreach (WordOM.Cell c in table.Range.Cells)
                    {
                        if (c.ColumnIndex == column && c.RowIndex == row)
                        {
                            //foreach (WordOM.Row r in c.Row)
                            //{
                            //    if (r.Index == row)
                            //    {
                            c.Range.Text = value;
                            //    }
                            //}
                        }
                    }

                    //for (int rI = 1; rI < table.Rows.Count; rI ++)
                    //{
                    //    if (rI == row)
                    //    {

                    //    }
                    //}
                }
            }
            else
            {
                foreach (WordOM.Table table in _document.Tables)
                {
                    if (String.Equals(name, table.Title, StringComparison.OrdinalIgnoreCase))
                    {
                        WordOM.Cell cell = table.Cell(row, column);
                        cell.Range.Text = value;
                    }
                }
            }
        }

        public string ReadTableCell(int row, int column, string tableName = "")
        {
            if (String.IsNullOrEmpty(tableName))
            {
                
            }
            else
            {
                foreach (WordOM.Table table in _document.Tables)
                {
                    if (String.Equals(tableName, table.Title, StringComparison.OrdinalIgnoreCase))
                    {
                        var cell = table.Cell(row, column);
                        return cell.Range.Text;
                    }
                }
            }

            return string.Empty;
        }

        public bool IsRangeReadOnly()
        {
            return _document.ProtectionType != WordOM.WdProtectionType.wdNoProtection;
        }

        public int TurnOffProtection(string password)
        {
            WordOM.WdProtectionType oldType = _document.ProtectionType;
            _document.Unprotect(password);
            return (int) oldType;
        }

        public void TurnOnProtection(int type, string password)
        {
            if (_document == null || _document.ProtectionType != WordOM.WdProtectionType.wdNoProtection)
                return;

            _document.Protect((WordOM.WdProtectionType) type, true, password);
            ToggleDocumentEditableShade();
        }

        public void InsertParagraphBreak()
        {
            _document.Application.Selection.TypeParagraph();
        }

        public void InsertBackspace()
        {
            _document.Application.Selection.TypeBackspace();
        }

        public void RemoveHeader()
        {
            foreach (WordOM.Section sec in _document.Sections)
            {
                foreach (WordOM.HeaderFooter header in sec.Headers)
                {
                    header.Range.Delete();
                }
            }
            //_document.Sections[1].Headers[
            //    WordOM.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Delete();            
        }

        public void RemoveFooter()
        {
            foreach (WordOM.Section sec in _document.Sections)
            {
                foreach (WordOM.HeaderFooter footer in sec.Footers)
                {
                    footer.Range.Delete();
                }
            }

            //_document.Sections[1].Footers[WordOM.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Delete();
        }

        public void SwitchScreenUpdating(bool show)
        {
            _document.Application.ScreenUpdating = show;
        }

        public void DeleteImage(string imageAltText)
        {
            //for (var i = 1; i <= shapesCount; i++) //VSTO arrays start at 1
            //{
            //    var j = i;
            //    var currentImage =
            //        _document.Sections[1].Headers[
            //            WordOM.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[j];
            //    if (imageAltText.Contains(currentImage.AlternativeText))
            //    {
            //    }
            //}

            foreach (WordOM.Shape s in _document.Shapes)
            {
                if (String.Equals(s.AlternativeText, imageAltText, StringComparison.InvariantCultureIgnoreCase))
                    s.Delete();
            }
        }

        public void Activate()
        {
            _document.Activate();
        }

        public void SetMargins(float leftMargin, float rightMargin)
        {
            _document.Activate();
            //_document.Application.Selection.ParagraphFormat.LeftIndent = margin;
            _document.Application.Selection.PageSetup.LeftMargin = leftMargin;
            _document.Application.Selection.PageSetup.RightMargin = rightMargin;
            _document.Application.Selection.ParagraphFormat.SpaceBeforeAuto = 0;
            _document.Application.Selection.ParagraphFormat.SpaceAfterAuto = 0;
        }

        public void UnlinkDocumentFooterAndHeader()
        {
            _document.Activate();
            if (_document.Application.Selection.HeaderFooter != null)
                _document.Application.Selection.HeaderFooter.LinkToPrevious = false;
        }

        public void SetTopMargins(float margin)
        {
            _document.Activate();
            _document.Application.Selection.PageSetup.TopMargin = margin;
            _document.Application.Selection.ParagraphFormat.SpaceBeforeAuto = 0;
            _document.Application.Selection.ParagraphFormat.SpaceAfterAuto = 0;
        }
        
        public void SendEmail(string subject, string body, string recipients)
        {
            var app = new OutlookOM.Application();
            OutlookOM.MailItem email = app.CreateItem(OutlookOM.OlItemType.olMailItem);
            var rs = recipients.Split(';');
            foreach (var r in rs)
            {
                email.Recipients.Add(r);
            }
            email.Subject = subject;
            email.HTMLBody = body;
            email.Send();
        }

        public void DisplayMessage(string text, string caption)
        {
            throw new NotImplementedException();
        }

        private string DownloadLocalImage(string serverRelativeUrl)
        {
            using (var clientContext = new ClientContext(Settings.Default.SharePointContextUrl))
            {
                string temp = Path.GetTempFileName();

                var fileInformation = File.OpenBinaryDirect(clientContext, ("/" + serverRelativeUrl.ToLower().Replace(clientContext.Url.ToLower(), string.Empty)).Replace("//","/")); // :( yep
                var stream = fileInformation.Stream;

                using (var output = System.IO.File.OpenWrite(temp))
                {
                    stream.CopyTo(output);
                }
                return temp;
            }
        }

        public void SetMargins(float margin)
        {
            throw new NotImplementedException();
        }

        public void ToggleDocumentEditableShade()
        {
            _document.ActiveWindow.View.ShadeEditableRanges = 0;
        }
    }
}