using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.Word.Helpers;
using WordOM = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using File = Microsoft.SharePoint.Client.File;
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
            var find =  _document.Application.Selection.Find;

            var r = _document.Application.Selection.Move();
            

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
            var contentControls = _document.SelectContentControlsByTag(tag);
            if (contentControls == null) return;
            foreach (WordOM.ContentControl contentControl in contentControls)
            {

                if (contentControl.Type == WordOM.WdContentControlType.wdContentControlCheckBox)
                {
                    if (!String.IsNullOrEmpty(value)) contentControl.Checked = Boolean.Parse(value);
                }
                else
                {
                    value = value.Replace("\r\n", ((char)11).ToString());

                    contentControl.Range.Text = value;
                }

                if (value != null)
                {
                    contentControl.SetPlaceholderText(null, null, value);

                }

            }
        }

        public void PopulateControlInBookmark(string bookmarkName, string tag, string value)
        {
            if (!_document.Bookmarks.Exists(bookmarkName)) return;

            var contentControls = _document.Bookmarks[bookmarkName].Range.ContentControls;
            if (contentControls == null) return;
            foreach (WordOM.ContentControl contentControl in contentControls)
            {
                contentControl.Range.Text = value;
            }
        }

        public void DeleteControl(string tag)
        {
            var contentControls = _document.SelectContentControlsByTag(tag);
            if (contentControls == null) return;
            foreach (WordOM.ContentControl contentControl in contentControls)
            {
                contentControl.Delete(true);
            }
        }

        public string ReadContentControlValue(string tag)
        {
            var contentControls = _document.SelectContentControlsByTag(tag);
            if (contentControls.Count < 1)
                return string.Empty;


            if (contentControls[1].Type == WordOM.WdContentControlType.wdContentControlCheckBox)
            {
                return (contentControls.Count < 1) ? String.Empty : contentControls[1].Checked.ToString();
            }
            else
            {
                return (contentControls.Count < 1) ? String.Empty : contentControls[1].Range.Text;
            }
        }
        
        public void RenameControl(string tag, string newTag)
        {
            var contentControls = _document.SelectContentControlsByTag(tag);
            if (contentControls == null) return;
            foreach (WordOM.ContentControl contentControl in contentControls)
            {
                contentControl.Tag = newTag;
            }
        }

        public void RenameControlInBookmark(string tag, string newTag, string bookmarkName)
        {
            if (!_document.Bookmarks.Exists(bookmarkName)) return;

            var contentControls = _document.Bookmarks[bookmarkName].Range.ContentControls;
            if (contentControls == null) return;
            foreach (WordOM.ContentControl contentControl in contentControls)
            {
                contentControl.Tag = newTag;
            }
        }

        public void DeletePage(int pageNumber = -1)
        {
            if (pageNumber > 0)
                _document.Application.Selection.GoTo(WordOM.WdGoToItem.wdGoToPage, WordOM.WdGoToDirection.wdGoToAbsolute,pageNumber);

            _document.Bookmarks[@"\Page"].Range.Delete();
        }

        public void DisplayMessage(string text, string caption)
        {
            throw new NotImplementedException();
        }

        public void CloseInformationPanel(bool delayClose = false)
        {
            try
            {
                if (delayClose)
                {
                    Task.Factory.StartNew(() => System.Threading.Thread.Sleep(2000)).ContinueWith(task => CloseInformationPanel(false));
                }
                else
                {
                    _document.Application.DisplayDocumentInformationPanel = false;
                    _document.ActiveWindow.View.Zoom.Percentage = 100; //todo: THIS DOES NOT BELONG HERE, EXTRACT TO ITS OWN METHOD!
                }
            }
            catch (Exception ex)
            {
                string b = "";

            }
          
        }
        
        private string DownloadLocalImage(string serverRelativeUrl)
        {
            using (var clientContext = new ClientContext(Properties.Settings.Default.SharePointContextUrl))
            {
                var temp = System.IO.Path.GetTempFileName();

                FileInformation fileInformation = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, "/" +  serverRelativeUrl.Replace(clientContext.Url,string.Empty));
                var stream = fileInformation.Stream;

                using (FileStream output = System.IO.File.OpenWrite(temp))
                {
                    stream.CopyTo(output);
                }


                return temp;
            }
        }

        //this method changes the cover page and company logo images that are selected on the wizard
        public void ChangeDocumentImages(IList<string> imageAltText, IList<string> imageUrls)
        {
            var localImageUrls = new Dictionary<string, string>();
            var shapesCount =
                _document.Sections[1].Headers[
                    WordOM.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.Count;
            var deleteableShapes = new List<WordOM.Shape>();

            for (var i = 1; i <= shapesCount; i++) //VSTO arrays start at 1
            {
                var j = i;
                var currentImage =
                   _document.Sections[1].Headers[
                        WordOM.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes[j];
                if (imageAltText.Contains(currentImage.AlternativeText))
                {
                    var imageIndex = imageAltText.IndexOf(currentImage.AlternativeText);

                    
                    if (String.IsNullOrEmpty(imageUrls[imageIndex])) continue;

                    if (!localImageUrls.ContainsKey(imageAltText[imageIndex]))
                    {
                        var localUrl = DownloadLocalImage(imageUrls[imageIndex]);
                        localImageUrls.Add(imageAltText[imageIndex], localUrl);
                    }
                    var key = imageAltText[imageIndex];
                    var newImage = _document.Sections[1].Footers[WordOM.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Shapes.AddPicture(
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

            var cBookmark = _document.Bookmarks[name];
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

            var bs = _document.Application.Selection.Bookmarks.Add(newName + "start");
            bs.Range.InsertAfter(" ");
            _document.Application.Selection.Move(WordOM.WdUnits.wdCharacter, 1);
            var be = _document.Application.Selection.Bookmarks.Add(newName + "end");
            _document.Range(bs.Start, bs.End).Paste();

            var contentRange = _document.Range(bs.Start, be.End);
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

        public void InsertFile(string serverRelativeUrl)
        {
            //enhancement - we could not delete the file at the end of this method
            //enhancement - if we did not delete this file we could compqare its created date agasint the modified date in sharepoint and only take a copy if updates have occured.
            using (var clientContext = new ClientContext(Properties.Settings.Default.SharePointContextUrl))
            {
                var temp = System.IO.Path.GetTempFileName();

                FileInformation fileInformation = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, serverRelativeUrl);
                var stream = fileInformation.Stream;

                using (FileStream output = System.IO.File.OpenWrite(temp))
                {
                    stream.CopyTo(output);
                }

                _document.Application.Selection.InsertFile(temp, "", false, false);

                System.IO.File.Delete(temp);
            }

        }

        public bool MoveCursorToStartOfBookmark(string name)
        {
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
            _document.Application.Selection.MoveUp(WordOM.WdUnits.wdLine, count);
        }

        public void MoveCursorRight(int count)
        {
            _document.Application.Selection.MoveRight(WordOM.WdUnits.wdCharacter, count);
        }

        public void MoveCursorLeft(int count)
        {
            _document.Application.Selection.MoveLeft(WordOM.WdUnits.wdCharacter, count);
        }

        public void InsertPageBreak()
        {
            _document.Application.Selection.InsertBreak(WordOM.WdBreakType.wdSectionBreakNextPage);
        }

        


        public void InsertSectionBreak()
        {
            _document.Application.Selection.InsertBreak(WordOM.WdBreakType.wdSectionBreakContinuous);
        }

        public void SetMargins(float margin)
        {
            throw new NotImplementedException();
        }


        public void ChangePageOrientToPortrait()
        {
            if (_document.Application.Selection.get_Information(WordOM.WdInformation.wdWithInTable))
            {
                _document.Application.Selection.Tables[1].Select();
                _document.Application.Selection.Collapse(WordOM.WdCollapseDirection.wdCollapseEnd);
            }

            //if cursor didnt move out of table dont contintue else error will generate.
            if (_document.Application.Selection.get_Information(WordOM.WdInformation.wdWithInTable))
                return;

            _document.Application.Selection.PageSetup.Orientation = WordOM.WdOrientation.wdOrientPortrait;

        }

        public void ChangePageOrientToLandscape()
        {
            _document.Application.Selection.PageSetup.Orientation = WordOM.WdOrientation.wdOrientLandscape;
        }

        public void MoveCursorPastStartOfBookmark(string name, int moveCharacterCount)
        {
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
            using (var clientContext = new ClientContext(Properties.Settings.Default.SharePointContextUrl))
            {
                var temp = System.IO.Path.GetTempFileName();

                FileInformation fileInformation = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, serverRelativeUrl);
                var stream = fileInformation.Stream;

                using (FileStream output = System.IO.File.OpenWrite(temp))
                {
                    stream.CopyTo(output);
                }
                var d =_document.Application.Documents.Add(temp);
                var doc = new OfficeDocument(d);
                System.IO.File.Delete(temp);

                return doc;
            }
        }

        public void OpenFile(string cacheName, string serverRelativeUrl)
        {
            if (serverRelativeUrl == null)
            {
                _document.Application.Documents.Add();
                return;                
            }
            using (var clientContext = new ClientContext(Properties.Settings.Default.SharePointContextUrl))
            {
                var temp = System.IO.Path.GetTempFileName();

                FileInformation fileInformation = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, serverRelativeUrl);
                var stream = fileInformation.Stream;

                using (FileStream output = System.IO.File.OpenWrite(temp))
                {
                    stream.CopyTo(output);
                }

                var d = _document.Application.Documents.Add(temp);
                d.UpdateOrCreatePropertyValue(BusinessLogic.Helpers.Constants.WordDocumentProperties.CacheName, cacheName);
                System.IO.File.Delete(temp);
            }
        }
        
        public void TypeText(string text, string style = "")
        {
            if (!String.IsNullOrEmpty(style))
            {

                
                var s = _document.Styles[style];
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
            if (_document.Application.Selection.Range.ListParagraphs.Count > 0)
            {
                var listpara = _document.Application.Selection.Range.ListParagraphs[1];
                listpara.Range.ListFormat.ApplyListTemplate(listpara.Range.ListFormat.ListTemplate, false);    
            }
        }

        public void InsertStyleSeperator()
        {
            _document.Application.Selection.InsertStyleSeparator();
        }

        public void MoveToStartOfDocument()
        {
            _document.Application.Selection.HomeKey(WordOM.WdUnits.wdStory);
        }

        public void MoveToEndOfDocument()
        {
            _document.Application.Selection.EndKey(WordOM.WdUnits.wdStory);
        }

        public void ImportPDF(string path)
        {
            var s = _document.Styles["normal"]; 
// ReSharper disable UseIndexedProperty
            _document.Application.Selection.set_Style(s);
// ReSharper restore UseIndexedProperty
            _document.Application.Selection.InlineShapes.AddOLEObject("AcroExch.Document.7",path,false,false);
           
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
                var t = _document.Application.Selection.Information[WordOM.WdInformation.wdWithInTable];
                bool isT = false;
                Boolean.TryParse(t.ToString(), out isT);
                if (isT)
                {
                    if (isColumnCount)
                    {
                        return _document.Application.Selection.Tables[1].Columns.Count;
                    }
                    return _document.Application.Selection.Tables[1].Rows.Count;
                }
            }
            else
            {
                foreach (WordOM.Table table in _document.Tables)
                {
                    if (String.Equals(name, table.Title, StringComparison.OrdinalIgnoreCase))
                    {
                        if (isColumnCount)
                        {
                            return table.Columns.Count;
                        }
                        return table.Rows.Count;

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
                bool isT = false;
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
                    if (String.Equals(name, table.Title, StringComparison.OrdinalIgnoreCase))
                    {
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
            }
        }

        public void InsertTableRow(int insertRowPosition, string name = "")
        {
            if (String.IsNullOrEmpty(name))
            {
                var t = _document.Application.Selection.Information[WordOM.WdInformation.wdWithInTable];
                bool isT = false;
                Boolean.TryParse(t.ToString(), out isT);
                if (isT)
                {
                    var table = _document.Application.Selection.Tables[1];
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
                var t = _document.Application.Selection.Information[WordOM.WdInformation.wdWithInTable];
                bool isT = false;
                Boolean.TryParse(t.ToString(), out isT);
                if (isT)
                {
                    var table = _document.Application.Selection.Tables[1];
                    table.Select();
                    _document.Application.Selection.InsertCells(Microsoft.Office.Interop.Word.WdInsertCells.wdInsertCellsEntireColumn);
                }
            }
            else
            {
                foreach (WordOM.Table table in _document.Tables)
                {
                    if (String.Equals(tableName, table.Title, StringComparison.OrdinalIgnoreCase))
                    {
                       var col =  table.Columns[insertBeforeColumnIndex];
                        
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
                var t = _document.Application.Selection.Information[WordOM.WdInformation.wdWithInTable];
                bool isT = false;
                Boolean.TryParse(t.ToString(), out isT);
                if (isT)
                {
                    //var cell = _document.Application.Selection.Tables[1].Cell(row, column);
                    //cell.Range.Text = value;

                    //_document.Application.Selection.Tables[1].Columns[1].Cells

                    var table = _document.Application.Selection.Tables[1];
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
                        var cell = table.Cell(row, column);
                        cell.Range.Text = value;
                    }
                }
            }
        }

        public string ReadTableCell(int row, int column, string tableName = "")
        {
            if (String.IsNullOrEmpty(tableName))
            {
                throw new NotImplementedException();
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
            if (_document == null || _document.ProtectionType == WordOM.WdProtectionType.wdNoProtection)
                return (int)WordOM.WdProtectionType.wdNoProtection;

            WordOM.WdProtectionType oldType = _document.ProtectionType;
            _document.Unprotect(password);
            return (int) oldType;
        }

        public void TurnOnProtection(int type, string password)
        {
            if (_document == null || _document.ProtectionType == WordOM.WdProtectionType.wdNoProtection)
                return;

            _document.Protect((WordOM.WdProtectionType) type, true, password);
            ToggleDocumentEditableShade();
        }

        public void ToggleDocumentEditableShade()
        {
            _document.ActiveWindow.View.ShadeEditableRanges = 0;
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
            //_document.Application.Selection.ParagraphFormat.LeftIndent = margin;
            _document.Application.Selection.PageSetup.LeftMargin = leftMargin;
            _document.Application.Selection.PageSetup.RightMargin = rightMargin;
            _document.Application.Selection.ParagraphFormat.SpaceBeforeAuto = 0;
            _document.Application.Selection.ParagraphFormat.SpaceAfterAuto = 0;
        }

        public void UnlinkDocumentFooterAndHeader()
        {
            if (_document.Application.Selection.HeaderFooter!=null)
            _document.Application.Selection.HeaderFooter.LinkToPrevious = false;
        }

        public void SetTopMargins(float margin)
        {
            _document.Application.Selection.PageSetup.TopMargin = margin;
            _document.Application.Selection.ParagraphFormat.SpaceBeforeAuto = 0;
            _document.Application.Selection.ParagraphFormat.SpaceAfterAuto = 0;
        }

        

        public void SendEmail(string body)
        {
            var app = new OutlookOM.Application();
            OutlookOM.MailItem email = app.CreateItem(OutlookOM.OlItemType.olMailItem);
            email.Recipients.Add("maxwell.jiang@oamps.com.au");
            email.Subject = "Message subject";
            //email.Body = body;
            email.HTMLBody = body;
            email.Send();
        }
    }
}
