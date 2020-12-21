using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using LogsheetDatamodelLibrary;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelAdmin
{
    public partial class RibbonLsdm
    {

        private void ValidationItemsSearchMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameValidItems;
                const int titleRow = TitleRowValidItems;
                const int resultColumn = ResultColumnValidItems;

                var keyword = "" + _cells.GetCell("B4").Value;

                _cells.ClearTableRange(tableName);

                var i = titleRow + 1;

                List<ValidationItem> itemList = ValidationItem.Read(keyword, keyword);

                foreach (var item in itemList)
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                        _cells.GetCell(1, i).Value = "" + item.SourceName;
                        _cells.GetCell(2, i).Value = "" + item.Id;
                        _cells.GetCell(3, i).Value = "" + item.Description;
                        _cells.GetCell(4, i).Value = "" + item.SourceTable;
                        _cells.GetCell(5, i).Value = "" + item.SourceColumn;
                        _cells.GetCell(6, i).Value = "" + item.Sortable;
                        _cells.GetCell(7, i).Value = "" + item.DistinctFilter;

                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Searched;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ValidationItemsSearchMethod()", ex.Message);
                    }
                    finally
                    {
                        _cells.GetCell(2, i).Select();
                        i++;
                    }
                }

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Resources.Error_ErrorFound, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse.cs:ValidationItemsSearchMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
        private void ValidationItemsSearchEachMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameValidItems;
                const int titleRow = TitleRowValidItems;
                const int resultColumn = ResultColumnValidItems;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(1, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                        var id = MyUtilities.ToInteger("" + _cells.GetCell(2, i).Value);

                        var item = ValidationItem.Read(id);

                        if (item == null)
                            throw new ArgumentException(Resources.Error_ItemNotFound, nameof(item));
                        _cells.GetCell(1, i).Value = "" + item.SourceName();
                        _cells.GetCell(2, i).Value = "" + item.Id;
                        _cells.GetCell(3, i).Value = "" + item.Description;
                        _cells.GetCell(4, i).Value = "" + item.SourceTable;
                        _cells.GetCell(5, i).Value = "" + item.SourceColumn;
                        _cells.GetCell(6, i).Value = "" + item.Sortable;
                        _cells.GetCell(7, i).Value = "" + item.DistinctFilter;

                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Searched;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ValidationItemsSearchEachMethod()", ex.Message);
                    }
                    finally
                    {
                        _cells.GetCell(2, i).Select();
                        i++;
                    }
                }

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Resources.Error_ErrorFound, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse.cs:ValidationItemsSearchEachMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
        private void ValidationItemsUpdateMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameValidItems;
                const int titleRow = TitleRowValidItems;
                const int resultColumn = ResultColumnValidItems;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(1, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                        var sourceName = "" + _cells.GetCell(1, i).Value;
                        var id = MyUtilities.ToIntegerNull("" + _cells.GetCell(2, i).Value);
                        var description = "" + _cells.GetCell(3, i).Value;
                        var sourceTable = "" + _cells.GetCell(4, i).Value;
                        var sourceColumn = "" + _cells.GetCell(5, i).Value;
                        var sortable = MyUtilities.IsTrue("" + _cells.GetCell(6, i).Value);
                        var distinctFilter = MyUtilities.IsTrue("" + _cells.GetCell(7, i).Value);

                        var item = new ValidationItem(sourceName, id, description, sourceTable, sourceColumn, sortable, distinctFilter);


                        var result = ValidationItem.Create(item);
                        if (result.Message.Equals(Resources.Results_Failed))
                        {
                            var message = Resources.Error_ErrorUppercase;
                            if (result.Errors != null)
                                message = result.Errors.Aggregate(message, (current, e) => current + (" " + e));

                            throw new Exception(message);
                        }
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Success;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ValidationItemsUpdateMethod()", ex.Message);
                    }
                    finally
                    {
                        _cells.GetCell(2, i).Select();
                        i++;
                    }
                }

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Resources.Error_ErrorFound, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse.cs:ValidationItemsUpdateMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void ValidationItemsDeleteMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameValidItems;
                const int titleRow = TitleRowValidItems;
                const int resultColumn = ResultColumnValidItems;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(2, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                        var id = "" + _cells.GetCell(2, i).Value;
                        if (string.IsNullOrWhiteSpace(id))
                            throw new ArgumentNullException(nameof(id), Resources.Error_NullValue);

                        ValidationItem.Delete(MyUtilities.ToIntegerNull(id));

                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Deleted;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ValidationItemsDeleteMethod()", ex.Message);
                    }
                    finally
                    {
                        _cells.GetCell(2, i).Select();
                        i++;
                    }
                }

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Resources.Error_ErrorFound, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse.cs:ValidationItemsDeleteMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
    }
}
