using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

        private void MeasureTypeSearchMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var tableName = TableNameMeasureType;
                var titleRow = TitleRowMeasureType;
                var resultColumn = ResultColumnMeasureType;

                var id = "" + _cells.GetCell("B4").Value;
                var description = "" + _cells.GetCell("B5").Value;

                _cells.ClearTableRange(tableName);

                var i = titleRow + 1;

                List<MeasureType> itemList;
                if (!string.IsNullOrWhiteSpace(id))
                    itemList = MeasureType.Read(MyUtilities.ToInteger(id));
                else if (!string.IsNullOrWhiteSpace(description))
                    itemList = MeasureType.Read(description);
                else
                    itemList = MeasureType.Read();

                foreach (var item in itemList)
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                        _cells.GetCell(1, i).Value = "" + item.Id;
                        _cells.GetCell(2, i).Value = "" + item.Description;

                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Searched;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:MeasureTypeSearchMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:MeasureTypeSearchMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
        private void MeasureTypeSearchEachMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var tableName = TableNameMeasureType;
                var titleRow = TitleRowMeasureType;
                var resultColumn = ResultColumnMeasureType;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(1, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                        var id = MyUtilities.ToInteger("" + _cells.GetCell(1, i).Value);

                        var item = MeasureType.ReadFirst(id);

                        if(item == null)
                            throw new ArgumentException(Resources.Error_ItemNotFound, nameof(item));
                        _cells.GetCell(1, i).Value = item.Id;
                        _cells.GetCell(2, i).Value = item.Description;

                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Searched;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:MeasureTypeSearchEachMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:MeasureTypeSearchEachMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
        private void MeasureTypeUpdateMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var tableName = TableNameMeasureType;
                var titleRow = TitleRowMeasureType;
                var resultColumn = ResultColumnMeasureType;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(2, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                        var id = "" + _cells.GetCell(1, i).Value;
                        var description = "" + _cells.GetCell(2, i).Value;
                        var item = new MeasureType();
                        if(!string.IsNullOrWhiteSpace(id))
                            item.Id = MyUtilities.ToInteger(id);
                        if (string.IsNullOrWhiteSpace(description))
                            throw new ArgumentNullException(nameof(description), Resources.Error_NullValue);
                        item.Description = description;

                        var result = MeasureType.Create(item);
                        if (result.Message.Equals(Resources.Results_Failed))
                        {
                            var message = Resources.Error_ErrorUppercase;
                            if (result.Errors != null)
                                message = result.Errors.Aggregate(message, (current, e) => current + (" " + e));

                            throw new Exception(message);
                        }
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Success;

                        try
                        {
                            if (!string.IsNullOrWhiteSpace(id)) continue;
                            var newItem = MeasureType.ReadFirst(description);
                            _cells.GetCell(1, i).Value = newItem.Id;
                        }
                        catch (Exception)
                        {
                            _cells.GetCell(resultColumn, i).Style = StyleConstants.Warning;
                            _cells.GetCell(resultColumn, i).Value = Resources.Results_Success;
                        }
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:MeasureTypeUpdateMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:MeasureTypeUpdateMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void MeasureTypeDeleteMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var tableName = TableNameMeasureType;
                var titleRow = TitleRowMeasureType;
                var resultColumn = ResultColumnMeasureType;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(1, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                        
                        var id = "" + _cells.GetCell(1, i).Value;
                        if(string.IsNullOrWhiteSpace(id))
                            throw new ArgumentNullException(nameof(id), Resources.Error_NullValue);

                        MeasureType.Delete(MyUtilities.ToInteger(id));

                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Deleted;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:MeasureTypeDeleteMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:MeasureTypeDeleteMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
    }
}
