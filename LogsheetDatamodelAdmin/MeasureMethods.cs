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

        private void MeasureSearchMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var tableName = TableNameMeasure;
                var titleRow = TitleRowMeasure;
                var resultColumn = ResultColumnMeasure;

                var id = "" + _cells.GetCell("B4").Value;
                var code = "" + _cells.GetCell("B5").Value;
                var keyword = "" + _cells.GetCell("D4").Value;

                _cells.ClearTableRange(tableName);

                var i = titleRow + 1;

                List<Measure> itemList;
                if (!string.IsNullOrWhiteSpace(id))
                    itemList = Measure.Read(MyUtilities.ToInteger(id));
                else if (!string.IsNullOrWhiteSpace(code))
                    itemList = Measure.Read(code);
                else
                    itemList = Measure.Read(code, keyword);

                foreach (var item in itemList)
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                        _cells.GetCell(1, i).Value = "" + item.Id;
                        _cells.GetCell(2, i).Value = "" + item.Code;
                        _cells.GetCell(3, i).Value = "" + item.Name;
                        _cells.GetCell(4, i).Value = "" + item.Description;
                        _cells.GetCell(5, i).Value = "" + item.Units;
                        _cells.GetCell(6, i).Value = "" + item.ActiveStatus;
                        _cells.GetCell(7, i).Value = "" + item.MeasureTypeId;

                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Searched;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:MeasureSearchMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:MeasureSearchMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
        private void MeasureSearchEachMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var tableName = TableNameMeasure;
                var titleRow = TitleRowMeasure;
                var resultColumn = ResultColumnMeasure;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(1, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                        var id = MyUtilities.ToInteger("" + _cells.GetCell(1, i).Value);

                        var item = Measure.ReadFirst(id);

                        if (item == null)
                            throw new ArgumentException(Resources.Error_ItemNotFound, nameof(item));
                        _cells.GetCell(1, i).Value = "" + item.Id;
                        _cells.GetCell(2, i).Value = "" + item.Code;
                        _cells.GetCell(3, i).Value = "" + item.Name;
                        _cells.GetCell(4, i).Value = "" + item.Description;
                        _cells.GetCell(5, i).Value = "" + item.Units;
                        _cells.GetCell(6, i).Value = "" + item.ActiveStatus;
                        _cells.GetCell(7, i).Value = "" + item.MeasureTypeId();

                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Searched;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:MeasureSearchEachMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:MeasureSearchEachMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
        private void MeasureUpdateMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var tableName = TableNameMeasure;
                var titleRow = TitleRowMeasure;
                var resultColumn = ResultColumnMeasure;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(2, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                        var id = "" + _cells.GetCell(1, i).Value;
                        var code = "" + _cells.GetCell(2, i).Value;
                        var typeId = MyUtilities.GetCodeKey("" + _cells.GetCell(7, i).Value);

                        var item = new Measure();

                        if (!string.IsNullOrWhiteSpace(id))
                            item.Id = MyUtilities.ToInteger(id);
                        if (string.IsNullOrWhiteSpace(code))
                            throw new ArgumentNullException(nameof(code), Resources.Error_NullValue);

                        item.Code = code;
                        item.Name = "" + _cells.GetCell(3, i).Value; 
                        item.Description = "" + _cells.GetCell(4, i).Value; 
                        item.Units = "" + _cells.GetCell(5, i).Value; 
                        item.ActiveStatus = MyUtilities.IsTrue("" + _cells.GetCell(6, i).Value);

                        if (!string.IsNullOrWhiteSpace(typeId))
                            item.MeasureTypeId = MyUtilities.ToInteger(typeId);

                        var result = Measure.Create(item);
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
                            var newItem = Measure.ReadFirst(MyUtilities.ToInteger(typeId), code);
                            _cells.GetCell(1, i).Value = newItem.Id;
                        }
                        catch (Exception ex)
                        {
                            _cells.GetCell(resultColumn, i).Style = StyleConstants.Warning;
                            _cells.GetCell(resultColumn, i).Value = Resources.Results_Success;
                            Debugger.LogError("RibbonEllipse.cs:MeasureUpdateMethod()", ex.Message);
                        }
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:MeasureUpdateMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:MeasureUpdateMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void MeasureDeleteMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var tableName = TableNameMeasure;
                var titleRow = TitleRowMeasure;
                var resultColumn = ResultColumnMeasure;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(2, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                        var id = "" + _cells.GetCell(1, i).Value;
                        if (string.IsNullOrWhiteSpace(id))
                            throw new ArgumentNullException(nameof(id), Resources.Error_NullValue);

                        Measure.Delete(MyUtilities.ToInteger(id));

                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Deleted;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:MeasureDeleteMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:MeasureDeleteMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
    }
}
