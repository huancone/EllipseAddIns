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
        private void AttributeSearchMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameAttribute;
                const int titleRow = TitleRowAttribute;
                const int resultColumn = ResultColumnAttribute;

                var modelId = "" + _cells.GetCell("B4").Value;
                var hideInactive = MyUtilities.IsTrue("" + _cells.GetCell("B5").Value);
                _cells.ClearTableRange(tableName);

                List<ModelAttribute> itemList = ModelAttribute.Read(modelId, hideInactive);

                var i = titleRow + 1;
                foreach (var item in itemList)
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                        _cells.GetCell(1, i).Value = "" + item.ModelId;
                        _cells.GetCell(2, i).Value = "" + item.Id;
                        _cells.GetCell(3, i).Value = "" + item.Description;
                        _cells.GetCell(4, i).Value = "" + item.DataType;
                        _cells.GetCell(5, i).Value = "" + item.SheetIndex;
                        _cells.GetCell(6, i).Value = "" + item.MaxLength;
                        _cells.GetCell(7, i).Value = "" + item.MaxPrecision;
                        _cells.GetCell(8, i).Value = "" + item.MaxScale;
                        _cells.GetCell(9, i).Value = "" + item.AllowNull;
                        _cells.GetCell(10, i).Value = "" + item.DefaultValue;
                        _cells.GetCell(11, i).Value = "" + item.ActiveStatus;
                        _cells.GetCell(12, i).Value = "" + item.MeasureId;
                        _cells.GetCell(13, i).Value = "" + item.ValidationItemId;


                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Searched;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:AttributeSearchMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:AttributeSearchMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void AttributeSearchEachMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameAttribute;
                const int titleRow = TitleRowAttribute;
                const int resultColumn = ResultColumnAttribute;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(1, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                       
                        var modelId = "" + _cells.GetCell(1, i).Value;
                        var attributeId = "" + _cells.GetCell(2, i).Value;

                        var item = ModelAttribute.Read(modelId, attributeId);

                        if (item == null)
                            throw new ArgumentException(Resources.Error_ItemNotFound, nameof(item));
                        _cells.GetCell(1, i).Value = "" + item.ModelId();
                        _cells.GetCell(2, i).Value = "" + item.Id;
                        _cells.GetCell(3, i).Value = "" + item.Description;
                        _cells.GetCell(4, i).Value = "" + item.DataType;
                        _cells.GetCell(5, i).Value = "" + item.SheetIndex;
                        _cells.GetCell(6, i).Value = "" + item.MaxLength;
                        _cells.GetCell(7, i).Value = "" + item.MaxPrecision;
                        _cells.GetCell(8, i).Value = "" + item.MaxScale;
                        _cells.GetCell(9, i).Value = "" + item.AllowNull;
                        _cells.GetCell(10, i).Value = "" + item.DefaultValue;
                        _cells.GetCell(11, i).Value = "" + item.ActiveStatus;
                        _cells.GetCell(12, i).Value = "" + item.MeasureId();
                        _cells.GetCell(13, i).Value = "" + item.ValidationItemId();

                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Searched;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:AttributeSearchEachMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:AttributeSearchEachMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void AttributeUpdateMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameAttribute;
                const int titleRow = TitleRowAttribute;
                const int resultColumn = ResultColumnAttribute;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(1, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                        var modelId = "" + _cells.GetCell(1, i).Value;
                        var id = "" + _cells.GetCell(2, i).Value;
                        var description = "" + _cells.GetCell(3, i).Value;
                        var dataType = "" + _cells.GetCell(4, i).Value;
                        var sheetIndex = MyUtilities.ToIntegerNull(_cells.GetCell(5, i).Value, MyUtilities.ConversionConstants.DefaultNullAndEmpty);
                        var maxLength = MyUtilities.ToIntegerNull(_cells.GetCell(6, i).Value, MyUtilities.ConversionConstants.DefaultNullAndEmpty);
                        var maxPrecision = MyUtilities.ToIntegerNull(_cells.GetCell(7, i).Value, MyUtilities.ConversionConstants.DefaultNullAndEmpty);
                        var maxScale = MyUtilities.ToIntegerNull(_cells.GetCell(8, i).Value, MyUtilities.ConversionConstants.DefaultNullAndEmpty);
                        var allowNull = MyUtilities.IsTrue(_cells.GetCell(9, i).Value);
                        var defaultValue = "" + _cells.GetCell(10, i).Value;
                        var activeStatus = MyUtilities.IsTrue(_cells.GetCell(11, i).Value);
                        int? measureId = MyUtilities.ToIntegerNull(MyUtilities.GetCodeKey("" + _cells.GetCell(12, i).Value));
                        int? validItemId = MyUtilities.ToIntegerNull(MyUtilities.GetCodeKey("" + _cells.GetCell(13, i).Value));

                        var item = new ModelAttribute(modelId, id, description, dataType, sheetIndex, maxLength, maxPrecision, maxScale, allowNull, defaultValue, activeStatus, measureId, validItemId);

                        var reply = ModelAttribute.Create(item);
                        if (reply.Message.Equals(Resources.Results_Failed))
                            throw new Exception(reply.GetStringErrors());

                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Success;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:AttributeUpdateMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:AttributeUpdateMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void AttributeDeleteMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameAttribute;
                const int titleRow = TitleRowAttribute;
                const int resultColumn = ResultColumnAttribute;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(2, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                        var modelId = "" + _cells.GetCell(1, i).Value;
                        var id = "" + _cells.GetCell(2, i).Value;

                        if (string.IsNullOrWhiteSpace(modelId))
                            throw new ArgumentNullException(nameof(modelId), Resources.Error_NullValue);
                        if (string.IsNullOrWhiteSpace(id))
                            throw new ArgumentNullException(nameof(id), Resources.Error_NullValue);

                        ModelAttribute.Delete(modelId, id);

                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Deleted;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:AttributeDeleteMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:AttributeDeleteMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
    }
}
