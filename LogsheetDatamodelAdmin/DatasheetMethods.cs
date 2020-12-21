using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using LogsheetDatamodelLibrary;
using LogsheetDatamodelLibrary.Configuration;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelAdmin
{
    public partial class RibbonLsdm
    {
        private void DatasheetSearchMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameDatasheet;
                const int titleRow = TitleRowDatasheet;
                const int resultColumn = ResultColumnDatasheet;

                var modelId = "" + _cells.GetCell("B4").Value;
                var activeOnly = "" + _cells.GetCell("B5").Value;
                var strStartDate = "" + _cells.GetCell("B6").Value;
                var strFinishDate = "" + _cells.GetCell("B7").Value;
                var startDate = MyUtilities.ToDate(strStartDate);
                var finishDate = MyUtilities.ToDate(strFinishDate);



                _cells.ClearTableRange(tableName);
                _cells.DeleteTableRange(tableName);
                _cells.ClearRange((titleRow - 1) + ":" + (titleRow - 1));
                var i = titleRow + 1;

                if(string.IsNullOrWhiteSpace(modelId))
                    throw new ArgumentNullException(nameof(modelId), Resources.Error_InvalidId);
                Datamodel model = Datamodel.ReadFirst(modelId);
                var attributesHeader = model.PullModelAttributes();
                List<Datasheet> itemList = Datasheet.ReadHeader(modelId, startDate, finishDate);


                //Table Structure Creation
                _cells.GetCell(1, titleRow).Value = LsdmResource.Datasheet_Date;
                _cells.GetCell(1, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, titleRow).Value = LsdmResource.Datasheet_Shift;
                _cells.GetCell(2, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, titleRow).Value = LsdmResource.Datasheet_SequenceId;
                _cells.GetCell(3, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.SetValidationList(_cells.GetCell(2, titleRow + 1), ValidationSheetName, 1, false);
                var curColumn = resultColumn;
                foreach (var att in attributesHeader)
                {
                    if(_cells.GetCell(curColumn, titleRow).Comment != null)
                        _cells.GetCell(curColumn, titleRow).Comment.Delete();
                    _cells.GetCell(curColumn, titleRow).AddComment(att.Description);
                    _cells.GetCell(curColumn, titleRow).Value = att.Id;
                    _cells.GetCell(curColumn, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                    _cells.GetCell(curColumn, titleRow - 1).Value = att.DataType;
                    _cells.GetCell(curColumn, titleRow - 1).Style = _cells.GetStyle(StyleConstants.Select);
                    curColumn++;
                }


                _cells.GetCell(resultColumn + attributesHeader.Count, titleRow).Value = Resources.Title_ResultCamelcase;
                _cells.GetCell(resultColumn + attributesHeader.Count, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn + attributesHeader.Count, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn + attributesHeader.Count, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                //
                foreach (var item in itemList)
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                        _cells.GetCell(1, i).Value = MyUtilities.ToString(item.Date);
                        _cells.GetCell(2, i).Value = "" + item.Shift;
                        _cells.GetCell(3, i).Value = "" + item.SequenceId;
                        curColumn = resultColumn;
                        foreach (var vObj in item.ValueObjects)
                        {
                            if (!vObj.AttributeId.Equals("" + _cells.GetCell(curColumn, titleRow).Value, StringComparison.OrdinalIgnoreCase))
                                _cells.GetCell(curColumn, i).Style = StyleConstants.Error;

                            if (vObj.DataType().Equals(DataTypes.Date, StringComparison.OrdinalIgnoreCase))
                                _cells.GetCell(curColumn, i).Value = MyUtilities.ToString(vObj.Value, MyUtilities.DateTime.DateDefaultFormat);
                            else if (vObj.DataType().Equals(DataTypes.DateTime, StringComparison.OrdinalIgnoreCase))
                                _cells.GetCell(curColumn, i).Value = MyUtilities.ToString(vObj.Value, MyUtilities.DateTime.DateTimeDefaultFormat);
                            else
                                _cells.GetCell(curColumn, i).Value = vObj.Value;

                            curColumn++;
                        }
                        _cells.GetCell(resultColumn + attributesHeader.Count, i).Value = Resources.Results_Searched;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn + attributesHeader.Count, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn + attributesHeader.Count, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:DatasheetSearchMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:DatasheetSearchMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void DatasheetUpdateMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameDatasheet;
                const int titleRow = TitleRowDatasheet;
                const int resultColumn = ResultColumnDatasheet;

                var modelId = "" + _cells.GetCell("B4").Value;

                if (string.IsNullOrWhiteSpace(modelId))
                    throw new ArgumentNullException(nameof(modelId), Resources.Error_InvalidId);

                List<ModelAttribute> attributesHeader = ModelAttribute.Read(modelId);

                _cells.ClearTableRangeColumn(tableName, resultColumn + attributesHeader.Count);
                var curRow = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(1, curRow).Value))
                {
                    var curColumn = resultColumn;
                    try
                    {
                        var date = MyUtilities.ToDate("" + _cells.GetCell(1, curRow).Value);
                        var shift = "" + _cells.GetCell(2, curRow).Value;
                        var sequenceId = "" + _cells.GetCell(3, curRow).Value;

                        Datasheet datasheet = Datasheet.ReadFirstHeader(modelId, date, shift, sequenceId);
                        if (datasheet == null)
                        {
                            var replyHeader = Datasheet.CreateHeader(modelId, null, date, shift, sequenceId, LsdmConfig.Login.User);
                            if (replyHeader.Message.Equals(Resources.Results_Failed))
                                throw new Exception(replyHeader.GetStringErrors());
                            datasheet = Datasheet.ReadFirstHeader(modelId, date, shift, sequenceId);
                        }
                        
                        string curHeader = "" + _cells.GetCell(curColumn, titleRow).Value;
                        while (!string.IsNullOrWhiteSpace(curHeader) && !curHeader.Equals(Resources.Title_ResultCamelcase))
                        {
                            string dataType = "" + _cells.GetCell(curColumn, titleRow - 1).Value;

                            var vObj = new ValueObject();
                            vObj.ModelId = modelId;
                            vObj.SheetId = datasheet.Id;
                            vObj.AttributeId = curHeader;
                            if (dataType.Equals(DataTypes.Date, StringComparison.OrdinalIgnoreCase))
                            {
                                if(_cells.IsNumberFormatDate(curColumn, curRow))
                                    vObj.SetValue((double)(MyUtilities.ToDate(_cells.GetCell(curColumn, curRow).Value)));
                                else
                                    vObj.SetValue((MyUtilities.ToDate("" + _cells.GetCell(curColumn, curRow).Value)));
                            }
                            else if (dataType.Equals(DataTypes.DateTime, StringComparison.OrdinalIgnoreCase))
                            {
                                if (_cells.IsNumberFormatDate(curColumn, curRow))
                                    vObj.SetValue((double)(MyUtilities.ToDateTime(_cells.GetCell(curColumn, curRow).Value)));
                                else
                                    vObj.SetValue((MyUtilities.ToDateTime("" + _cells.GetCell(curColumn, curRow).Value)));
                            }
                            else if (dataType.Equals(DataTypes.Numeric, StringComparison.OrdinalIgnoreCase))
                                vObj.SetValue(MyUtilities.ToDecimal(_cells.GetCell(curColumn, curRow).Value));
                            else
                                vObj.SetValue("" + _cells.GetCell(curColumn, curRow).Value);
                            datasheet.ValueObjects.Add(vObj);

                            curColumn++;
                            curHeader = "" + _cells.GetCell(curColumn, titleRow).Value;
                        }

                        var reply = datasheet.PushValueObjects();

                        if (reply.Message.StartsWith(LsdmResource.Results_Success))
                        {
                            _cells.GetCell(resultColumn + attributesHeader.Count, curRow).Style = StyleConstants.Success;
                            _cells.GetCell(resultColumn + attributesHeader.Count, curRow).Value = reply.Message;
                        }
                        else
                        {
                            _cells.GetCell(resultColumn + attributesHeader.Count, curRow).Style = StyleConstants.Success;
                            _cells.GetCell(resultColumn + attributesHeader.Count, curRow).Value = reply.Message + " " + reply.GetStringWarnings() + " " + reply.GetStringErrors();
                        }
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn + attributesHeader.Count, curRow).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn + attributesHeader.Count, curRow).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:DatasheetUpdateMethod()", ex.Message);
                    }
                    finally
                    {
                        _cells.GetCell(2, curRow).Select();
                        curRow++;
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Resources.Error_ErrorFound, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse.cs:DatasheetUpdateMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void DatasheetDeleteMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameDatasheet;
                const int titleRow = TitleRowDatasheet;
                const int resultColumn = ResultColumnDatasheet;

                var modelId = "" + _cells.GetCell("B4").Value;

                if (string.IsNullOrWhiteSpace(modelId))
                    throw new ArgumentNullException(nameof(modelId), Resources.Error_InvalidId);

                List<ModelAttribute> attributesHeader = ModelAttribute.Read(modelId);

                _cells.ClearTableRangeColumn(tableName, resultColumn + attributesHeader.Count);
                var curRow = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(1, curRow).Value))
                {
                    try
                    {
                        var date = MyUtilities.ToDate("" + _cells.GetCell(1, curRow).Value);
                        var shift = "" + _cells.GetCell(2, curRow).Value;
                        var sequenceId = "" + _cells.GetCell(3, curRow).Value;

                        Datasheet datasheet = Datasheet.ReadFirstHeader(modelId, date, shift, sequenceId);
                        if (datasheet == null)
                        {
                            var replyHeader = Datasheet.CreateHeader(modelId, null, date, shift, sequenceId, LsdmConfig.Login.User);
                            if (replyHeader.Message.Equals(Resources.Results_Failed))
                                throw new Exception(replyHeader.GetStringErrors());
                            datasheet = Datasheet.ReadFirstHeader(modelId, date, shift, sequenceId);
                        }

                        var reply = Datasheet.Delete(datasheet.Id);

                        if (reply.Message.StartsWith(LsdmResource.Results_Success))
                        {
                            _cells.GetCell(resultColumn + attributesHeader.Count, curRow).Style = StyleConstants.Success;
                            _cells.GetCell(resultColumn + attributesHeader.Count, curRow).Value = reply.Message;
                        }
                        else
                        {
                            _cells.GetCell(resultColumn + attributesHeader.Count, curRow).Style = StyleConstants.Success;
                            _cells.GetCell(resultColumn + attributesHeader.Count, curRow).Value = reply.Message + " " + reply.GetStringWarnings() + " " + reply.GetStringErrors();
                        }
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn + attributesHeader.Count, curRow).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn + attributesHeader.Count, curRow).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:DatasheetUpdateMethod()", ex.Message);
                    }
                    finally
                    {
                        _cells.GetCell(2, curRow).Select();
                        curRow++;
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Resources.Error_ErrorFound, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse.cs:DatasheetUpdateMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

    }
}
