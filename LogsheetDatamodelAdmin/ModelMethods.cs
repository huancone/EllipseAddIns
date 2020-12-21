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
    

        private void ModelSearchMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var tableName = TableNameModel;
                var titleRow = TitleRowModel;
                var resultColumn = ResultColumnModel;

                var id = "" + _cells.GetCell("B4").Value;
                var keyword = "" + _cells.GetCell("B5").Value;

                _cells.ClearTableRange(tableName);

                List<Datamodel> itemList;
                if (!string.IsNullOrWhiteSpace(id))
                    itemList = Datamodel.Read(id);
                else if (!string.IsNullOrWhiteSpace(keyword))
                    itemList = Datamodel.Read(id, keyword);
                else
                    itemList = Datamodel.Read();

                var i = titleRow + 1;
                foreach (var item in itemList)
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                        _cells.GetCell(1, i).Value = "" + item.Id;
                        _cells.GetCell(2, i).Value = "" + item.Description;
                        _cells.GetCell(3, i).Value = "" + item.ActiveStatus;
                        _cells.GetCell(4, i).Value = "" + item.CreationDate;
                        _cells.GetCell(5, i).Value = "" + item.CreationUser;
                        _cells.GetCell(6, i).Value = "" + item.LastModDate;
                        _cells.GetCell(7, i).Value = "" + item.LastModUser;

                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Searched;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ModelSearchMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:ModelSearchMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void ModelSearchEachMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var tableName = TableNameModel;
                var titleRow = TitleRowModel;
                var resultColumn = ResultColumnModel;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(1, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                       
                        var id = "" + _cells.GetCell(1, i).Value;

                        var item = Datamodel.ReadFirst(id);

                        if (item == null)
                            throw new ArgumentException(Resources.Error_ItemNotFound, nameof(item));
                        _cells.GetCell(1, i).Value = "" + item.Id;
                        _cells.GetCell(2, i).Value = "" + item.Description;
                        _cells.GetCell(3, i).Value = "" + item.ActiveStatus;
                        _cells.GetCell(4, i).Value = "" + item.CreationDate;
                        _cells.GetCell(5, i).Value = "" + item.CreationUser;
                        _cells.GetCell(6, i).Value = "" + item.LastModDate;
                        _cells.GetCell(7, i).Value = "" + item.LastModUser;

                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Searched;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ModelSearchEachMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:ModelSearchEachMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void ModelUpdateMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var tableName = TableNameModel;
                var titleRow = TitleRowModel;
                var resultColumn = ResultColumnModel;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(1, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                        var id = "" + _cells.GetCell(1, i).Value;
                        var description = "" + _cells.GetCell(2, i).Value;
                        var status = MyUtilities.IsTrue("" + _cells.GetCell(3, i).Value);

                        var item = new Datamodel();
                        item.Id = id;
                        item.Description = description;
                        item.ActiveStatus = status;
                        item.CreationUser = _frmAuth.User;
                        item.LastModUser = _frmAuth.User;
                        var result = Datamodel.Create(item);
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
                        Debugger.LogError("RibbonEllipse.cs:ModelUpdateMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:ModelUpdateMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void ModelDeleteMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var tableName = TableNameModel;
                var titleRow = TitleRowModel;
                var resultColumn = ResultColumnModel;

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

                        Datamodel.Delete(id);

                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Deleted;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ModelDeleteMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:ModelDeleteMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
    }
}
