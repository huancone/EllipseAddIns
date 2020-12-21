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

        private void ValidationSourcesSearchMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameValidSources;
                const int titleRow = TitleRowValidSources;
                const int resultColumn = ResultColumnValidSources;

                var keyword = "" + _cells.GetCell("B4").Value;

                _cells.ClearTableRange(tableName);

                var i = titleRow + 1;

                List<ValidationSource> SourceList = ValidationSource.Read(keyword, keyword);

                foreach (var Source in SourceList)
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                        _cells.GetCell(1, i).Value = "" + Source.Name;
                        _cells.GetCell(2, i).Value = "" + Source.DbName;
                        _cells.GetCell(3, i).Value = "" + Source.DbUser;
                        _cells.GetCell(4, i).Value = "" + Source.DbPassword;
                        _cells.GetCell(5, i).Value = "" + Source.DbReference;
                        _cells.GetCell(6, i).Value = "" + Source.DbLink;
                        _cells.GetCell(7, i).Value = "" + Source.PasswordEncodedType;

                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Searched;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ValidationSourcesSearchMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:ValidationSourcesSearchMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
        private void ValidationSourcesSearchEachMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameValidSources;
                const int titleRow = TitleRowValidSources;
                const int resultColumn = ResultColumnValidSources;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(1, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;
                        var id = MyUtilities.ToInteger("" + _cells.GetCell(2, i).Value);

                        var Source = ValidationSource.Read(id);

                        if (Source == null)
                            throw new ArgumentException(Resources.Error_ItemNotFound, nameof(Source));
                        _cells.GetCell(1, i).Value = "" + Source.Name;
                        _cells.GetCell(2, i).Value = "" + Source.DbName;
                        _cells.GetCell(3, i).Value = "" + Source.DbUser;
                        _cells.GetCell(4, i).Value = "" + Source.DbPassword;
                        _cells.GetCell(5, i).Value = "" + Source.DbReference;
                        _cells.GetCell(6, i).Value = "" + Source.DbLink;
                        _cells.GetCell(7, i).Value = "" + Source.PasswordEncodedType;

                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Searched;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ValidationSourcesSearchEachMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:ValidationSourcesSearchEachMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
        private void ValidationSourcesUpdateMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameValidSources;
                const int titleRow = TitleRowValidSources;
                const int resultColumn = ResultColumnValidSources;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(1, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                        var sourceName = "" + _cells.GetCell(1, i).Value;
                        var dbName = "" + _cells.GetCell(2, i).Value;
                        var dbUser = "" + _cells.GetCell(3, i).Value;
                        var dbPassword = "" + _cells.GetCell(4, i).Value;
                        var dbReference = "" + _cells.GetCell(5, i).Value;
                        var dbLink = "" + _cells.GetCell(6, i).Value;
                        var passwordEncodedType = "" + _cells.GetCell(7, i).Value;

                        var Source = new ValidationSource(sourceName, dbName, dbUser, dbPassword, dbReference, dbLink, passwordEncodedType);


                        var result = ValidationSource.Create(Source);
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
                        Debugger.LogError("RibbonEllipse.cs:ValidationSourcesUpdateMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:ValidationSourcesUpdateMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void ValidationSourcesDeleteMethod()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                const string tableName = TableNameValidSources;
                const int titleRow = TitleRowValidSources;
                const int resultColumn = ResultColumnValidSources;

                _cells.ClearTableRangeColumn(tableName, resultColumn);

                var i = titleRow + 1;

                while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(2, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                        var sourceName = "" + _cells.GetCell(1, i).Value;
                        if (string.IsNullOrWhiteSpace(sourceName))
                            throw new ArgumentNullException(nameof(sourceName), Resources.Error_NullValue);

                        ValidationSource.Delete(sourceName);

                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Value = Resources.Results_Deleted;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Value = Resources.Error_ErrorUppercase + ": " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ValidationSourcesDeleteMethod()", ex.Message);
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
                Debugger.LogError("RibbonEllipse.cs:ValidationSourcesDeleteMethod()", ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
    }
}
