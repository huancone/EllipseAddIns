using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Connections;
using Excel = Microsoft.Office.Interop.Excel;
using ProgramacionExcelAddins.Classes;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Vsto.Excel;
using Debugger = SharedClassLibrary.Utilities.Debugger;

namespace ProgramacionExcelAddins
{
    public partial class RibbonEllipse
    {
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private ExcelStyleCells _cells;
        private Excel.Application _excelApp;
        private HistorialProgramacion _historia;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }
        public void LoadSettings()
        {
            var settings = new Settings();
            _eFunctions = new EllipseFunctions();
            _frmAuth = new FormAuthenticate();
            _excelApp = Globals.ThisAddIn.Application;

            var environments = Environments.GetEnvironmentList();
            foreach (var env in environments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnvironment.Items.Add(item);
            }

            //settings.SetDefaultCustomSettingValue("AutoSort", "Y");
            //settings.SetDefaultCustomSettingValue("OverrideAccountCode", "Maintenance");
            //settings.SetDefaultCustomSettingValue("IgnoreItemError", "N");
            //settings.SetDefaultCustomSettingValue("AllowBackgroundWork", "N");

            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //var overrideAccountCode = settings.GetCustomSettingValue("OverrideAccountCode");
            //if (overrideAccountCode.Equals("Maintenance"))
            //    cbAccountElementOverrideMntto.Checked = true;
            //else if (overrideAccountCode.Equals("Disable"))
            //    cbAccountElementOverrideDisable.Checked = true;
            //else if (overrideAccountCode.Equals("Always"))
            //    cbAccountElementOverrideAlways.Checked = true;
            //else if (overrideAccountCode.Equals("Default"))
            //    cbAccountElementOverrideDefault.Checked = true;
            //else
            //    cbAccountElementOverrideDefault.Checked = true;
            //cbAutoSortItems.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("AutoSort"));
            //cbIgnoreItemError.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("IgnoreItemError"));
            //cbAllowBackgroundWork.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("AllowBackgroundWork"));

            //
            settings.SaveCustomSettings();

            _historia = new HistorialProgramacion();
        }
        private void btnConsultar_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteQuery();
        }
        private void ExecuteQuery()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var titleRow = 1;
                var sqlQuery = @"SELECT * FROM SIGMDC.HISTORIAL_PROGRAMACION WHERE FECHA='20190103' AND GRUPO='CTC'";
                var tableName = "table";
                conectarSigcor();
                var dataReader = _eFunctions.GetQueryResult(sqlQuery);

                if (dataReader == null || dataReader.IsClosed)
                    return;

                var data = dataReader;
                //Cargo el encabezado de la tabla y doy formato
                for (var i = 0; i < dataReader.FieldCount; i++)
                    _cells.GetCell(i + 1, titleRow).Value2 = "'" + dataReader.GetName(i);

                _cells.FormatAsTable(_cells.GetRange(1, titleRow, dataReader.FieldCount, titleRow + 1), tableName);
                

                var currentRow = titleRow + 1;
                while (dataReader.Read())
                {
                    for (var i = 0; i < dataReader.FieldCount; i++)
                        _cells.GetCell(i + 1, currentRow).Value2 = "'" + dataReader[i].ToString().Trim();
                    currentRow++;
                }

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:GetQueryResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private void btnCargar_Click(object sender, RibbonControlEventArgs e)
        {
            // CargarQuery();
            Recorrer();
        }
        private void Recorrer()
        {
            try
            {
                if (_cells == null) _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                var i = 2;
                //Recorre la tabla
                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value) != null)
                {
                    
                    if (!TieneHistoria(_cells.GetEmptyIfNull(_cells.GetCell(1, i).Value.ToString()), _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value.ToString()), _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value.ToString())))
                    {
                        _cells.GetCell(8, i).Value = InsertarHistoria(
                            _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value.ToString()),
                            _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value.ToString()), 
                            _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value.ToString()), 
                            _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value.ToString()),
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value), 
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value)
                            );
                        if (_cells.GetCell(8, i).Value== "Registro realizado")
                            _cells.GetCell(8, i).Style = StyleConstants.Success;
                        else
                            _cells.GetCell(8, i).Style = StyleConstants.Error;
                    }
                    else
                    {
                        _cells.GetCell(8, i).Value = ModificarHistoria(
                            _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value.ToString()),
                            _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value.ToString()),
                            _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value.ToString()),
                            _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value.ToString()),
                            _cells.GetNullOrTrimmedValue(_cells.GetCell(5, i).Value),
                            _cells.GetNullOrTrimmedValue(_cells.GetCell(6, i).Value)
                            );

                        if (_cells.GetCell(8, i).Value == "Registro modificado")
                            _cells.GetCell(8, i).Style = StyleConstants.Warning;
                        else
                            _cells.GetCell(8, i).Style = StyleConstants.Error;
                    }
                    i += 1;
                  
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:GetQueryResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
                
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }
        public String InsertarHistoria (String fecha, String grupo, String idConcepto, String concepto, String valor1, String valor2)
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                if (valor2 == "" || valor2 == null ) valor2 = "NULL";
                if (valor1 == "" || valor1 == null) valor1 = "NULL";
                var sqlQuery = @"INSERT INTO SIGMDC.HISTORIAL_PROGRAMACION  VALUES ('"+ fecha + "','" + grupo + "','" + idConcepto + "','" + concepto + "'," + valor1 + "," + valor2 + ")";
                //return sqlQuery;
                conectarSigcor();
                var dataReader = _eFunctions.GetQueryResult(sqlQuery);
                return "Registro realizado";
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:GetQueryResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                //MessageBox.Show(@"Error de Insersion. " + ex.Message);
                return @"Error de Insersion. " + ex.Message;
               
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
            }

        }
        public String ModificarHistoria(String fecha, String grupo, String idConcepto, String concepto, String valor1, String valor2)
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                if (valor2 == "" || valor2 == null) valor2 = "NULL";
                if (valor1 == "" || valor1 == null) valor1 = "NULL";
                var sqlQuery = @"UPDATE SIGMDC.HISTORIAL_PROGRAMACION SET CONCEPTO='" + concepto + "', VALOR1=" + valor1 + ", VALOR2=" + valor2 + " WHERE FECHA='" + fecha + "' AND GRUPO = '" + grupo + "' AND ID_CONCEPTO = '" + idConcepto + "'";

                conectarSigcor();
                var dataReader = _eFunctions.GetQueryResult(sqlQuery);
                return "Registro modificado";
                
                

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:GetQueryResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                //MessageBox.Show(@"Error de Insersion. " + ex.Message);
                return @"Error de Insersion. " + ex.Message;

            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
            }

        }
        public bool TieneHistoria(String fecha, String grupo, String idConcepto)
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                var resultado = 0;
                var sqlQuery = @"SELECT COUNT(*) AS UNO FROM SIGMDC.HISTORIAL_PROGRAMACION WHERE FECHA='"+ fecha + "' AND GRUPO='"+grupo+ "'AND ID_CONCEPTO='" + idConcepto + "'";
                conectarSigcor();
                var dataReader = _eFunctions.GetQueryResult(sqlQuery);
                
                //Sino hay nada en el datarader, entonces es que no tiene historia
                if (dataReader == null) return false ;
                
                while (dataReader.Read())
                {
                   resultado = Int32.Parse(dataReader[0].ToString());
                }
                
                //Si tiene un registro es decir que existe, sino, entonces retorna falso
                if (resultado != 0) return true; else return false;
                
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:GetQueryResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
                return true;
            }
            finally
            {
                _eFunctions.CloseConnection();
            }
        }

        private void conectarSigcor()
        {
            _eFunctions.SetDBSettings("SIGCOPRD", "consulbo", "consulbo");
        }

        private void btnEliminar_Click(object sender, RibbonControlEventArgs e)
        {
            EliminarHistoria();
        }

        private void EliminarHistoria()
        {
            try
            {
                if (_cells == null) _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                var i = 2;
                var fecha="";
                var grupo="";
                var idConcepto = "";
                //Recorre la tabla
                while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value) != null)
                {
                    fecha = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value.ToString());
                    grupo = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value.ToString());
                    idConcepto = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value.ToString());
                    if (TieneHistoria(fecha, grupo, idConcepto))
                    {
                        var sqlQuery = @"DELETE FROM SIGMDC.HISTORIAL_PROGRAMACION  WHERE FECHA='" + fecha + "' AND GRUPO = '" + grupo + "' AND ID_CONCEPTO = '" + idConcepto + "'";
                        conectarSigcor();
                        var dataReader = _eFunctions.GetQueryResult(sqlQuery);
                        _cells.GetCell(8, i).Value = "Registro eliminado!";
                    }
                    else
                    {
                        _cells.GetCell(8, i).Value = "No existe informacion para este registro";
                    }
                    i += 1;

                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:GetQueryResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);

            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }
    }
}
