using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using System.Web.Services.Ellipse;
using System.Threading;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel; 
using EllipseStdTextClassLibrary;
using EllipseReferenceCodesClassLibrary;

namespace EllipseStdTextExcelAddIn
{
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
    public partial class RibbonEllipse
    {

        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions = new EllipseFunctions();
        private FormAuthenticate _frmAuth = new FormAuthenticate();
        private Excel.Application _excelApp;
        private const string SheetName01 = "StdText";
        private const string SheetName02 = "ReferenceCodes";
        private const int TitleRow01 = 5;
        private const int TitleRow02 = 5;
        private const int ResultColumn01 = 6;
        private const int ResultColumn02 = 10;
        private const string TableName01 = "StdTextTable";
        private const string TableName02 = "RefCodeTable";
        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
           
            var environments = Environments.GetEnvironmentList();
            foreach (var env in environments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnvironment.Items.Add(item);
            }
        }

        private void btnFormatSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        private void btnGetHeaderAndText_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                {
                    GetStdText(true, true);
                }
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnUpdateHeaderAndText_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                {
                    SetStdText(true, true);
                }
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnGetHeaderOnly_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                {
                    GetStdText(true, false);
                }
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnSetHeaderOnly_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment= drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                {
                    SetStdText(true, false);
                }
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnGetTextOnly_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                {
                    GetStdText(false, true);
                }
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnSetTextOnly_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() == DialogResult.OK)
                {
                    SetStdText(false, true);
                }
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        public void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
               
                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "STD TEXT - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);


                _cells.GetCell(1, TitleRow01).Value = "TYPE[2]";
                _cells.GetCell(1, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(2, TitleRow01).Value = "DISTRICT[4]";
                _cells.GetCell(2, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(3, TitleRow01).Value = "ID[8]";
                _cells.GetCell(3, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(4, TitleRow01).Value = "HEADER";
                _cells.GetCell(4, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(5, TitleRow01).Value = "TEXT";
                _cells.GetCell(5, TitleRow01).Style = StyleConstants.TitleRequired;

                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow01+1 , ResultColumn01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO LA HOJA 2 - REFERENCE CODES
                _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "STD TEXT - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);


                _cells.GetCell(1, TitleRow02).Value = "ENT. TYPE [3]";
                _cells.GetCell(1, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, TitleRow02).Value = "ENT. VALUE 1";
                _cells.GetCell(2, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(3, TitleRow02).Value = "ENT. VALUE 2";
                _cells.GetCell(3, TitleRow02).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, TitleRow02).Value = "ENT. NO";
                _cells.GetCell(4, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(5, TitleRow02).Value = "ENT. SEQ";
                _cells.GetCell(5, TitleRow02).Style = StyleConstants.TitleOptional;
                _cells.GetCell(6, TitleRow02).Value = "REF CODE";
                _cells.GetCell(6, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(7, TitleRow02).Value = "ST FLAG";
                _cells.GetCell(7, TitleRow02).Style = StyleConstants.TitleInformation;
                _cells.GetCell(8, TitleRow02).Value = "STDTEXT ID";
                _cells.GetCell(8, TitleRow02).Style = StyleConstants.TitleOptional;
                _cells.GetCell(9, TitleRow02).Value = "STDTEXT TEXT";
                _cells.GetCell(9, TitleRow02).Style = StyleConstants.TitleOptional;

                _cells.GetCell(ResultColumn02, TitleRow02).Value = "RESULTADO";
                _cells.GetCell(ResultColumn02, TitleRow02).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow02 + 1, ResultColumn02, TitleRow02 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow02, ResultColumn02, TitleRow02 + 1), TableName02);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                
                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatSheet()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        public void GetStdText(bool getHeader, bool getText)
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;
            var districtCode = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
            var position = _frmAuth.EllipsePost;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(3, i).Value))
            {
                try
                {
                    var stdTextId = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value) +
                                    _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value) +
                                    _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);
                    var stdTextOpc = StdText.GetStdTextOpContext(districtCode, position, 100, Debugger.DebugWarnings);
                    if (getHeader)
                    {
                        
                        string headerText = StdText.GetHeader(urlService, stdTextOpc, stdTextId);
                        _cells.GetCell(4, i).Value = "" + headerText;
                    }

                    if (getText)
                    {
                        string bodyText = StdText.GetText(urlService, stdTextOpc, stdTextId);
                        _cells.GetCell(5, i).Value = "" + bodyText;
                    }

                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:GetStdText(bool, bool)", ex.Message);
                }
                finally
                {
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        public void SetStdText(bool setHeader, bool setText)
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;
            var districtCode = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
            var position = _frmAuth.EllipsePost;
            var stdTextOpc = StdText.GetStdTextOpContext(districtCode, position, 100, Debugger.DebugWarnings);
            var stdTextCustomOpc = StdText.GetCustomOpContext(districtCode, position, 100, Debugger.DebugWarnings);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(3, i).Value))
            {
                try
                {
                    var stdTextId = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value) +
                                    _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value) +
                                    _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);

                    var resultHeader = true;
                    var resultBody = true;

                    if (setHeader)
                    {
                        var headerText = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value);
                        resultHeader = StdText.SetHeader(urlService, stdTextOpc, stdTextId, headerText);
                    }

                    if (setText)
                    {
                        var bodyText = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value);
                        resultBody = StdText.SetText(urlService, stdTextCustomOpc, stdTextId, bodyText);
                    }

                    if (resultHeader && resultBody)
                    {
                        _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                        _cells.GetCell(ResultColumn01, i).Value = "ACTUALIZADO";
                    }
                    else
                    {
                        _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn01, i).Value = "ERROR";
                    }

                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:GetStdText(bool, bool)", ex.Message);
                }
                finally
                {
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        public void ReviewRefCodesList()
        {

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            district = string.IsNullOrWhiteSpace(district) ? "ICOR" : district;

            var rcOpContext = ReferenceCodeActions.GetRefCodesOpContext(district, _frmAuth.EllipsePost, 100, Debugger.DebugWarnings);
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow02 + 1;

            //Se encuentran problemas de implementación, debido a un comportamiento irregular del ODP en Windows. 
            //Las conexiones cerradas (EllipseFunctions.Close()) vuelven a la piscina (pool) de conexiones por un tiempo antes 
            //de ser completamente Cerradas (Close) y Dispuestas (Dispose), lo que ocasiona un desbordamiento del
            //número máximo de conexiones en el pool (100) y la nueva conexión alcanza el tiempo de espera (timeout) antes de
            //entrar en la cola del pool de conexiones arrojando un error 'Pooled Connection Request Timed Out'.
            //Para solucionarlo se fuerza el string de conexiones para que no genere una conexión que entre al pool.
            //Esto implica mayor tiempo de ejecución pero evita la excepción por el desbordamiento y tiempo de espera
            _eFunctions.SetConnectionPoolingType(false);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {

                    var entityType = "" + _cells.GetCell(1, i).Value;
                    var entityValue = "" + _cells.GetCell(2, i).Value + _cells.GetCell(3, i).Value;
                    var refNum = "" + _cells.GetCell(4, i).Value;
                    var seqNum = "" + _cells.GetCell(5, i).Value;

                    refNum = string.IsNullOrWhiteSpace(refNum) ? "001" : refNum.PadLeft(3, '0');
                    seqNum = string.IsNullOrWhiteSpace(seqNum) ? "001" : seqNum.PadLeft(3, '0'); 

                    if (entityType.Equals("WRQ"))
                        entityValue = entityValue.PadLeft(12, '0');

                    var refItem = ReferenceCodeActions.FetchReferenceCodeItem(_eFunctions, urlService, rcOpContext, entityType, entityValue, refNum, seqNum);
                    //GENERAL
                    _cells.GetCell(6, i).Value = "'" + refItem.RefCode;
                    _cells.GetCell(7, i).Value = "'" + refItem.StdTextFlag;
                    _cells.GetCell(8, i).Value = "'" + refItem.StdtxtId;
                    _cells.GetCell(9, i).Value = "'" + refItem.StdText;

                    _cells.GetCell(ResultColumn02, i).Value = "CONSULTADO";
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewRefCodesList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _eFunctions.SetConnectionPoolingType(true);
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        public void UpdateRefCodesList()
        {

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            district = string.IsNullOrWhiteSpace(district) ? "ICOR" : district;

            var rcOpContext = ReferenceCodeActions.GetRefCodesOpContext(district, _frmAuth.EllipsePost, 100, Debugger.DebugWarnings);
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow02 + 1;

            //Se encuentran problemas de implementación, debido a un comportamiento irregular del ODP en Windows. 
            //Las conexiones cerradas (EllipseFunctions.Close()) vuelven a la piscina (pool) de conexiones por un tiempo antes 
            //de ser completamente Cerradas (Close) y Dispuestas (Dispose), lo que ocasiona un desbordamiento del
            //número máximo de conexiones en el pool (100) y la nueva conexión alcanza el tiempo de espera (timeout) antes de
            //entrar en la cola del pool de conexiones arrojando un error 'Pooled Connection Request Timed Out'.
            //Para solucionarlo se fuerza el string de conexiones para que no genere una conexión que entre al pool.
            //Esto implica mayor tiempo de ejecución pero evita la excepción por el desbordamiento y tiempo de espera
            _eFunctions.SetConnectionPoolingType(false);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {

                    var entityType = "" + _cells.GetCell(1, i).Value;
                    var entityValue = "" + _cells.GetCell(2, i).Value + _cells.GetCell(3, i).Value;
                    var refNum = "" + _cells.GetCell(4, i).Value;
                    var seqNum = "" + _cells.GetCell(5, i).Value;
                    var refCode = "" + _cells.GetCell(6, i).Value;
                    var stdTextId = "" + _cells.GetCell(8, i).Value;
                    var stdText = "" + _cells.GetCell(9, i).Value;
                    refNum = string.IsNullOrWhiteSpace(refNum) ? "001" : refNum.PadLeft(3, '0');
                    seqNum = string.IsNullOrWhiteSpace(seqNum) ? "001" : seqNum.PadLeft(3, '0'); 

                    if (entityType.Equals("WRQ"))
                        entityValue = entityValue.PadLeft(12, '0');

                    var refItem = new ReferenceCodeItem(entityType, entityValue, refNum, seqNum, refCode, stdTextId, stdText);
                    ReferenceCodeActions.ModifyRefCode(_eFunctions, urlService, rcOpContext, refItem);

                    _cells.GetCell(ResultColumn02, i).Value = "ACTUALIZADO";
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateRefCodesList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _eFunctions.SetConnectionPoolingType(true);
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void btnCleanTable_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                _cells.ClearTableRange(TableName01);
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
                _cells.ClearTableRange(TableName02);
        }
        private void btnReviewRefCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName02))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewRefCodesList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewRefCodesList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnUpdateRefCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName02))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(UpdateRefCodesList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:UpdateRefCodesList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                if (_cells != null) _cells.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }
        }
        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }
}
