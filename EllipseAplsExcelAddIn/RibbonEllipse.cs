using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Threading;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseAplsExcelAddIn
{
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        EllipseFunctions _eFunctions = new EllipseFunctions();
        private Excel.Application _excelApp;

        private const string SheetName01 = "AplItems";
        private const string SheetName02 = "AplHeader";

        private const int TitleRow01 = 9;
        private const int TitleRow02 = 9;

        private const int ResultColumn01 = 28;
        private const int ResultColumn02 = 24;

        private const string TableName01 = "AplItemsTable";
        private const string TableName02 = "TaskTable";

        private const string ValidationSheetName = "ValidationSheet";
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

        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        private void CreateAplItems()
        {
            var service = new APLItemService.APLItemService();
            var opContext = new APLItemService.OperationContext();
            var request = new APLItemService.APLItemServiceRetrieveRequestDTO();
            var requiredAttributes = new APLItemService.APLItemServiceRetrieveRequiredAttributesDTO();
            
            var reply = service.retrieve(opContext, request, requiredAttributes, "");
            var reply2 = service.retrieveAPLItems(opContext, request);
            
        }
        private void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                #region CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();

                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "APLs - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                var districtList = Districts.GetDistrictList();

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = Districts.DefaultDistrict;
                _cells.SetValidationList(_cells.GetCell("B3"), districtList, ValidationSheetName, 1);

                _cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01).Style = StyleConstants.TitleOptional;

                //GENERAL
                _cells.GetCell(1, TitleRow01).Value = "WORK_GROUP";
                _cells.GetCell(1, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, TitleRow01).Value = "WORK_ORDER";
                _cells.GetCell(2, TitleRow01).AddComment("Ingrese solo el prefijo si quiere crear una orden con prefijo");
                _cells.GetCell(3, TitleRow01).Value = "WO_STATUS";
                _cells.GetCell(3, TitleRow01).Style = StyleConstants.TitleInformation;

                _cells.GetCell(4, TitleRow01 - 2).Value = "GENERAL";
                _cells.GetCell(4, TitleRow01).Value = "DESCRIPTION";
                _cells.GetCell(4, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(5, TitleRow01).Value = "EQUIPMENT";
                _cells.GetCell(5, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(6, TitleRow01).Value = "COMP_CODE";
                _cells.GetCell(7, TitleRow01).Value = "MOD_CODE";

                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
    }
}
