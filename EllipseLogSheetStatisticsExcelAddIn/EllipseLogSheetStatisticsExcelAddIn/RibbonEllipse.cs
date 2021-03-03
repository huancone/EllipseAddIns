using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading;
using Microsoft.Office.Tools.Ribbon;
using Screen = SharedClassLibrary.Ellipse.ScreenService;
using SharedClassLibrary;
using SharedClassLibrary.Utilities;
using System.Web.Services.Ellipse;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Vsto.Excel;


namespace EllipseLogSheetStatisticsExcelAddIn
{
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Thread _thread;

        Excel.Application _excelApp;

        private const string SheetName01 = "LogSheetStatistics";

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

            //settings.SetDefaultCustomSettingValue("OptionName1", "false");
            //settings.SetDefaultCustomSettingValue("OptionName2", "OptionValue2");
            //settings.SetDefaultCustomSettingValue("OptionName3", "OptionValue3");



            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, SharedResources.Settings_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            //var optionItem1Value = MyUtilities.IsTrue(settings.GetCustomSettingValue("OptionName1"));
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName2");
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName3");

            //cbCustomSettingOption.Checked = optionItem1Value;
            //optionItem2.Text = optionItem2Value;
            //optionItem3 = optionItem3Value;

            //
            settings.SaveCustomSettings();
        }
        private void btnFormatLogSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheetHeaderData();
        }
        private void btnLoadModel_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {

                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    
                    _thread = new Thread(FormatModelData);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:FormatModelData()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnCreateLogSheetStatistics_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {

                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(StartCreateLogSheet);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:FormatModelData()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }


        }

        public void FormatSheetHeaderData()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "MAINTAIN LOGSHEET STATISTICS - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "J2");

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

                //Cells.GetCell("A3").Value = "DISTRICT";
                //Cells.GetCell("A3").Style = Cells.GetStyle(StyleConstants.Option);
                //Cells.GetCell("B3").Style = Cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("A4").Value = "MODEL";
                _cells.GetCell("A4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B4").Style = _cells.GetStyle(StyleConstants.Select);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        public void StartCreateLogSheet()
        {
            try
            {

                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName01)
                    throw new Exception("\nLa hoja seleccionada no coincide con el modelo requerido");

                if (string.IsNullOrWhiteSpace(drpEnvironment.SelectedItem.Label))
                    throw new Exception("\nSeleccione un ambiente válido");

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var opContext = new Screen.OperationContext()
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = Debugger.DebugWarnings
                };

                var modelCode = "" + _cells.GetCell("B4").Value;
                var globalStartIndex = 7;
                var globalEndIndex = globalStartIndex;
                for (var i = globalStartIndex; _cells.GetCell(1, i).Value != null; i++)
                    globalEndIndex = i;

                //Preparación de los datos para envío

                //cargo los encabezados de fila y los equipos aceptados para el modelo
                var modelEquipments = GetModelEquipments(_eFunctions, modelCode);
                //vienen ordenados en el vector según el número de secuencia respectivo
                var modelHeaders = GetModelHeaders();

                if (modelEquipments == null || (modelHeaders == null || modelHeaders.Count == 0))
                    throw new Exception(@"No se pudo obtener información del modelo");

                    _cells.GetRange(1, globalStartIndex, 3, globalEndIndex).Style = StyleConstants.Normal;
                _cells.GetRange(modelHeaders.Count() + 1, globalStartIndex, modelHeaders.Count() + 1, globalEndIndex)
                    .Style = StyleConstants.Normal;
                //ordeno los valores de la hoja para turno
                var unorderRange = _cells.GetRange(1, globalStartIndex, modelHeaders.Count(), globalEndIndex);
                unorderRange.Sort(unorderRange.Columns[1, Type.Missing], Excel.XlSortOrder.xlAscending,
                    unorderRange[2, Type.Missing], Type.Missing);


                var logStartIndex = globalStartIndex;
                //para propósitos de resaltar en la hoja errores y warnings en bloques

                var previousDate = "" + _cells.GetCell(1, globalStartIndex).Value;
                var previousShift = "" + _cells.GetCell(2, globalStartIndex).Value;


                var matrixValues = new List<string[]>();
                //almacenará los valores de forma como estén en la hoja de excel, excluyendo los que no pertenezcan al modelo
                var finalValues = new List<string[]>();
                //almacenará los valores de forma ordenada que será la que se enviará finalmente

                for (var i = globalStartIndex; i <= globalEndIndex + 1; i++)
                {
                    //si es fecha-turno diferente envíe lo que tiene
                    if (!Convert.ToString("" + _cells.GetCell(1, i).Value).Equals(previousDate) ||
                        !Convert.ToString("" + _cells.GetCell(2, i).Value).Equals(previousShift))
                    {
                        //
                        foreach (var mEq in modelEquipments)
                        {
                            var inList = false; //garantizará que exista al menos uno en ceros
                            foreach (var mVal in matrixValues)
                            {
                                if (mVal[3] != mEq) continue;
                                inList = true;
                                finalValues.Add(mVal);
                            }

                            if (inList) continue;
                            var rowValues = new string[modelHeaders.Count()];
                            rowValues[0] = "S"; //ACTION
                            rowValues[1] = previousDate;
                            rowValues[2] = previousShift;
                            rowValues[3] = mEq;
                            for (var j = 4; j < rowValues.Length; j++)
                                rowValues[j] = "";
                            finalValues.Add(rowValues);
                        }

                        //marcar con I el registro anterior al duplicado para indicar al screen que debe hacer la acción
                        for (var k = 0; k < finalValues.Count() - 1; k++)
                            if (finalValues.ElementAt(k)[3].Equals(finalValues.ElementAt(k + 1)[3]))
                                finalValues.ElementAt(k)[0] = "I";




                        var logEndIndex = i - 1;
                        var logResult = CreateLogSheet(opContext, modelCode, previousDate, previousShift, finalValues);

                        if (logResult.StartsWith("SUCCESS"))
                        {
                            _cells.GetRange(1, logStartIndex, 1, logEndIndex).Style = StyleConstants.Success;
                            _cells.GetRange(1, logStartIndex, 1, logEndIndex).BorderAround2();
                        }
                        else if (logResult.StartsWith("ERROR"))
                        {
                            _cells.GetRange(1, logStartIndex, 1, logEndIndex).Style = StyleConstants.Error;
                            _cells.GetRange(1, logStartIndex, 1, logEndIndex).BorderAround2();
                        }
                        else
                        {
                            _cells.GetRange(1, logStartIndex, 1, logEndIndex).Style = StyleConstants.Warning;
                            _cells.GetRange(1, logStartIndex, 1, logEndIndex).BorderAround2();
                        }

                        _cells.GetCell(modelHeaders.Count() + 1, i - 1).Value = logResult;
                        //reinicie los objetos del screen
                        matrixValues = new List<string[]>();
                        //almacenará los valores de forma como estén en la hoja de excel, excluyendo los que no pertenezcan al modelo
                        finalValues = new List<string[]>();
                        //almacenará los valores de forma ordenada que será la que se enviará finalmente
                        logStartIndex = i;
                    }

                    previousDate = "" + _cells.GetCell(1, i).Value;
                    previousShift = "" + _cells.GetCell(2, i).Value;
                    //si sigue en la misma fecha-turno siga
                    //valido que el registro de la fila exista en el modelo
                    if (modelEquipments.Contains(("" + _cells.GetCell(3, i).Value).Trim()))
                        //si existe se añade a la lista a ser agregado
                    {
                        var rowValues = new string[modelHeaders.Count() + 1];
                        rowValues[0] = ""; //para ACTION
                        rowValues[1] = "" + _cells.GetCell(1, i).Value; //DATE
                        rowValues[2] = "" + _cells.GetCell(2, i).Value; //SHIFT

                        for (var j = 3; j < rowValues.Length; j++)
                            rowValues[j] = "" + _cells.GetCell(j, i).Value;
                        //lo adiciono a la lista de registros aceptados (no ordenada)
                        matrixValues.Add(rowValues);
                    }
                    else
                        //si no existe se resalta el error y se continúa el proceso ignorando el registro (no será cargado)
                        _cells.GetCell(3, i).Style = StyleConstants.Error;

                }

                //
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Rows.AutoFit();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:createLogSheet()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                _cells?.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }
        public string CreateLogSheet(Screen.OperationContext opSheet, string modelCode, string modelDate, string modelShift, List<string[]> sheetData)
        {
            //Proceso del screen
            var proxySheet = new Screen.ScreenService();
            var requestSheet = new Screen.ScreenSubmitRequestDTO();
            var arrayFields = new ArrayScreenNameValue();

            //Selección de ambiente
            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            //Aseguro que no esté en alguna pantalla antigua
            _eFunctions.RevertOperation(opSheet, proxySheet);
            //ejecutamos el programa
            var replySheet = proxySheet.executeScreen(opSheet, "MSO435");

            //validamos el ingreso al programa
            if (replySheet.mapName != "MSM435A" || _excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName01)
                return "ERROR:" + "No se pudo establecer comunicación con el servicio";

            arrayFields.Add("OPTION1I", "1");
            arrayFields.Add("MODEL_CODE1I", modelCode);
            arrayFields.Add("STAT_DATE1I", modelDate);
            arrayFields.Add("SHIFT1I", modelShift);
            //arrayFields.Add("MODEL_MODE1I",""); //no usado
            //arrayFields.Add("RUN_ID1I", ""); //no usado

            requestSheet.screenFields = arrayFields.ToArray();
            requestSheet.screenKey = "1";

            replySheet = proxySheet.submit(opSheet, requestSheet);

            _eFunctions.CheckReplyWarning(replySheet);//si hay debug activo muestra el warning de lo contrario depende del proceso del OP



            if (replySheet != null && !_eFunctions.CheckReplyError(replySheet) && replySheet.mapName == "MSM435B")
            {

                //Creamos la nueva pantalla de envío reutilizando los elementos anteriores
                requestSheet = new Screen.ScreenSubmitRequestDTO();
                arrayFields = new ArrayScreenNameValue();

                //ingresamos los elementos (name, value) para los campos a enviar   
                arrayFields.Add("STAT_DATE2I", modelDate);
                arrayFields.Add("SHIFT2I", modelShift);

                var screenIndex = 1;
                foreach (var row in sheetData)
                {
                    if (screenIndex > 7)
                    {
                        //enviar Screen
                        requestSheet.screenFields = arrayFields.ToArray();
                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);
                        arrayFields = new ArrayScreenNameValue();
                        //
                        if (replySheet != null && replySheet.mapName != "MSM435B")
                            break;
                        screenIndex = 1;
                    }

                    //eS(screenIndex) = fv
                    arrayFields.Add("ACTION2I" + screenIndex, row[0]);
                    arrayFields.Add("PLANT_NO2I" + screenIndex, row[3]);
                    arrayFields.Add("OPERATOR2I" + screenIndex, row[4]);
                    arrayFields.Add("ACCOUNT_CODE2I" + screenIndex, row[5]);
                    arrayFields.Add("WORK_ORDER2I" + screenIndex, row[6]);
                    for (var i = 7; i < row.Length; i++)
                        arrayFields.Add("INPUT_" + (i - 6) + "2I" + screenIndex, row[i]);
                    //
                    if (row[0] == "I")
                    {
                        //enviar Screen
                        requestSheet.screenFields = arrayFields.ToArray();
                        requestSheet.screenKey = "1";
                        replySheet = proxySheet.submit(opSheet, requestSheet);
                        var field = arrayFields.GetField("ACTION2I" + screenIndex);//Se reinicia el valor para que al enviar no vuelva a hacer insert, sino simplemente continúe con el screen
                        field.value = "";
                        //
                        if (replySheet != null && replySheet.mapName != "MSM435B")
                            break;

                        if (screenIndex >= 7) //es una condición especial cuando se añade estando en el último registro, porque el sistema envía y cambia el screen de una vez
                        {
                            screenIndex = 0; //se iguala a cero porque al terminar el bucle exterior sube el index a 1, que es lo que se necesitaría para la siguiente iteración
                            arrayFields = new ArrayScreenNameValue();
                        }
                    }
                    screenIndex++;
                }
                requestSheet = new Screen.ScreenSubmitRequestDTO
                {
                    screenFields = arrayFields.ToArray(),
                    screenKey = "1"
                };

                replySheet = proxySheet.submit(opSheet, requestSheet);
                _eFunctions.CheckReplyWarning(replySheet);//si hay debug activo muestra el warning de lo contrario depende del proceso del OP

                if (replySheet != null && !_eFunctions.CheckReplyError(replySheet) && replySheet.mapName == "MSM435A")
                    return "SUCCESS:" + "Se han cargado exitosamente los datos";
                if (replySheet != null && _eFunctions.CheckReplyError(replySheet))
                    return "ERROR:" + replySheet.message;
                return "ERROR:" + "Se produjo un error al intentar cargar los datos";
            }

            if (replySheet == null)
                return "ERROR:" + "No se puede establecer conexión con el programa MSM435B";
            if (replySheet.mapName != "MSM435B" || replySheet.message.Substring(0, 2) == "X2")
                return "ERROR:" + replySheet.message;
            return "ERROR:" + replySheet.message;


            //---fin proceso del screen
        }
        public void FormatModelData()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                _cells.ClearRange("A6", "AZ65536");
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                var modelCode = "" + _cells.GetCell("B4").Value;
                //encabezados
                var sqlQuery1 = Queries.GetDefaultHeaderData(modelCode, _eFunctions.DbReference, _eFunctions.DbLink);
                //equipos
                var sqlQuery2 = Queries.GetQueryDefaultModelData(modelCode, _eFunctions.DbReference, _eFunctions.DbLink);
                //Igual que el query de getModelEquipment

                if (Debugger.DebugQueries)
                {
                    _cells.GetCell("L1").Value = sqlQuery1;
                    _cells.GetCell("M1").Value = sqlQuery2;
                }

                var drHeaders = _eFunctions.GetQueryResult(sqlQuery1);
                var headerValues = new List<ModelHeaderNameValue>();


                if (drHeaders == null || drHeaders.IsClosed)
                    throw new Exception("No se han encontrado datos para el modelo especificado");

                _cells.GetCell(1, 6).Value = "DATE";
                _cells.GetCell(1, 6).AddComment("YYYYMMDD");
                _cells.GetCell(1, 6).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, 6).Value = "SHIFT";
                _cells.GetCell(2, 6).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, 6).Value = "EQUIP_REF";
                _cells.GetCell(3, 6).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, 6).Value = "OPERATOR";
                _cells.GetCell(4, 6).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, 6).Value = "ACCOUNT_CODE"; //siempre bloqueado en los existentes
                _cells.GetCell(5, 6).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(6, 6).Value = "W/O"; //siempre bloqueado en los existentes
                _cells.GetCell(6, 6).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                var i = 7;
                while (drHeaders.Read())
                {
                    var hv = new ModelHeaderNameValue
                    {
                        Name = ("" + drHeaders["HEADER_NAME"]).Trim(),
                        Type = ("" + drHeaders["VALUE_TYPE"]).Trim(),
                        Index = int.Parse("" + drHeaders["INDICE"])
                    };

                    if (hv.Name.Equals("")) continue;

                    _cells.GetCell(i, 6).Value = hv.Name;
                    _cells.GetCell(i, 6).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                    i++;
                    headerValues.Add(hv);
                }

                var drEquipments = _eFunctions.GetQueryResult(sqlQuery2);

                if (drEquipments == null || drEquipments.IsClosed)
                    return;

                var j = 7;
                var arrayHeader = headerValues.ToArray();

                while (drEquipments.Read())
                {

                    var rv = new ModelRowValue
                    {
                        Code = ("" + drEquipments["ENTRY_GRP"]),
                        EquipReference = ("" + drEquipments["EQ_REFERENCE"]),
                        Operator =
                        {
                            Flag = ("" + drEquipments["OPERATOR_FLG"]).Equals("O") ||
                                   ("" + drEquipments["OPERATOR_FLG"]).Equals("M"),
                            Value = ("" + drEquipments["OPERATOR_ID"])
                        },
                        Account =
                        {
                            Flag = ("" + drEquipments["ACCOUNT_FLG"]).Equals("O") ||
                                   ("" + drEquipments["ACCOUNT_FLG"]).Equals("M"),
                            Value = ("" + drEquipments["ACCOUNT_CODE"])
                        },
                        WorkOrder =
                        {
                            Flag = ("" + drEquipments["WORK_ORDER_FLG"]).Equals("O") ||
                                   ("" + drEquipments["WORK_ORDER_FLG"]).Equals("M"),
                            Value = ("" + drEquipments["WORK_ORDER"])
                        },
                        Source =
                        {
                            Flag = ("" + drEquipments["SOURCE_LOC_FLG"]).Equals("O") ||
                                   ("" + drEquipments["SOURCE_LOC_FLG"]).Equals("M"),
                            Value = ("" + drEquipments["SOURCE_LOC"])
                        },
                        Destination =
                        {
                            Flag = ("" + drEquipments["DEST_LOC_FLG"]).Equals("O") ||
                                   ("" + drEquipments["DEST_LOC_FLG"]).Equals("M"),
                            Value = ("" + drEquipments["DEST_LOC"])
                        },
                        Material =
                        {
                            Flag = ("" + drEquipments["MATERIAL_FLG"]).Equals("O") ||
                                   ("" + drEquipments["MATERIAL_FLG"]).Equals("M"),
                            Value = ("" + drEquipments["MATERIAL_CODE"])
                        }
                    };
                    //Flags de esta sección O: Optional, M: Mandatory, N: Not Required
                    //Flags de esta sección I: Input, O: Output, B: Both
                    for (var k = 0; k < 10; k++)
                    {
                        rv.Inputs[k].Flag = ("" + drEquipments["STAT_IO_FLG_" + (k + 1)]).Equals("I") ||
                                            ("" + drEquipments["STAT_IO_FLG_" + (k + 1)]).Equals("B");
                        rv.Inputs[k].Value = ("" + drEquipments["STAT_VALUE_" + (k + 1)]);
                    }

                    //Escribo los valores obtenidos según el valor predeterminado
                    _cells.GetCell(3, j).Value = "'" + rv.Code.Trim();
                    if (rv.EquipReference != "")
                        _cells.GetCell(3, j).AddComment(rv.EquipReference);
                    //operator
                    _cells.GetCell(4, j).Value = (rv.Operator.Value.Equals("") ? "" : "'" + rv.Operator.Value);
                    _cells.GetCell(4, j).Style = (rv.Operator.Flag ? "Normal" : StyleConstants.Disabled);
                    //account
                    _cells.GetCell(5, j).Value = (rv.Account.Value.Equals("") ? "" : "'" + rv.Account.Value);
                    _cells.GetCell(5, j).Style = (rv.Account.Flag ? "Normal" : StyleConstants.Disabled);
                    //wo
                    _cells.GetCell(6, j).Value = (rv.WorkOrder.Value.Equals("") ? "" : "'" + rv.WorkOrder.Value);
                    _cells.GetCell(6, j).Style = (rv.WorkOrder.Flag ? "Normal" : StyleConstants.Disabled);

                    //asigna el valor por defecto según corresponda
                    for (var k = 0; k < arrayHeader.Length; k++)
                    {
                        if (arrayHeader[k].Type.Equals("SS")) //source
                            _cells.GetCell(7 + k, j).Value = rv.Source.Value;
                        else if (arrayHeader[k].Type.Equals("SD")) //destination
                            _cells.GetCell(7 + k, j).Value = rv.Destination.Value;
                        else if (arrayHeader[k].Type.Equals("ML")) //material
                            _cells.GetCell(7 + k, j).Value = rv.Material.Value;
                        else
                            _cells.GetCell(7 + k, j).Value = 0;

                        _cells.GetCell(7 + k, j).Style = (rv.Inputs[k].Flag ? "Normal" : StyleConstants.Disabled);
                    }
                    j++;
                }

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetModelData()", ex.Message);
                MessageBox.Show(ex.Message);
            }
            finally
            {
				_eFunctions.CloseConnection();
                _cells?.SetCursorDefault();
            }
        }
        public List<string> GetModelHeaders()
        {
            try
            {
                var headerList = new List<string>();
                const int headerRow = 6;
                var i = 1;
                while ("" + _cells.GetCell(i, headerRow).Value != "")
                {
                    headerList.Add("" + _cells.GetCell(i, headerRow));
                    i++;
                }
                return headerList;
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:getModelHeader()", ex.Message);
                return null;
            }
        }
        public List<string> GetModelEquipments(EllipseFunctions ef, string modelCode)
        {
			try
            {
                List<string> equipmentList = null;
                //Igual que el query sqlQuery2 de setSheetModelData
                var sqlQuery = Queries.GetQueryDefaultModelData(modelCode, _eFunctions.DbReference, _eFunctions.DbLink);

                if (Debugger.DebugQueries)
                    _cells.GetCell("M1").Value = sqlQuery;

                var drEquipments = ef.GetQueryResult(sqlQuery);

                if (drEquipments != null && !drEquipments.IsClosed)
                {
                    equipmentList = new List<string>();
                    while (drEquipments.Read())
                        equipmentList.Add("" + drEquipments["ENTRY_GRP"].ToString().Trim());
                }

                return equipmentList;
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetModelData()", ex.Message);
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }


    }
}
