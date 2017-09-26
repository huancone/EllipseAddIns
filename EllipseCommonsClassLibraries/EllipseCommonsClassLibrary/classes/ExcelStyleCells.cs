using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

// ReSharper disable LoopCanBeConvertedToQuery

namespace EllipseCommonsClassLibrary
{
    public class ExcelStyleCells
    {
        private readonly Application _excelApp;
        private Worksheet _excelSheet;

        private bool _alwaysActiveSheet;
        private CultureInfo _oldCultureInfo;
        /// <summary>
        /// Constructor de la clase. Si alwaysActiveSheet es true La clase estará sujeta a la hoja activa con la que se esté trabajando, si es false, estará sujeta exclusivamente a la hoja activa desde la que se invoca este constructor
        /// </summary>
        /// <param name="excelApp">Microsoft.Office.Interop.Excel.Application Aplicación Excel en ejecución</param>
        /// <param name="alwaysActiveSheet">bool: Determina si se ejecutará según la hoja activa de excel</param>
        public ExcelStyleCells(Application excelApp, bool alwaysActiveSheet = true)
        {
            _excelApp = excelApp;
            _alwaysActiveSheet = alwaysActiveSheet;
            try
            {
                //Si hay un libro activo (Ej. Office 2013+ inicia sin libro activo)
                if (_excelApp.ActiveWorkbook == null) return;
                if (Debugger.ForceRegionalization)
                    SetEllipseDefaultCulture();//Se adiciona instrucción para evitar conflictos de símbolos por diferencias de lenguaje
                _excelSheet = (Worksheet) _excelApp.ActiveWorkbook.ActiveSheet;
                CreateStyles();
            }
            catch(Exception ex)
            {
                Debugger.LogError("Se ha producido un error al intentar inicializar la clase Commons>ExcelStyleCells. ", ex.Message);
            }
        }

        /// <summary>
        /// Constructor de la clase. La clase estará sujeta solamente a la hoja de trabajo ingresada en SheetName
        /// </summary>
        /// <param name="excelApp">Microsoft.Office.Interop.Excel.Application Aplicación Excel en ejecución</param>
        /// <param name="sheetName">string: Especifica el nombre de la hoja para la que se le realizarán las acciones con esta clase</param>
        public ExcelStyleCells(Application excelApp, string sheetName)
        {
            _excelApp = excelApp;
            _alwaysActiveSheet = false;
            try
            {
                //Si hay un libro activo (Ej. Office 2013+ inicia sin libro activo)
                if (_excelApp.ActiveWorkbook == null) return;
                if (Debugger.ForceRegionalization)
                    SetEllipseDefaultCulture();
                _excelSheet = (Worksheet)_excelApp.ActiveWorkbook.ActiveSheet;
                foreach (Worksheet sheet in _excelApp.ActiveWorkbook.Sheets)
                {
                    if (sheet.Name != sheetName) continue;
                    _excelSheet = sheet;
                    break;
                }
                CreateStyles();
            }
            catch(Exception ex)
            {
                Debugger.LogError("Se ha producido un error al intentar inicializar la clase Commons>ExcelStyleCells. ", ex.Message);
            }
        }

        /// <summary>
        /// Establece como hoja de trabajo a la hoja activa
        /// </summary>
        /// <returns>bool: true si hay una hoja activa disponible</returns>
        public bool SetActiveSheet()
        {
            try
            {
                _excelSheet = (Worksheet)_excelApp.ActiveWorkbook.ActiveSheet;
                return true;
            }
            catch(Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:SetActiveSheet", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                return false;
            }
        }

        /// <summary>
        /// Establece como hoja de trabajo a la hoja ingresada en sheetName. Si no existe no hace cambios
        /// </summary>
        /// <returns>bool: true si hay una hoja que coincida con sheetName</returns>
        public bool SetActiveSheet(string sheetName)
        {
            try
            {
                foreach (Worksheet ws in _excelApp.ActiveWorkbook.Sheets)
                {
                    if (ws.Name != sheetName) continue;
                    _excelSheet = ws;
                    return true;
                }
                return false;

            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:SetActiveSheet", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                return false;
            }
        }
        /// <summary>
        /// Establece el valor de alwaysActiveSheet (siempre la hoja activa) como True o False. Si es true, ejecutará siempre las acciones en la hoja activa. Si es false, ejecutará las acciones en la hoja que se haya especificado en el momento de su creación o la hoja establecida en la clase
        /// </summary>
        /// <param name="value">bool: valor booleano a asignar a alwaysActiveSheet</param>
        public void SetAlwaysActiveSheet(bool value)
        {
            _alwaysActiveSheet = value;
        }

        /// <summary>
        /// Cambia el valor de alwaysActiveSheet. Si está en true, lo cambia a false y viceversa
        /// </summary>
        /// <returns>bool: Estado final del valor de alwaysActiveSheet</returns>
        public bool ToggleAlwaysActiveSheet()
        {
            _alwaysActiveSheet = !_alwaysActiveSheet;
            return _alwaysActiveSheet;
        }

        //CELLS

        /// <summary>
        /// Obtiene la celda de una hoja a partir de la columna y fila de la misma Ej: (4, 3) Columna 4, Fila 3
        /// </summary>
        /// <param name="column">long: columna de la celda </param>
        /// <param name="row">long: fila de la celda</param>
        /// <returns>Excel.Range Celda solicitada</returns>
        public Range GetCell(long column, long row)
        {
            var excelSheet = (_alwaysActiveSheet && _excelSheet != null) ? (Worksheet)_excelApp.ActiveWorkbook.ActiveSheet : _excelSheet;
            return excelSheet != null ? (Range)excelSheet.Cells[row, column] : null;
        }

        /// <summary>
        /// Obtiene la celda de una hoja a partir del nombre de la celda Ej: (A2) Columna A, Fila 2
        /// </summary>
        /// <param name="cell">string: nombre de la celda Ej. (A2) </param>
        /// <returns>Microsoft.Office.Interop.Excel.Range Celda solicitada</returns>
        public Range GetCell(string cell)
        {
            var excelSheet = (_alwaysActiveSheet && _excelSheet != null) ? (Worksheet)_excelApp.ActiveWorkbook.ActiveSheet : _excelSheet;
            return excelSheet != null ? excelSheet.Range[cell] : null;
        }

        /// <summary>
        /// Borra los datos de una celda y devuelve si la acción se realizó o no
        /// </summary>
        /// <param name="cell">string: nombre de la celda Ej. (A2) </param>
        /// <returns>bool: La acción se realizó sin problemas</returns>
        public bool ClearCell(string cell)
        {
            try
            {
                var cellsRange = GetCell(cell);
                if(cellsRange != null)
                    cellsRange.Clear();
                return true;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:clearCell(string)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                return false;
            }
        }
        /// <summary>
        /// Borra los datos de una celda y devuelve si la acción se realizó o no
        /// </summary>
        /// <param name="column">long: columna de la celda </param>
        /// <param name="row">long: fila de la celda</param>
        /// <returns>bool: La acción se realizó sin problemas</returns>
        public bool ClearCell(long column, long row)
        {
            try
            {
                var cellsRange = GetCell(column, row);
                if(cellsRange != null)
                    cellsRange.Clear();
                return true;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:clearCell(long, long)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                return false;
            }
        }


        
        //RANGES

        /// <summary>
        /// Obtiene el rango de una hoja entre una celda inicial y una celda final Ej: ("A2", "F8")
        /// </summary>
        /// <param name="startRange">string: celda inicial Ej. ("A2") </param>
        /// <param name="endRange">string: celda final Ej. ("F8") </param>
        /// <returns>Microsoft.Office.Interop.Excel.Range Rango solicitado</returns>
        public Range GetRange(string startRange, string endRange)
        {
            Worksheet excelSheet = (_alwaysActiveSheet && _excelSheet != null) ? (Worksheet)_excelApp.ActiveWorkbook.ActiveSheet : _excelSheet;
            return excelSheet != null ? excelSheet.Range[startRange + ":" + endRange] : null;
        }

        /// <summary>
        /// Obtiene el rango de una hoja entre una celda inicial y una celda final Ej: (1, 2, 6, 8) para ("A2", "F8")
        /// </summary>
        /// <param name="startColumn">long: columna de la celda inicial Ej. (1) </param>
        /// <param name="startRow">long: fila de la celda inicial Ej. (2) </param>
        /// <param name="endColumn">long: columna de la celda final Ej. (6) </param>
        /// <param name="endRow">long: fila de la celda final Ej. (8) </param>
        /// <returns>Microsoft.Office.Interop.Excel.Range Rango solicitado</returns>
        public Range GetRange(long startColumn, long startRow, long endColumn, long endRow)
        {
            Worksheet excelSheet = (_alwaysActiveSheet && _excelSheet != null) ? (Worksheet)_excelApp.ActiveWorkbook.ActiveSheet : _excelSheet;
            return excelSheet != null ? excelSheet.Range[GetCell(startColumn, startRow), GetCell(endColumn, endRow)] : null;
        }

        /// <summary>
        /// Obtiene el rango de una hoja según el nombre dado (Ej: rangos como tablas)
        /// </summary>
        /// <param name="rangeName">string: nombre dado al rango (Ej: ValuesTable</param>
        /// <returns></returns>
        public Range GetRange(string rangeName)
        {
            Worksheet excelSheet = (_alwaysActiveSheet && _excelSheet != null) ? (Worksheet)_excelApp.ActiveWorkbook.ActiveSheet : _excelSheet;
            return excelSheet != null ? excelSheet.Range[rangeName] : null;
        }

        /// <summary>
        /// Borra los datos de un rango y devuelve si la acción se realizó o no
        /// </summary>
        /// <param name="startRange">string: nombre de la celda inicial Ej. ("A2") </param>
        /// <param name="endRange">string: nombre de la celda final Ej. ("F8") </param>
        /// <returns>bool: La acción se realizó sin problemas</returns>
        public bool ClearRange(string startRange, string endRange)
        {
            try
            {
                Range cellsRange = GetRange(startRange, endRange);
                cellsRange.Clear();
                return true;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:clearRange(string, string)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                return false;
            }
        }
        /// <summary>
        /// Borra los datos de un rango y devuelve si la acción se realizó o no
        /// </summary>
        /// <param name="startColumn">long: columna inicial del rango </param>
        /// <param name="startRow">long: fila inicial del rango</param>
        /// <param name="endColumn">long: columna inicial del rango </param>
        /// <param name="endRow">long: fila inicial del rango</param>
        /// <returns>bool: La acción se realizó sin problemas</returns>
        public bool ClearRange(long startColumn, long startRow, long endColumn, long endRow)
        {
            try
            {
                var cellsRange = GetRange(startColumn, startRow, endColumn, endRow);
                cellsRange.Clear();
                return true;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:clearRange(long, long, long, long)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                return false;
            }
        }
        /// <summary>
        /// Combina un rango desde la celda inciial dada hasta la celda final dada
        /// </summary>
        /// <param name="startRange">Celda inicial del rango a combinar Ej. ("A2")</param>
        /// <param name="endRange">Celda final del rango a combinar Ej. ("F8")</param>
        public void MergeCells(string startRange, string endRange)
        {
            var cellsRange = GetRange(startRange, endRange);
            cellsRange.Merge();
            cellsRange.WrapText = true;
        }

        /// <summary>
        /// Combina un rango desde la celda inciial dada hasta la celda final dada
        /// </summary>
        /// <param name="startColumn">long: columna inicial del rango </param>
        /// <param name="startRow">long: fila inicial del rango</param>
        /// <param name="endColumn">long: columna inicial del rango </param>
        /// <param name="endRow">long: fila inicial del rango</param>
        public void MergeCells(long startColumn, long startRow, long endColumn, long endRow)
        {
            var cellsRange = GetRange(startColumn, startRow, endColumn, endRow);
            cellsRange.Merge();
            cellsRange.WrapText = true;
        }

        //SHEETS
        /// <summary>
        /// Establece un nuevo nombre a una Hoja dada según el índice Ej. Hoja(index) = newSheetName
        /// </summary>
        /// <param name="index">int: índice de hoja a cambiar nombre</param>
        /// <param name="newSheetName">string: nuevo nombre para la hoja</param>
        /// <returns>true: si se realiza la acción. false: si no se realiza la acción por algún error o porque no existe la hoja ingresada</returns>
        public bool SetSheetName(int index, string newSheetName)
        { //TO CHECK
            try
            {
                ((Worksheet)_excelApp.ActiveWorkbook.Sheets[index]).Name = newSheetName;
                
                return true;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells::setSheetName(int, string)", ex.Message);
                return false;
            }
        }
        /// <summary>
        /// Establece un nuevo nombre a una hoja según el nombre antiguo dado Ej. Hoja(oldSheetName) = newSheetName
        /// </summary>
        /// <param name="oldSheetName">string: Nombre de hoja a renombrar</param>
        /// <param name="newSheetName">string: Nuevo nombre dado</param>
        /// <returns></returns>
        public bool SetSheetName(string oldSheetName, string newSheetName)
        { //TO CHECK
            try
            {
                foreach (Worksheet ws in _excelApp.ActiveWorkbook.Sheets)
                {
                    if (ws.Name != oldSheetName) continue;
                    ws.Name = newSheetName;
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells::setSheetName(int, string)", ex.Message);
                return false;
            }
        }
        /// <summary>
        /// Obtiene la hoja del libro según el nombre indicado. Si la hoja no existe retorna null
        /// </summary>
        /// <param name="worksheetName">string: nombre de la hoja de trabajo</param>
        /// <returns></returns>
        public Worksheet GetWorksheet(string worksheetName)
        {
            foreach (Worksheet sheet in _excelApp.ActiveWorkbook.Sheets)
            {
                if (sheet.Name != worksheetName) continue;
                return sheet;
            }
            return null;
        }
        /// <summary>
        /// Obtiene la hoja del libro según el nombre indicado. Si la hoja no existe retorna null
        /// </summary>
        /// <param name="index">int: índice de la hoja de trabajo (índice inicial 1)</param>
        /// <returns></returns>
        public Worksheet GetWorksheet(int index)
        { 
            try
            {
                return (Worksheet)_excelApp.ActiveWorkbook.Sheets[index];
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells::setSheetName(int, string)", ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Establece la visibilidad de la hoja.
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <param name="visible">bool: true para Visible, false para oculto</param>
        public void SetWorksheetVisibility(string worksheetName, bool visible)
        {
            var sheet = GetWorksheet(worksheetName);
            if (sheet == null)
                return;
            if (visible)
                sheet.Visible = XlSheetVisibility.xlSheetVisible;
            if (!visible)
                sheet.Visible = XlSheetVisibility.xlSheetHidden;
        }

        /// <summary>
        /// Crea una nueva hoja de trabajo al final del libro activo
        /// </summary>
        /// <param name="worksheetName">string: Nombre de la nueva hoja de trabajo</param>
        /// <returns></returns>
        public Worksheet CreateNewWorksheet(string worksheetName)
        {
            var currentSheetIndex = ((Worksheet) _excelApp.ActiveWorkbook.ActiveSheet).Index;
            
            _excelApp.ActiveWorkbook.Worksheets.Add(After: _excelApp.ActiveWorkbook.Sheets[_excelApp.ActiveWorkbook.Sheets.Count]);
            ((Worksheet)_excelApp.ActiveWorkbook.Sheets[_excelApp.ActiveWorkbook.Worksheets.Count]).Select(Type.Missing);

            ((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name = worksheetName;
            var newSheet = (Worksheet)_excelApp.ActiveWorkbook.ActiveSheet;

            ((Worksheet)_excelApp.ActiveWorkbook.Sheets[currentSheetIndex]).Select(Type.Missing);
            return newSheet;

        }
    
        //STYLES

        /// <summary>
        /// Obtiene un estilo a partir de un nombre de estilo dado. El nombre del estilo coincide con los valores de estilos existentes en StyleConstants.StyleName. Si el estilo ingresado no existe, devuelve el estilo Normal
        /// </summary>
        /// <param name="styleName">string: nombre de estilo a obtener</param>
        /// <returns>Microsoft.Office.Interop.Excel.Style styleName El estilo solicitado</returns>
        public Style GetStyle(string styleName)
        {
            try
            {
                // ReSharper disable once LoopCanBePartlyConvertedToQuery
                foreach (Style style in _excelApp.ActiveWorkbook.Styles)
                    if (style.Name == styleName)
                        return style;
                //si no existe, retorna el normal
                if (!StyleConstants.GetStyleListName().Contains(styleName))
                    return GetStyleNormal();
                CreateStyles();
                return GetStyle(styleName);
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:getStyle(string)", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                return GetStyleNormal();
            }
        }
        /// <summary>
        /// Obtiene el estilo Normal predeterminado
        /// </summary>
        /// <returns>Microsoft.Office.Interop.Excel.Style styleName El estilo solicitado</returns>
        public Style GetStyleNormal()
        {
            try
            {
                foreach (Style style in _excelApp.ActiveWorkbook.Styles)
                    if (style.Name == StyleConstants.Normal)
                        return style;
                //si no existe, crea nuevamente los estilos e invoca recursividad
                CreateStyles();
                return GetStyleNormal();
            }
            catch(Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:getStyleNormal()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                CreateStyles();
                return GetStyleNormal();
            }
        }
        /// <summary>
        /// Obtiene el estilo Error predeterminado
        /// </summary>
        /// <returns>Microsoft.Office.Interop.Excel.Style styleName El estilo solicitado</returns>
        public Style GetStyleError()
        {
            try
            {
                foreach (Style style in _excelApp.ActiveWorkbook.Styles)
                    if (style.Name == StyleConstants.Error)
                        return style;
                //si no existe, crea nuevamente los estilos e invoca recursividad
                CreateStyles();
                return GetStyleError();
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:getStyleError()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                CreateStyles();
                return GetStyleError();
            }
        }
        /// <summary>
        /// Obtiene el estilo Warning predeterminado
        /// </summary>
        /// <returns>Microsoft.Office.Interop.Excel.Style styleName El estilo solicitado</returns>
        public Style GetStyleWarning()
        {
            try
            {
                foreach (Style style in _excelApp.ActiveWorkbook.Styles)
                    if (style.Name == StyleConstants.Warning)
                        return style;
                //si no existe, crea nuevamente los estilos e invoca recursividad
                CreateStyles();
                return GetStyleWarning();
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:getStyleWarning()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                CreateStyles();
                return GetStyleWarning();
            }
        }
        /// <summary>
        /// Obtiene el estilo Success predeterminado
        /// </summary>
        /// <returns>Microsoft.Office.Interop.Excel.Style styleName El estilo solicitado</returns>
        public Style GetStyleSuccess()
        {
            try
            {
                foreach (Style style in _excelApp.ActiveWorkbook.Styles)
                    if (style.Name == StyleConstants.Success)
                        return style;
                //si no existe, crea nuevamente los estilos e invoca recursividad
                CreateStyles();
                return GetStyleSuccess();
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:getStyleSuccess()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                CreateStyles();
                return GetStyleSuccess();
            }
        }
        /// <summary>
        /// Crea los estilos predeterminados que va a tener la clase
        /// </summary>
        private void CreateStyles()
        {
            try
            {
                //Normal
                if (!ExistStyle(StyleConstants.Normal))
                {
                    var styleNormal = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.Normal, Type.Missing);
                    styleNormal.NumberFormat = "General";
                }
                //Success
                if (!ExistStyle(StyleConstants.Success))
                {
                    var styleSuccess = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.Success, Type.Missing);
                    styleSuccess.Font.Name = "Calibri";
                    styleSuccess.Font.Size = 10;
                    styleSuccess.Font.Color = ColorTranslator.ToOle(Color.Black);
                    styleSuccess.NumberFormat = "General";
                    styleSuccess.Interior.Color = ColorTranslator.ToOle(Color.LightGreen);
                }
                //Warning
                if (!ExistStyle(StyleConstants.Warning))
                {
                    var styleWarning = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.Warning, Type.Missing);
                    styleWarning.Font.Name = "Calibri";
                    styleWarning.Font.Size = 10;
                    styleWarning.Font.Color = ColorTranslator.ToOle(Color.Black);
                    styleWarning.NumberFormat = "General";
                    styleWarning.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                }
                //Error
                if (!ExistStyle(StyleConstants.Error))
                {
                    var styleError = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.Error, Type.Missing);
                    styleError.Font.Name = "Calibri";
                    styleError.Font.Size = 10;
                    styleError.Font.Color = ColorTranslator.ToOle(Color.White);
                    styleError.NumberFormat = "General";
                    styleError.Interior.Color = ColorTranslator.ToOle(Color.Red);
                }
                //HeaderDefault
                if (!ExistStyle(StyleConstants.HeaderDefault))
                {
                    var styleHeaderDefault = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.HeaderDefault, Type.Missing);
                    styleHeaderDefault.Font.Name = "MS Sans Serif";
                    styleHeaderDefault.Font.Size = 13;
                    styleHeaderDefault.Font.Bold = true;
                    styleHeaderDefault.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    styleHeaderDefault.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    styleHeaderDefault.NumberFormat = "General";
                }
                //HeaderSize17
                if (!ExistStyle(StyleConstants.HeaderSize17))
                {
                    var styleHeaderSize17 = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.HeaderSize17, Type.Missing);
                    styleHeaderSize17.Font.Name = "MS Sans Serif";
                    styleHeaderSize17.Font.Size = 17;
                    styleHeaderSize17.Font.Bold = true;
                    styleHeaderSize17.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    styleHeaderSize17.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    styleHeaderSize17.NumberFormat = "General";
                }
                //TitleDefault
                if (!ExistStyle(StyleConstants.TitleDefault))
                {
                    var styleTitleDefault = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.TitleDefault, Type.Missing);
                    styleTitleDefault.Font.Name = "Calibri";
                    styleTitleDefault.Font.Size = 10;
                    styleTitleDefault.Font.Bold = true;
                    styleTitleDefault.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    styleTitleDefault.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    styleTitleDefault.NumberFormat = "General";
                }
                //TitleRequired
                if (!ExistStyle(StyleConstants.TitleRequired))
                {
                    var styleTitleRequired = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.TitleRequired, Type.Missing);
                    styleTitleRequired.Font.Name = "Calibri";
                    styleTitleRequired.Font.Size = 10;
                    styleTitleRequired.Font.Bold = true;
                    styleTitleRequired.Font.Color = ColorTranslator.ToOle(Color.White);
                    styleTitleRequired.NumberFormat = "General";
                    styleTitleRequired.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    styleTitleRequired.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    styleTitleRequired.Interior.Color = ColorTranslator.ToOle(Color.DarkBlue);
                }
                //TitleOptional
                if (!ExistStyle(StyleConstants.TitleOptional))
                {
                    var styleTitleOptional = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.TitleOptional, Type.Missing);
                    styleTitleOptional.Font.Name = "Calibri";
                    styleTitleOptional.Font.Size = 10;
                    styleTitleOptional.Font.Bold = true;
                    styleTitleOptional.Font.Color = ColorTranslator.ToOle(Color.White);
                    styleTitleOptional.NumberFormat = "General";
                    styleTitleOptional.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    styleTitleOptional.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    styleTitleOptional.Interior.Color = ColorTranslator.ToOle(Color.Black);
                }
                //TitleInformation
                if (!ExistStyle(StyleConstants.TitleInformation))
                {
                    var styleTitleInformation = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.TitleInformation, Type.Missing);
                    styleTitleInformation.Font.Name = "Calibri";
                    styleTitleInformation.Font.Size = 10;
                    styleTitleInformation.Font.Bold = true;
                    styleTitleInformation.Font.Color = ColorTranslator.ToOle(Color.Black);
                    styleTitleInformation.NumberFormat = "General";
                    styleTitleInformation.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    styleTitleInformation.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    styleTitleInformation.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                }
                //TitleAction
                if (!ExistStyle(StyleConstants.TitleAction))
                {
                    var styleTitleAction = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.TitleAction, Type.Missing);
                    styleTitleAction.Font.Name = "Calibri";
                    styleTitleAction.Font.Size = 10;
                    styleTitleAction.Font.Bold = true;
                    styleTitleAction.Font.Color = ColorTranslator.ToOle(Color.White);
                    styleTitleAction.NumberFormat = "General";
                    styleTitleAction.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    styleTitleAction.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    styleTitleAction.Interior.Color = ColorTranslator.ToOle(Color.Red);
                }
                //TitleAdditional
                if (!ExistStyle(StyleConstants.TitleAdditional))
                {
                    var styleTitleAdditional = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.TitleAdditional, Type.Missing);
                    styleTitleAdditional.Font.Name = "Calibri";
                    styleTitleAdditional.Font.Size = 10;
                    styleTitleAdditional.Font.Bold = true;
                    styleTitleAdditional.Font.Color = ColorTranslator.ToOle(Color.White);
                    styleTitleAdditional.NumberFormat = "General";
                    styleTitleAdditional.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    styleTitleAdditional.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    styleTitleAdditional.Interior.Color = ColorTranslator.ToOle(Color.LightGreen);
                }
                //TitleResult
                if (!ExistStyle(StyleConstants.TitleResult))
                {
                    var styleTitleResult = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.TitleResult, Type.Missing);
                    styleTitleResult.Font.Name = "Calibri";
                    styleTitleResult.Borders.Color = ColorTranslator.ToOle(Color.Gray);
                    styleTitleResult.Borders.LineStyle = XlLineStyle.xlContinuous;
                    styleTitleResult.Borders.Weight = 3d;
                    styleTitleResult.Borders[XlBordersIndex.xlDiagonalDown].LineStyle = XlLineStyle.xlLineStyleNone;
                    styleTitleResult.Borders[XlBordersIndex.xlDiagonalUp].LineStyle = XlLineStyle.xlLineStyleNone;
                    styleTitleResult.Font.Size = 10;
                    styleTitleResult.Font.Bold = true;
                    styleTitleResult.Font.Color = ColorTranslator.ToOle(Color.Black);
                    styleTitleResult.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    styleTitleResult.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    styleTitleResult.NumberFormat = "General";
                    styleTitleResult.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                }
                //Option
                if (!ExistStyle(StyleConstants.Option))
                {
                    var styleTitleOption = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.Option, Type.Missing);
                    styleTitleOption.Font.Name = "Calibri";
                    //styleTitleOption.Borders.Color = ColorTranslator.ToOle(Color.Black);
                    //styleTitleOption.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    //styleTitleOption.Borders.Weight = 3d;
                    //styleTitleOption.Borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    //styleTitleOption.Borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    //styleTitleOption.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 2d;
                    //styleTitleOption.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 2d;
                    //styleTitleOption.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = 2d;
                    //styleTitleOption.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 2d;
                    styleTitleOption.Font.Size = 10;
                    styleTitleOption.Font.Bold = true;
                    styleTitleOption.Font.Color = ColorTranslator.ToOle(Color.DarkBlue);
                    styleTitleOption.NumberFormat = "General";
                    styleTitleOption.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    styleTitleOption.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    styleTitleOption.Interior.Color = ColorTranslator.ToOle(Color.Gray);
                }
                //Select
                if (!ExistStyle(StyleConstants.Select))
                {
                    var styleTitleSelect = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.Select, Type.Missing);
                    styleTitleSelect.Font.Name = "Calibri";
                    styleTitleSelect.Borders.Color = ColorTranslator.ToOle(Color.Black);
                    styleTitleSelect.Borders.LineStyle = XlLineStyle.xlContinuous;
                    styleTitleSelect.Borders.Weight = 3d;
                    styleTitleSelect.Borders[XlBordersIndex.xlDiagonalDown].LineStyle = XlLineStyle.xlLineStyleNone;
                    styleTitleSelect.Borders[XlBordersIndex.xlDiagonalUp].LineStyle = XlLineStyle.xlLineStyleNone;
                    styleTitleSelect.Font.Size = 10;
                    styleTitleSelect.Font.Bold = true;
                    styleTitleSelect.Font.Color = ColorTranslator.ToOle(Color.Black);
                    styleTitleSelect.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    styleTitleSelect.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    styleTitleSelect.NumberFormat = "General";
                }
                //Disabled
                if (!ExistStyle(StyleConstants.Disabled))
                {
                    var styleTitleDisabled = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.Disabled, Type.Missing);
                    styleTitleDisabled.Font.Name = "Calibri";
                    styleTitleDisabled.Font.Size = 10;
                    styleTitleDisabled.Font.Bold = true;
                    styleTitleDisabled.Font.Color = ColorTranslator.ToOle(Color.Black);
                    styleTitleDisabled.NumberFormat = "General";
                    styleTitleDisabled.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    styleTitleDisabled.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    styleTitleDisabled.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                }
                //Time
                if (!ExistStyle(StyleConstants.Time))
                {
                    var styleTime = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.Time, Type.Missing);
                    styleTime.NumberFormat = "HH:mm:ss";
                }
                //ItalicSmall
                if (!ExistStyle(StyleConstants.ItalicSmall))
                {
                    var styleItalicSmall = _excelApp.ActiveWorkbook.Styles.Add(StyleConstants.ItalicSmall, Type.Missing);
                    styleItalicSmall.Font.Name = "Calibri";
                    styleItalicSmall.Font.Size = 8;
                    styleItalicSmall.Font.Italic = true;
                    styleItalicSmall.Font.Color = ColorTranslator.ToOle(Color.Black);
                    styleItalicSmall.NumberFormat = "General";
                    styleItalicSmall.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    styleItalicSmall.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    styleItalicSmall.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                }
            }
            catch(Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:createStyles()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                var styleConstantsList = StyleConstants.GetStyleListName();
                foreach (Style style in _excelApp.ActiveWorkbook.Styles)
                    if(styleConstantsList.Contains(style.Name))
                        style.Delete();
                CreateStyles();
            }
        }

        /// <summary>
        /// Indica si el estilo ya existe en el libro de excel
        /// </summary>
        /// <param name="styleName">Nombre del estilo a buscar</param>
        /// <returns>true si styleName existe, false si no existe un estilo con ese nombre</returns>
        private bool ExistStyle(string styleName)
        {
            foreach (Style style in _excelApp.ActiveWorkbook.Styles)
                if (style.Name.Equals(styleName))
                    return true;

            return false;
        }

        /// <summary>
        /// Adiciona una lista de validación a la celda o rango especificada
        /// </summary>
        /// <param name="targetRange">targetRange: Celda o rango para adicionarle la lista de validación</param>
        /// <param name="validationValues">List(string): lista de validación para adicionar al rango</param>
        /// <param name="showError">bool: Indica si muestra o no el diálogo de error al ingresar un valor erróneo</param>
        public void SetValidationList(Range targetRange, List<string> validationValues, bool showError = true)
        {
            string separator;
            //si uso los separadores del sistema
            if (_excelApp.UseSystemSeparators)
            {
                separator = LanguageSettingConstants.ListSeparator;
                //si el separador de lista y el separador decimal son iguales
                if (LanguageSettingConstants.ListSeparator.Equals(LanguageSettingConstants.DecimalSeparator))
                    separator = LanguageSettingConstants.DecimalSeparator.Equals(",") ? ";" : ",";
            }
            else
            {
                separator = _excelApp.DecimalSeparator.Equals(",") ? ";" : ",";
                
            }
            var list = string.Join(separator, validationValues);

            targetRange.Validation.Delete();
            targetRange.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, list, Type.Missing);
            targetRange.Validation.IgnoreBlank = true;
            targetRange.Validation.ShowError = true;
        }

        /// <summary>
        /// Adiciona una lista de validación a la celda o rango especificada
        /// </summary>
        /// <param name="targetRange">targetRange: Celda o rango para adicionarle la lista de validación</param>
        /// <param name="validationSheetName">string: Nombre de la hoja de datos de validación</param>
        /// <param name="validationColumnIndex">int: índice de la columna de datos de la hoja de validación dada</param>
        /// <param name="validationValues">List(string): lista de validación para adicionar al rango</param>
        /// <param name="showError">bool: Indica si muestra o no el diálogo de error al ingresar un valor erróneo</param>
        public void SetValidationList(Range targetRange, List<string> validationValues, string validationSheetName, int validationColumnIndex, bool showError = true)
        {
            Worksheet validationSheet = null;
            foreach (Worksheet sheet in _excelApp.ActiveWorkbook.Sheets)
            {
                if (sheet.Name != validationSheetName) continue;
                validationSheet = sheet;
                break;
            }
            if (validationSheet == null)
                throw new Exception(@"La hoja de validación ingresada no existe");
            
            var i = 1;

            var validCells = new ExcelStyleCells(_excelApp, validationSheetName);
            validCells.SetAlwaysActiveSheet(false);
            var rangeColumn = validCells.GetCell(validationColumnIndex, 1);
            rangeColumn = rangeColumn.EntireColumn;
            rangeColumn.Delete(Type.Missing);
            foreach(var value in validationValues)
            {
                validCells.GetCell(validationColumnIndex, i).Value = "'" + value;
                i++;
            }

            var columnName = GetExcelColumnName(validationColumnIndex);
            var formula = "='" + validationSheetName + "'!$"+columnName+":$"+columnName;
            targetRange.Validation.Delete();
            targetRange.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, formula, Type.Missing);
            targetRange.Validation.IgnoreBlank = true;
            targetRange.Validation.ShowError = showError;
        }

        /// <summary>
        /// Adiciona una lista de validación a la celda o rango especificada a partir de una columna de datos de una hoja dada
        /// </summary>
        /// <param name="targetRange">targetRange: Celda o rango para adicionarle la lista de validación</param>
        /// <param name="validationSheetName">string: Nombre de la hoja de datos de validación</param>
        /// <param name="validationColumnIndex">int: índice de la columna de datos de la hoja de validación dada</param>
        /// <param name="showError">bool: Indica si muestra o no el diálogo de error al ingresar un valor erróneo</param>
        public void SetValidationList(Range targetRange, string validationSheetName, int validationColumnIndex, bool showError = true)
        {
            Worksheet validationSheet = null;
            foreach (Worksheet sheet in _excelApp.ActiveWorkbook.Sheets)
            {
                if (sheet.Name != validationSheetName) continue;
                validationSheet = sheet;
                break;
            }

            if (validationSheet == null)
                throw new Exception(@"La hoja de validación ingresada no existe");

            var validCells = new ExcelStyleCells(_excelApp, validationSheetName);
            validCells.SetAlwaysActiveSheet(false);

            var columnName = GetExcelColumnName(validationColumnIndex);
            var formula = "='" + validationSheetName + "'!$" + columnName + ":$" + columnName;
            targetRange.Validation.Delete();
            targetRange.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, formula, Type.Missing);
            targetRange.Validation.IgnoreBlank = true;
            targetRange.Validation.ShowError = showError;
        }
        /// <summary>
        /// Adiciona una lista de validación a la hoja de validación especificada
        /// </summary>
        /// <param name="validationSheetName">string: Nombre de la hoja de datos de validación</param>
        /// <param name="validationColumnIndex">int: índice de la columna de datos de la hoja de validación dada</param>
        /// <param name="validationValues">List(string): lista de validación para adicionar al rango</param>
        public void SetValidationList(List<string> validationValues, string validationSheetName, int validationColumnIndex)
        {
            Worksheet validationSheet = null;
            foreach (Worksheet sheet in _excelApp.ActiveWorkbook.Sheets)
            {
                if (sheet.Name != validationSheetName) continue;
                validationSheet = sheet;
                break;
            }
            if (validationSheet == null)
                throw new Exception(@"La hoja de validación ingresada no existe");

            var i = 1;

            var validCells = new ExcelStyleCells(_excelApp, validationSheetName);
            validCells.SetAlwaysActiveSheet(false);
            var rangeColumn = validCells.GetCell(validationColumnIndex, 1);
            rangeColumn = rangeColumn.EntireColumn;
            rangeColumn.Delete(Type.Missing);
            foreach (var value in validationValues)
            {
                validCells.GetCell(validationColumnIndex, i).Value = "" + value;
                i++;
            }
        }
        /// <summary>
        /// Devuelve un string con el valor del objeto eliminando los espacios vacíos
        /// </summary>
        /// <param name="value">Object: objeto con el valor a obtener</param>
        /// <returns>string: Trim(value) o null si value es nulo. Si el valor está vacío o solo tiene espacios vacíos, devuelve un string vació</returns>
        public string GetNullOrTrimmedValue(object value)
        {
            return value == null ? null : Convert.ToString(value).Trim();
        }

        /// <summary>
        /// Devuelve un string con el valor de value eliminando los espacios vacíos, o null si value está vacío
        /// </summary>
        /// <param name="value">Object: objeto con el valor a obtener</param>
        /// <returns>string: Trim(value) o null si value es nulo. Si el valor está vacío o solo tiene espacios vacíos, devuelve null</returns>
        public string GetNullIfTrimmedEmpty(object value)
        {
            return string.IsNullOrWhiteSpace(Convert.ToString(value)) ? null : value.ToString().Trim();
        }

        /// <summary>
        /// Devuelve un string con el valor de value eliminando los espacios vacíos, o vacío si value es null
        /// </summary>
        /// <param name="value"></param>
        /// <returns>string: Trim(value) o vacío si value es nulo</returns>
        public string GetEmptyIfNull(object value)
        {
            return string.IsNullOrWhiteSpace(Convert.ToString(value)) ? "" : value.ToString().Trim();
        }

        

        /// <summary>
        /// Da formato a un rango especificado para que se comporte como una tabla en excel
        /// </summary>
        /// <param name="sourceRange">Range: Rango a formatear como tabla</param>
        /// <param name="tableName">string: Nombre dado a la tabla</param>
        /// <returns>ListObject: Tabla del Rango en forma de objeto ListObject</returns>
        public ListObject FormatAsTable(Range sourceRange, string tableName)
        {
            try//lo crea con un estilo predeterminado
            {
                sourceRange.Worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange,
                    sourceRange, Type.Missing, XlYesNoGuess.xlYes, Type.Missing).Name =
                    tableName;
                sourceRange.Worksheet.ListObjects[tableName].TableStyle =
                    StyleConstants.TableStyleConstants.DefaultTableStyle;
                return sourceRange.Worksheet.ListObjects[tableName]; //get table
            }
            catch (Exception)
            {
                sourceRange.Worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange,
                    sourceRange, Type.Missing, XlYesNoGuess.xlYes, Type.Missing).Name =
                    tableName;
                return sourceRange.Worksheet.ListObjects[tableName]; //get table
            }
        }
        public void DeleteTableRange(string tableRangeName)
        {
            try
            {
                var tableRange = GetRange(tableRangeName);

                tableRange.ListObject.Delete();
            }
            catch (Exception)
            {
                // ignored
            }
        }

        /// <summary>
        /// Elimina todas las filas de una tabla de rango dejando su encabezado
        /// </summary>
        /// <param name="tableRangeName">string: Nombre dado a la tabla rango</param>
        public void ClearTableRange(string tableRangeName)
        {
            try
            {
                var tableRange = GetRange(tableRangeName);

                var numberFormat = tableRange.ListObject.ListRows[1].Range;
                tableRange.EntireRow.Delete();
                ////para conservar el numberformat de la tabla
                tableRange.ListObject.ListRows[1].Range.Style = StyleConstants.Normal;
                tableRange.ListObject.ListRows[1].Range.NumberFormat = numberFormat;

            }
            catch (Exception)
            {
                // ignored
            }
        }

        /// <summary>
        /// Elimina el texto y formato de todas las celdas de una columna pertenecientes a una tabla rango
        /// </summary>
        /// <param name="tableRangeName">string: Nombre dado a la tabla rango</param>
        /// <param name="columnIndex">int: índice de la columna dentro de la tabla rango</param>
        public void ClearTableRangeColumn(string tableRangeName, int columnIndex)
        {
            try
            {
                var tableRange = GetRange(tableRangeName);
                GetRange(tableRange.ListObject.ListColumns[columnIndex].Range.Column,
                    tableRange.ListObject.ListColumns[columnIndex].Range.Row + 1,
                    tableRange.ListObject.ListColumns[columnIndex].Range.Column,
                    tableRange.ListObject.ListColumns[columnIndex].Range.Row + tableRange.ListObject.ListRows.Count)
                    .Clear();

                GetRange(tableRange.ListObject.ListColumns[columnIndex].Range.Column,
                    tableRange.ListObject.ListColumns[columnIndex].Range.Row + 1,
                    tableRange.ListObject.ListColumns[columnIndex].Range.Column,
                    tableRange.ListObject.ListColumns[columnIndex].Range.Row + tableRange.ListObject.ListRows.Count)
                    .Style = StyleConstants.Normal;
            }
            catch (Exception)
            {
                //ignored
            }
        }
        /// <summary>
        /// Elimina el texto y formato de todas las celdas de una columna pertenecientes a una tabla rango
        /// </summary>
        /// <param name="tableRangeName">string: Nombre dado a la tabla rango</param>
        /// <param name="columnName">string: Título del encabezado de la columna dentro de la tabla rango</param>
        public void ClearTableRangeColumn(string tableRangeName, string columnName)
        {
            try
            {
                var tableRange = GetRange(tableRangeName);
                GetRange(tableRange.ListObject.ListColumns[columnName].Range.Column,
                    tableRange.ListObject.ListColumns[columnName].Range.Row + 1,
                    tableRange.ListObject.ListColumns[columnName].Range.Column,
                    tableRange.ListObject.ListColumns[columnName].Range.Row + tableRange.ListObject.ListRows.Count)
                    .Clear();

                GetRange(tableRange.ListObject.ListColumns[columnName].Range.Column,
                    tableRange.ListObject.ListColumns[columnName].Range.Row + 1,
                    tableRange.ListObject.ListColumns[columnName].Range.Column,
                    tableRange.ListObject.ListColumns[columnName].Range.Row + tableRange.ListObject.ListRows.Count)
                    .Style = StyleConstants.Normal;
            }
            catch (Exception)
            {
                //ignored
            }

        }

        /// <summary>
        /// Obtiene la letra de una columna Excel según el índice ingresado (Ej. columnNumber: 8, resultado: H)
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        private static string GetExcelColumnName(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        /// <summary>
        /// Indica si el separador de decimales usado es punto.
        /// </summary>
        /// <returns>true: si el separador es punto ".", false si es otro símbolo</returns>
        public bool IsDecimalDotSeparator(Application excelApp = null)
        {
            if (excelApp == null)
                excelApp = _excelApp;
            //si uso los separadores del sistema
            var separator = excelApp.UseSystemSeparators 
                ? LanguageSettingConstants.DecimalSeparator : excelApp.DecimalSeparator;
            return separator.Equals(".");
        }

        public void SetCursorWait(Application excelApp = null)
        {
            if (excelApp == null)
                excelApp = _excelApp;
            excelApp.Cursor = XlMousePointer.xlWait;
        }

        public void SetCursorDefault(Application excelApp = null)
        {
            if (excelApp == null)
                excelApp = _excelApp;
            excelApp.Cursor = XlMousePointer.xlDefault;
        }

        public void SetEllipseDefaultCulture()
        {
            _oldCultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
        }

        public void ResetCurrentCulture()
        {
            if (_oldCultureInfo != null)
                System.Threading.Thread.CurrentThread.CurrentCulture = _oldCultureInfo;
        }
    }

    /// <summary>
    /// Estilos Predeterminados del Sistemas
    /// </summary>
    public static class StyleConstants
    {
        //PASOS PARA AGREGAR UN ESTILO A LA CLASE
        //1. Agregar la variable de constante en esta clase
        //2. Adicionarla a la lista del método GetStyleName
        //3. Adicionarla a CreateStyle

        public static string Normal = "MyNormal";
        public static string Success = "Success";
        public static string Warning = "Warning";
        public static string Error = "Error";
        public static string HeaderDefault = "HeaderDefault";
        public static string HeaderSize17 = "HeaderSize17";
        public static string TitleDefault = "TitleDefault";
        public static string TitleRequired = "TitleRequired";
        public static string TitleOptional = "TitleOptional";
        public static string TitleInformation = "TitleInformation";
        public static string TitleAction = "TitleAction";
        public static string TitleAdditional = "TitleAdditional";
        public static string TitleResult = "TitleResult";
        public static string Option = "Option";
        public static string Select = "Select";
        public static string Disabled = "Disabled";
        public static string Time = "Time";
        public static string ItalicSmall = "ItalicSmall";
        public static List<string> GetStyleListName()
        {
            var styleConstantsList = new List<string>
            {
                Normal, 
                Success, 
                Warning, 
                Error, 
                HeaderDefault, 
                HeaderSize17, 
                TitleDefault, 
                TitleRequired, 
                TitleOptional, 
                TitleInformation, 
                TitleAction, 
                TitleAdditional, 
                TitleResult, 
                Option, 
                Select, 
                Disabled, 
                Time, 
                ItalicSmall
            };

            return styleConstantsList;
        }

        public static class TableStyleConstants
        {
            public static string DefaultTableStyle = "TableStyleLight8";
        }
    }
    //Formatos de Número para Celdas del Sistema
    public static class NumberFormatConstants
    {
        public static string General = "General";
        public static string Text = "@";
        public static string Integer = "0";
    }

    public static class LanguageSettingConstants
    {
        public static string ListSeparator = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
        public static string DecimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;

    }

}
