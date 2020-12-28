using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using EllipseWorkRequestClassLibrary.WorkRequestService;
using EllipseReferenceCodesClassLibrary;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Ellipse;

namespace EllipseWorkRequestClassLibrary
{

    public static class WorkRequestReferenceCodesActions
    {
        public static ReplyMessage CreateReferenceCodes(EllipseFunctions eFunctions, string urlService, OperationContext opContext, string requestId, WorkRequestReferenceCodes wrRefCodes)
        {
            //Corresponde a la misma acción de modificar, excepto que se garantiza que todos los RefCodes sean actualizados con la nueva información
            return ModifyReferenceCodes(eFunctions, urlService, opContext, requestId, wrRefCodes);
        }

        private static List<ReferenceCodeItem> GetNotNullRefCodeList(string entityType, string entityValue, WorkRequestReferenceCodes wrRefCodes)
        {
            var refItemList = new List<ReferenceCodeItem>();

            var riStockCode01 = new ReferenceCodeItem(entityType, entityValue, "001", "001", wrRefCodes.StockCode1, null, wrRefCodes.StockCode1Qty) { ShortName = "StockCode 01" };
            var riStockCode02 = new ReferenceCodeItem(entityType, entityValue, "001", "002", wrRefCodes.StockCode2, null, wrRefCodes.StockCode2Qty) { ShortName = "StockCode 02" };
            var riStockCode03 = new ReferenceCodeItem(entityType, entityValue, "001", "003", wrRefCodes.StockCode3, null, wrRefCodes.StockCode3Qty) { ShortName = "StockCode 03" };
            var riStockCode04 = new ReferenceCodeItem(entityType, entityValue, "001", "004", wrRefCodes.StockCode4, null, wrRefCodes.StockCode4Qty) { ShortName = "StockCode 04" };
            var riStockCode05 = new ReferenceCodeItem(entityType, entityValue, "001", "005", wrRefCodes.StockCode5, null, wrRefCodes.StockCode5Qty) { ShortName = "StockCode 05" };

            var riHorasHombre = new ReferenceCodeItem(entityType, entityValue, "006", "001", wrRefCodes.HorasQty, null, wrRefCodes.HorasHombre) { ShortName = "Horas Hombre" };
            var riDuracionTarea = new ReferenceCodeItem(entityType, entityValue, "007", "001", wrRefCodes.DuracionTarea) { ShortName = "Duracion Tarea" };
            var riEquipoDetenido = new ReferenceCodeItem(entityType, entityValue, "008", "001", wrRefCodes.EquipoDetenido) { ShortName = "Equipo Detenido" };
            var riWorkOrderOrigen = new ReferenceCodeItem(entityType, entityValue, "009", "001", wrRefCodes.WorkOrderOrigen) { ShortName = "OT de Inspección" };
            var riRaisedReprogramada = new ReferenceCodeItem(entityType, entityValue, "010", "001", wrRefCodes.RaisedReprogramada) { ShortName = "Raised Reprogramada" };
            var riCambioHora = new ReferenceCodeItem(entityType, entityValue, "011", "001", wrRefCodes.CambioHora) { ShortName = "Cambio Hora" };
            var riFechaPlanInicial = new ReferenceCodeItem(entityType, entityValue, "012", "001", wrRefCodes.FechaPlanInicial) { ShortName = "Fecha Plan Inicial" };
            var riFechaPlanFinal = new ReferenceCodeItem(entityType, entityValue, "013", "001", wrRefCodes.FechaPlanFinal) { ShortName = "Fecha Plan Final" };
            var riFechaEjecucionInicial = new ReferenceCodeItem(entityType, entityValue, "014", "001", wrRefCodes.FechaEjecucionInicial) { ShortName = "Fecha Ejecución Inicial" };
            var riFechaEjecucionFinal = new ReferenceCodeItem(entityType, entityValue, "015", "001", wrRefCodes.FechaEjecucionFinal) { ShortName = "Fecha Ejecución Final" };
            var riCalificacionEncuesta = new ReferenceCodeItem(entityType, entityValue, "016", "001", wrRefCodes.CalificacionEncuesta) { ShortName = "Calificación Encuesta" };
            var riWorkOrderReparacion = new ReferenceCodeItem(entityType, entityValue, "017", "001", wrRefCodes.WorkOrderReparacion) { ShortName = "OT de Reparación" };

            if (!(wrRefCodes.StockCode1 == null && wrRefCodes.StockCode1Qty == null))
                refItemList.Add(riStockCode01);
            if (!(wrRefCodes.StockCode2 == null && wrRefCodes.StockCode2Qty == null))
                refItemList.Add(riStockCode02);
            if (!(wrRefCodes.StockCode3 == null && wrRefCodes.StockCode3Qty == null))
                refItemList.Add(riStockCode03);
            if (!(wrRefCodes.StockCode4 == null && wrRefCodes.StockCode4Qty == null))
                refItemList.Add(riStockCode04);
            if (!(wrRefCodes.StockCode5 == null && wrRefCodes.StockCode5Qty == null))
                refItemList.Add(riStockCode05);

            if (wrRefCodes.HorasHombre != null || wrRefCodes.HorasQty != null)
                refItemList.Add(riHorasHombre);
            if (wrRefCodes.DuracionTarea != null)
                refItemList.Add(riDuracionTarea);
            if (wrRefCodes.EquipoDetenido != null)
                refItemList.Add(riEquipoDetenido);
            if (wrRefCodes.WorkOrderOrigen != null)
                refItemList.Add(riWorkOrderOrigen);
            if (wrRefCodes.RaisedReprogramada != null)
                refItemList.Add(riRaisedReprogramada);
            if (wrRefCodes.CambioHora != null)
                refItemList.Add(riCambioHora);
            if (wrRefCodes.FechaPlanInicial != null)
                refItemList.Add(riFechaPlanInicial);
            if (wrRefCodes.FechaPlanFinal != null)
                refItemList.Add(riFechaPlanFinal);
            if (wrRefCodes.FechaEjecucionInicial != null)
                refItemList.Add(riFechaEjecucionInicial);
            if (wrRefCodes.FechaEjecucionFinal != null)
                refItemList.Add(riFechaEjecucionFinal);
            if (wrRefCodes.CalificacionEncuesta != null)
                refItemList.Add(riCalificacionEncuesta);
            if (wrRefCodes.WorkOrderReparacion != null)
                refItemList.Add(riWorkOrderReparacion);

            return refItemList;
        }
        public static ReplyMessage ModifyReferenceCodes(EllipseFunctions eFunctions, string urlService, OperationContext opContext, string requestId, WorkRequestReferenceCodes wrRefCodes)
        {
            long defaultLong;
            if (long.TryParse(requestId, out defaultLong))
                requestId = requestId.PadLeft(12, '0');

            var refCodeOpContext = ReferenceCodeActions.GetRefCodesOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings);

            const string entityType = "WRQ";
            var entityValue = requestId;

            var reply = new ReplyMessage();
            var error = new List<string>();

            var refItemList = GetNotNullRefCodeList(entityType, entityValue, wrRefCodes);

            foreach (var item in refItemList)
            {
                try
                {
                    if (item.RefCode == null)
                        continue;
                    var replyRefCode = ReferenceCodeActions.ModifyRefCode(eFunctions, urlService, refCodeOpContext, item);
                    var stdTextId = replyRefCode.stdTxtKey;
                    if (string.IsNullOrWhiteSpace(stdTextId))
                        throw new Exception("No se recibió respuesta");
                }
                catch (Exception ex)
                {
                    error.Add("Error al actualizar " + item.ShortName + ": " + ex.Message);
                }
            }

            reply.Errors = error.ToArray();
            return reply;
        }
        [SuppressMessage("ReSharper", "InconsistentNaming")]
        public static WorkRequestReferenceCodes GetWorkRequestReferenceCodes(EllipseFunctions eFunctions, string urlService, OperationContext opContext, string requestId)
        {
            long defaultLong;
            if (long.TryParse(requestId, out defaultLong))
                requestId = requestId.PadLeft(12, '0');

            var wrRefCodes = new WorkRequestReferenceCodes();

            var rcOpContext = ReferenceCodeActions.GetRefCodesOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings);
            const string entityType = "WRQ";
            var entityValue = requestId;

            //Se encuentran problemas de implementación, debido a un comportamiento irregular del ODP en Windows. 
            //Las conexiones cerradas (EllipseFunctions.Close()) vuelven a la piscina (pool) de conexiones por un tiempo antes 
            //de ser completamente Cerradas (Close) y Dispuestas (Dispose), lo que ocasiona un desbordamiento del
            //número máximo de conexiones en el pool (100) y la nueva conexión alcanza el tiempo de espera (timeout) antes de
            //entrar en la cola del pool de conexiones arrojando un error 'Pooled Connection Request Timed Out'.
            //Para solucionarlo se fuerza el string de conexiones para que no genere una conexión que entre al pool.
            //Esto implica mayor tiempo de ejecución pero evita la excepción por el desbordamiento y tiempo de espera
            var newef = new EllipseFunctions(eFunctions);
            newef.SetConnectionPoolingType(false);

            var item001_01 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "001", "001");
            var item001_02 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "001", "002");
            var item001_03 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "001", "003");
            var item001_04 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "001", "004");
            var item001_05 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "001", "005");

            var item002 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "002", "001");
            var item006 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "006", "001");
            var item007 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "007", "001");
            var item008 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "008", "001");
            var item009 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "009", "001");
            var item010 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "010", "001");
            var item011 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "011", "001");

            var item012 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "012", "001");
            var item013 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "013", "001");
            var item014 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "014", "001");
            var item015 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "015", "001");

            var item016 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "016", "001");
            var item017 = ReferenceCodeActions.FetchReferenceCodeItem(newef, urlService, rcOpContext, entityType, entityValue, "017", "001");

            wrRefCodes.StockCode1 = item001_01.RefCode;            /*    wrRefCodes.StockCode1;             //001_9001  */
            wrRefCodes.StockCode1Qty = item001_01.StdText;         /*    wrRefCodes.StockCode1Qty;          //001_001   */
            wrRefCodes.StockCode2 = item001_02.RefCode;            /*    wrRefCodes.StockCode2;             //001_9002  */
            wrRefCodes.StockCode2Qty = item001_02.StdText;         /*    wrRefCodes.StockCode2Qty;          //001_002   */
            wrRefCodes.StockCode3 = item001_03.RefCode;            /*    wrRefCodes.StockCode3;             //001_9003  */
            wrRefCodes.StockCode3Qty = item001_03.StdText;         /*    wrRefCodes.StockCode3Qty;          //001_003   */
            wrRefCodes.StockCode4 = item001_04.RefCode;            /*    wrRefCodes.StockCode4;             //001_9004  */
            wrRefCodes.StockCode4Qty = item001_04.StdText;         /*    wrRefCodes.StockCode4Qty;          //001_004   */
            wrRefCodes.StockCode5 = item001_05.RefCode;            /*    wrRefCodes.StockCode5;             //001_9005  */
            wrRefCodes.StockCode5Qty = item001_05.StdText;         /*    wrRefCodes.StockCode5Qty;          //001_005   */
            wrRefCodes.NumeroComponente = item002.RefCode;         /*    wrRefCodes.NumeroComponente;       //002_001   */
            wrRefCodes.HorasHombre = item006.StdText;              /*    wrRefCodes.HorasHombre;            //006_9001  */
            wrRefCodes.HorasQty = item006.RefCode;                 /*    wrRefCodes.HorasQty;               //006_001   */
            wrRefCodes.DuracionTarea = item007.RefCode;            /*    wrRefCodes.DuracionTarea;          //007_001   */
            wrRefCodes.EquipoDetenido = item008.RefCode;           /*    wrRefCodes.EquipoDetenido;         //008_001   */
            wrRefCodes.WorkOrderOrigen = item009.RefCode;          /*    wrRefCodes.WorkOrderOrigen;        //009_001   */
            wrRefCodes.RaisedReprogramada = item010.RefCode;       /*    wrRefCodes.RaisedReprogramada;     //010_001   */
            wrRefCodes.CambioHora = item011.RefCode;               /*    wrRefCodes.CambioHora;             //011_001   */
            wrRefCodes.FechaPlanInicial = item012.RefCode;         /*    wrRefCodes.FechaPlanInicial;       //012_001   */
            wrRefCodes.FechaPlanFinal = item013.RefCode;           /*    wrRefCodes.FechaPlanFinal;         //013_001   */
            wrRefCodes.FechaEjecucionInicial = item014.RefCode;    /*    wrRefCodes.FechaEjecucionInicial;  //014_001   */
            wrRefCodes.FechaEjecucionFinal = item015.RefCode;      /*    wrRefCodes.FechaEjecucionFinal;    //015_001   */
            wrRefCodes.CalificacionEncuesta = item016.RefCode;     /*    wrRefCodes.CalificacionEncuesta;   //016_001   */
            wrRefCodes.WorkOrderReparacion = item017.RefCode;      /*    wrRefCodes.WorkOrderReparacion;    //017_001   */

            newef.CloseConnection();
            return wrRefCodes;
        }
    }

}
