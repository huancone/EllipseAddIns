using System;
using System.Collections.Generic;
using EllipseStdTextClassLibrary;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Utilities;

namespace EllipseReferenceCodesClassLibrary
{
    public static class ReferenceCodeActions
    {
        public static bool DebugErrors = false;//Muestra las alertas de debugger. Default: false


        public static List<ReferenceCodeEntity> FetchReferenceCodeEntities(EllipseFunctions ef, string entityType)
        {
            var sqlQuery = Queries.FetchReferenceCodeEntities(ef.DbReference, ef.DbLink, entityType);
            var drReference = ef.GetQueryResult(sqlQuery);
            var entityList = new List<ReferenceCodeEntity>();

            if (drReference == null || drReference.IsClosed) return entityList;
            while (drReference.Read())
            {
                var request = new ReferenceCodeEntity
                {
                    EntityType = drReference["ENTITY_TYPE"].ToString().Trim(),
                    RefNo = drReference["REF_NO"].ToString().Trim(),
                    RepeatCount = drReference["REPEAT_COUNT"].ToString().Trim(),
                    FieldType = drReference["FIELD_TYPE"].ToString().Trim(),
                    ShortName = drReference["SHORT_NAMES"].ToString().Trim(),
                    ScreenLiteral = drReference["SCREEN_LITERAL"].ToString().Trim(),
                    Length = Convert.ToInt16(drReference["LENGTH"].ToString().Trim()),
                    StdTextFlag = MyUtilities.IsTrue(drReference["STD_TEXT_FLAG"].ToString().Trim())
                };

                entityList.Add(request);
            }
            return entityList;
        }

        public static List<ReferenceCodeItem> FetchReferenceCodeItems(EllipseFunctions ef, string urlService, RefCodesService.OperationContext opContext, string entityType, string entityValue, string refNo = null)
        {
            var sqlQuery = Queries.FetchReferenceCodeItems(ef.DbReference, ef.DbLink, entityType, entityValue, refNo);
            var drReference = ef.GetQueryResult(sqlQuery);
            var itemList = new List<ReferenceCodeItem>();

            if (drReference == null || drReference.IsClosed) return itemList;
            while (drReference.Read())
            {
                var item = new ReferenceCodeItem
                {
                    EntityType = drReference["ENTITY_TYPE"].ToString().Trim(),
                    EntityValue = drReference["ENTITY_VALUE"].ToString().Trim(),
                    RefNo = drReference["REF_NO"].ToString().Trim(),
                    SeqNum = drReference["SEQ_NUM"].ToString().Trim(),
                    RefCode = drReference["REF_CODE"].ToString().Trim(),
                    FieldType = drReference["FIELD_TYPE"].ToString().Trim(),
                    ShortName = drReference["SHORT_NAMES"].ToString().Trim(),
                    ScreenLiteral = drReference["SCREEN_LITERAL"].ToString().Trim(),
                    StdtxtId = drReference["STD_TXT_KEY"].ToString().Trim(),
                    StdTextFlag = MyUtilities.IsTrue(drReference["STD_TEXT_FLAG"].ToString().Trim())
                };
                if (item.StdTextFlag && !string.IsNullOrWhiteSpace(item.StdtxtId))
                    item.StdText = StdText.GetText(urlService, StdText.GetStdTextOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings), "RC" + item.StdtxtId);
                itemList.Add(item);
            }
            return itemList;
        }

        public static ReferenceCodeItem FetchReferenceCodeItem(EllipseFunctions ef, string urlService, RefCodesService.OperationContext opContext, string entityType, string entityValue, string refNo, string seqNum)
        {
            var sqlQuery = Queries.FetchReferenceCodeItems(ef.DbReference, ef.DbLink, entityType, entityValue, refNo, seqNum);
            var drReference = ef.GetQueryResult(sqlQuery);

            if (drReference == null || drReference.IsClosed || !drReference.Read()) 
                return new ReferenceCodeItem();
            
            var item = new ReferenceCodeItem
            {
                EntityType = drReference["ENTITY_TYPE"].ToString().Trim(),
                EntityValue = drReference["ENTITY_VALUE"].ToString().Trim(),
                RefNo = drReference["REF_NO"].ToString().Trim(),
                SeqNum = drReference["SEQ_NUM"].ToString().Trim(),
                RefCode = drReference["REF_CODE"].ToString().Trim(),
                FieldType = drReference["FIELD_TYPE"].ToString().Trim(),
                ShortName = drReference["SHORT_NAMES"].ToString().Trim(),
                ScreenLiteral = drReference["SCREEN_LITERAL"].ToString().Trim(),
                StdtxtId = drReference["STD_TXT_KEY"].ToString().Trim(),
                StdTextFlag = MyUtilities.IsTrue(drReference["STD_TEXT_FLAG"].ToString().Trim())
            };

            try
            {
                if (item.StdTextFlag && !string.IsNullOrWhiteSpace(item.StdtxtId))

                    item.StdText = StdText.GetText(urlService, StdText.GetStdTextOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings), "RC" + item.StdtxtId);
            }
            catch (Exception ex)
            {
                throw new Exception("Error Fetching RefCode StdText. Entity " + item.EntityType + ", Value " + item.EntityValue + " , " + item.StdtxtId + ". " + ex.Message);
            }

            return item;
        }
        
        public static RefCodesService.RefCodesServiceModifyReplyDTO ModifyRefCode(EllipseFunctions ef, string urlService, RefCodesService.OperationContext opContext, ReferenceCodeItem refItem)
        {
            var proxySt = new RefCodesService.RefCodesService { Url = urlService + "/RefCodesService" };
            var request = new RefCodesService.RefCodesServiceModifyRequestDTO
            {
                entityType = refItem.EntityType,
                entityValue = refItem.EntityValue,
                refNo = refItem.RefNo,
                seqNum = refItem.SeqNum,
                refCode = refItem.RefCode,
                stdTxtKey = refItem.StdTextFlag ? refItem.StdtxtId : null,
            };
            try
            {
                var replyModify = proxySt.modify(opContext, request);
                
                //Actualizamos el StdText si existe
                try
                {
                    if (refItem.StdTextFlag && !string.IsNullOrWhiteSpace(refItem.StdtxtId)) //hay flag y se especifica el id
                        StdText.SetText(urlService, StdText.GetCustomOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings), "RC" + refItem.StdtxtId, refItem.StdText);
                    else if (refItem.StdTextFlag && string.IsNullOrWhiteSpace(refItem.StdtxtId)) //hay flag y NO se especifica el id
                    {
                        var item = FetchReferenceCodeItem(ef, urlService, opContext, refItem.EntityType, refItem.EntityValue, refItem.RefNo, refItem.SeqNum);
                        if (item.StdTextFlag && string.IsNullOrWhiteSpace(item.StdtxtId))
                            throw new Exception("Hay un error con el elemento. Especifica StdText pero no tiene un id asociado");
                        StdText.SetText(urlService, StdText.GetCustomOpContext(opContext.district, opContext.position, opContext.maxInstances, opContext.returnWarnings), "RC" + item.StdtxtId, refItem.StdText);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Error updating RefCode StdText. " + ex.Message);
                }

                return replyModify;
            }
            catch (Exception ex)
            {
                throw new Exception("Error updating RefCode. Entity " + refItem.EntityType + ", Value " + refItem.EntityValue + " , Ref No." + refItem.RefNo + ". " + ex.Message);
            }
        }


        /// <summary>
        /// Crea un nuevo operador de contexto para los métodos de la clase
        /// </summary>
        /// <param name="district">string: Distrito donde se va a crear el contexto</param>
        /// <param name="position">string: Posición donde se va a crear el contexto</param>
        /// <param name="maxInstances">int: Número máximo de instancias</param>
        /// <param name="returnWarnings">bool: True no ignora las advertencias</param>
        /// <returns></returns>
        public static RefCodesService.OperationContext GetRefCodesOpContext(string district, string position, int maxInstances, bool returnWarnings)
        {
            var opContext = new RefCodesService.OperationContext
            {
                district = district,
                position = position,
                maxInstances = maxInstances,
                maxInstancesSpecified = true,
                returnWarnings = returnWarnings,
                returnWarningsSpecified = true
            };

            return opContext;
        }

        public static RefCodesService.OperationContext GetRefCodesOpContext()
        {
            return new RefCodesService.OperationContext();
        }
    }

    
}
