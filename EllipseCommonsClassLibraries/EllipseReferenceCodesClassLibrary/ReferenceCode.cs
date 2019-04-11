using System;
using System.Collections.Generic;
using EllipseCommonsClassLibrary;
using EllipseStdTextClassLibrary;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseReferenceCodesClassLibrary
{
    public static class ReferenceCodeActions
    {
        public static bool DebugErrors = false;//Muestra las alertas de debugger. Default: false


        public static List<ReferenceCodeEntity> FetchReferenceCodeEntities(EllipseFunctions ef, string entityType)
        {
            var sqlQuery = Queries.FetchReferenceCodeEntities(ef.dbReference, ef.dbLink, entityType);
            var drReference = ef.GetQueryResult(sqlQuery);
            var entityList = new List<ReferenceCodeEntity>();

            if (drReference == null || drReference.IsClosed || !drReference.HasRows) return entityList;
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
            var sqlQuery = Queries.FetchReferenceCodeItems(ef.dbReference, ef.dbLink, entityType, entityValue, refNo);
            var drReference = ef.GetQueryResult(sqlQuery);
            var itemList = new List<ReferenceCodeItem>();

            if (drReference == null || drReference.IsClosed || !drReference.HasRows) return itemList;
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
            var sqlQuery = Queries.FetchReferenceCodeItems(ef.dbReference, ef.dbLink, entityType, entityValue, refNo, seqNum);
            var drReference = ef.GetQueryResult(sqlQuery);

            if (drReference == null || drReference.IsClosed || !drReference.HasRows) return new ReferenceCodeItem();
            drReference.Read();
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
                throw new Exception("Error updating RefCode StdText. Entity " + item.EntityType + ", Value " + item.EntityValue + " , " + item.StdtxtId + ". " + ex.Message);
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

    public static class Queries
    {
        public static string FetchReferenceCodeEntities(string dbReference, string dbLink, string entityType)
        {
            //escribimos el query
            var query = "" +
                        " SELECT RCE.ENTITY_TYPE," +
                        "     RCE.REF_NO," +
                        "     RCE.REPEAT_COUNT," +
                        "     RCE.FIELD_TYPE," +
                        "     RCE.SHORT_NAMES," +
                        "     RCE.SCREEN_LITERAL," +
                        "     RCE.LENGTH," +
                        "     RCE.STD_TEXT_FLAG" +
                        " FROM" +
                        "   " + dbReference + ".MSF070" + dbLink + " RCE" +
                        " WHERE RCE.ENTITY_TYPE = '" + entityType + "'";
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            return query;
        }

        public static string FetchReferenceCodeItems(string dbReference, string dbLink, string entityType, string entityValue, string refNo, string seqNum = null)
        {
            if (!string.IsNullOrWhiteSpace(refNo))
                refNo = " AND RC.REF_NO = '" + refNo.PadLeft(3, '0') + "'";
            if (!string.IsNullOrWhiteSpace(seqNum))
                seqNum = " AND RC.SEQ_NUM = '" + seqNum.PadLeft(3, '0') + "'";
            var query = "" +
                        " SELECT RC.ENTITY_TYPE, " +
                        "   RC.ENTITY_VALUE, " +
                        "   RC.REF_NO, " +
                        "   RC.SEQ_NUM, " +
                        "   RC.REF_CODE, " +
                        "   RCE.FIELD_TYPE, " +
                        "   RCE.SHORT_NAMES, " +
                        "   RCE.SCREEN_LITERAL, " +
                        "   RC.STD_TXT_KEY, " +
                        "   RCE.STD_TEXT_FLAG " +
                        " FROM " +
                        "     " + dbReference + ".MSF071" + dbLink + " RC LEFT JOIN " + dbReference + ".MSF070" + dbLink + " RCE " +
                        "         ON (RC.ENTITY_TYPE = RCE.ENTITY_TYPE AND RC.REF_NO = RCE.REF_NO) " +
                        " WHERE RCE.ENTITY_TYPE = '" + entityType + "' " +
                        " AND RC.ENTITY_VALUE = '" + entityValue + "' " +
                        " " + refNo +
                        " " + seqNum;
            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            return query;
        }


    }
}
