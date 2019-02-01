using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseRequisitionServiceExcelAddIn
{
    public class RequisitionClassLibrary
    {
        public class RequisitionHeader
        {
            public string DistrictCode;
            public string IndSerie;
            public string IreqNo;
            public string IreqType;
            public string RequestedBy;
            public string RequiredByPos;
            public string IssTranType;
            public string OrigWhouseId;
            public string PriorityCode;

            public string CostDistrictA;
            public string WorkOrderA;
            public string WorkProjectIndA;//Solo para el MSO140
            public string EquipmentA;
            public string ProjectA;
            public string CostCentreA;

            public string DelivInstrA;
            public string DelivInstrB;

            public string AllocPcA;
            public string RequiredByDate;
            public string AnswerB;
            public string AnswerD;

            public bool PartIssue;
            public bool PartIssueSpecified;
            public bool ProtectedInd;
            public bool PickTaskReq;
            public bool ProtectedIndSpecified;

            /// <summary>
            /// Compara el objeto encabezado RequisitionHeader con otro encabezado. Devuelve true si el encabezado es igual a objectHeader
            /// </summary>
            /// <param name="objectHeader"></param>
            /// <returns>bool: true si objectHeader es igual</returns>
            public bool Equals(RequisitionHeader objectHeader)
            {
                return DistrictCode == objectHeader.DistrictCode &&
                       IndSerie == objectHeader.IndSerie &&
                       //IreqNo == objectHeader.IreqNo && //este no se debe comparar
                       IreqType == objectHeader.IreqType &&
                       RequestedBy == objectHeader.RequestedBy &&
                       RequiredByPos == objectHeader.RequiredByPos &&
                       IssTranType == objectHeader.IssTranType &&
                       OrigWhouseId == objectHeader.OrigWhouseId &&
                       PriorityCode == objectHeader.PriorityCode &&
                       CostDistrictA == objectHeader.CostDistrictA &&
                       WorkOrderA == objectHeader.WorkOrderA &&
                       EquipmentA == objectHeader.EquipmentA &&
                       ProjectA == objectHeader.ProjectA &&
                       CostCentreA == objectHeader.CostCentreA &&
                       DelivInstrA == objectHeader.DelivInstrA &&
                       DelivInstrB == objectHeader.DelivInstrB &&
                       AllocPcA == objectHeader.AllocPcA &&
                       RequiredByDate == objectHeader.RequiredByDate &&
                       AnswerB == objectHeader.AnswerB &&
                       AnswerD == objectHeader.AnswerD &&
                       PartIssue == objectHeader.PartIssue &&
                       ProtectedInd == objectHeader.ProtectedInd &&
                       PickTaskReq == objectHeader.PickTaskReq &&
                       ProtectedIndSpecified == objectHeader.ProtectedIndSpecified;
            }

            public RequisitionService.RequisitionServiceCreateHeaderRequestDTO GetCreateHeaderRequest()
            {
                var request = new RequisitionService.RequisitionServiceCreateHeaderRequestDTO
                {
                    districtCode = DistrictCode,
                    ireqNo = IreqNo,
                    ireqType = IreqType,
                    requestedBy = RequestedBy,
                    requiredByPos = RequiredByPos,
                    issTranType = IssTranType,
                    origWhouseId = OrigWhouseId,
                    priorityCode = PriorityCode,
                    costDistrictA = CostDistrictA,
                    workOrderA = GetNewWorkOrderDto(WorkOrderA),
                    equipmentA = EquipmentA,
                    projectA = ProjectA,
                    costCentreA = CostCentreA,
                    delivInstrA = DelivInstrA,
                    delivInstrB = DelivInstrB,
                    allocPcA = AllocPcA,
                    requiredByDate = RequiredByDate,
                    answerB = AnswerB,
                    answerD = AnswerD,
                    partIssue = PartIssue,
                    partIssueSpecified = PartIssueSpecified,
                    protectedInd = ProtectedInd,
                    pickTaskReq = PickTaskReq,
                    protectedIndSpecified = ProtectedIndSpecified
                };

                return request;
            }
            public RequisitionService.RequisitionServiceCreateHeaderReplyDTO GetCreateHeaderReply()
            {
                var request = new RequisitionService.RequisitionServiceCreateHeaderReplyDTO
                {
                    districtCode = DistrictCode,
                    ireqNo = IreqNo,
                    ireqType = IreqType,
                    requestedBy = RequestedBy,
                    requiredByPos = RequiredByPos,
                    issTranType = IssTranType,
                    origWhouseId = OrigWhouseId,
                    priorityCode = PriorityCode,
                    costDistrictA = CostDistrictA,
                    workOrderA = GetNewWorkOrderDto(WorkOrderA),
                    equipmentA = EquipmentA,
                    projectA = ProjectA,
                    costCentreA = CostCentreA,
                    delivInstrA = DelivInstrA,
                    delivInstrB = DelivInstrB,
                    allocPcA = AllocPcA,
                    requiredByDate = RequiredByDate,
                    answerB = AnswerB,
                    answerD = AnswerD,
                    partIssue = PartIssue,
                    partIssueSpecified = PartIssueSpecified,
                    protectedInd = ProtectedInd,
                    pickTaskReq = PickTaskReq,
                    protectedIndSpecified = ProtectedIndSpecified
                };

                return request;
            }

            /// <summary>
            /// Obtiene un nuevo objeto de tipo WorkOrderDTO a partir del número de la orden
            /// </summary>
            /// <param name="no">string: Número de la orden de trabajo</param>
            /// <returns>WorkOrderDTO</returns>
            public static RequisitionService.WorkOrderDTO GetNewWorkOrderDto(string no)
            {
                var workOrderDto = new RequisitionService.WorkOrderDTO();
                if (string.IsNullOrWhiteSpace(no)) return workOrderDto;

                no = no.Trim();
                if (no.Length < 3)
                    throw new Exception(@"El número de orden no corresponde a una orden válida");
                workOrderDto.prefix = no.Substring(0, 2);
                workOrderDto.no = no.Substring(2, no.Length - 2);
                return workOrderDto;
            }
            /// <summary>
            /// Obtiene un nuevo objeto de tipo WorkOrderDTO a partir del número de la orden
            /// </summary>
            /// <param name="prefix">string: prefijo de la orden de trabajo</param>
            /// <param name="no">string: Número de la orden de trabajo</param>
            /// <returns>WorkOrderDTO</returns>
            public static RequisitionService.WorkOrderDTO GetNewWorkOrderDto(string prefix, string no)
            {
                var workOrderDto = new RequisitionService.WorkOrderDTO
                {
                    prefix = prefix,
                    no = no
                };

                return workOrderDto;
            }

        }

        public class RequisitionItem
        {
            public int Index;
            public string StockCode;
            public string UnitOfMeasure;
            public string ItemType;
            public decimal QuantityRequired;
            public decimal IssueRequisitionItem;
            public bool QuantityRequiredSpecified;
            public bool PartIssue;
            public bool RepairRequest;
            public bool RepairRequestProtect;
            public bool IssueDocoFlg;
            public bool PurchDocoFlg;
            public bool NarrativeExists;
            public string AlterStockCodeFlg;
            public bool IssueRequisitionItemSpecified;
            public bool RepairRequestSpecified;
            public bool RepairRequestProtectSpecified;
            public bool PartIssueSpecified;
            public bool IssueDocoFlgSpecified;
            public bool PurchDocoFlgSpecified;
            public bool NarrativeExistsSpecified;
            public bool AlterStockCodeFlgSpecified;
            public bool DirectOrderIndicator;


            public RequisitionItem()
            {
                IssueRequisitionItem = 0;
                RepairRequest = false;
                RepairRequestProtect = false;
                PartIssue = false;
                IssueDocoFlg = false;
                PurchDocoFlg = false;
                NarrativeExists = false;

                QuantityRequiredSpecified = true;

                IssueRequisitionItemSpecified = false;
                RepairRequestSpecified = false;
                RepairRequestProtectSpecified = false;
                PartIssueSpecified = false;
                IssueDocoFlgSpecified = false;
                PurchDocoFlgSpecified = false;
                NarrativeExistsSpecified = false;
                AlterStockCodeFlgSpecified = false;
            }

            public RequisitionService.RequisitionItemDTO GetRequisitionItemDto()
            {
                var item = new RequisitionService.RequisitionItemDTO
                {
                    stockCode = StockCode,
                    unitOfMeasure = UnitOfMeasure,
                    itemType = ItemType,
                    quantityRequired = QuantityRequired,
                    issueRequisitionItem = IssueRequisitionItem,
                    quantityRequiredSpecified = QuantityRequiredSpecified,
                    partIssue = PartIssue,
                    repairRequest = RepairRequest,
                    repairRequestProtect = RepairRequestProtect,
                    issueDocoFlg = IssueDocoFlg,
                    purchDocoFlg = PurchDocoFlg,
                    narrativeExists = NarrativeExists,
                    alterStockCode = AlterStockCodeFlg,
                    issueRequisitionItemSpecified = IssueRequisitionItemSpecified,
                    repairRequestSpecified = RepairRequestSpecified,
                    repairRequestProtectSpecified = RepairRequestProtectSpecified,
                    partIssueSpecified = PartIssueSpecified,
                    issueDocoFlgSpecified = IssueDocoFlgSpecified,
                    purchDocoFlgSpecified = PurchDocoFlgSpecified,
                    narrativeExistsSpecified = NarrativeExistsSpecified,
                    alterStockCodeFlgSpecified = AlterStockCodeFlgSpecified
                };


                return item;
            }
        }

        public static class SpecialRestriction
        {
            public static List<SpecialRestrictionItem> GetPositionRestrictions(EllipseFunctions eFunctions)
            {
                var listItems = new List<SpecialRestrictionItem>();
                var query = GetSpecialRestrictionsQuery();
                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                var drItemCodes = eFunctions.GetQueryResult(query);

                if (drItemCodes == null || drItemCodes.IsClosed || !drItemCodes.HasRows) return listItems;
                while (drItemCodes.Read())
                {
                    var item = new SpecialRestrictionItem
                    {
                        Position = drItemCodes["POSITION"].ToString().Trim(),
                        Code = drItemCodes["TABLE_CODE"].ToString().Trim(),
                        MandatoryWorkOrder = MyUtilities.IsTrue(drItemCodes["WO_MANDATORY_FLAG"].ToString().Trim())
                    };
                    listItems.Add(item);
                }
                
                return listItems;
            }
            public class SpecialRestrictionItem
            {
                public string Position;
                public bool MandatoryWorkOrder;
                public string Code;
            }
            public static string GetSpecialRestrictionsQuery()
            {
                var query = "WITH PPP_TABLE AS" +
                            " (" +
                            "     SELECT" +
                            " SUBSTR(ASSOC_REC, 1, 1) WO_MANDATORY_FLAG," +
                            " SUBSTR(ASSOC_REC, 11, 10) POSITION_1," +
                            " SUBSTR(ASSOC_REC, 21, 10) POSITION_2," +
                            " SUBSTR(ASSOC_REC, 31, 10) POSITION_3," +
                            " SUBSTR(ASSOC_REC, 41, 10) POSITION_4," +
                            " TABLE_CODE" +
                            "     FROM ELLIPSE.MSF010" +
                            " WHERE TABLE_TYPE = '+PPP' AND ACTIVE_FLAG = 'Y' " +
                            "     )," +
                            " MODEL_FIL AS" +
                            " (" +
                            "     SELECT * FROM PPP_TABLE UNPIVOT(POSITION FOR COLUMNAS IN (POSITION_1, POSITION_2, POSITION_3, POSITION_4))" +
                            "     )" +
                            " SELECT WO_MANDATORY_FLAG, POSITION, TABLE_CODE" +
                            " FROM MODEL_FIL" +
                            " WHERE TRIM(POSITION) IS NOT NULL";
                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
                return query;
            }
        }
    }
}
