using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using EllipseCommonsClassLibrary;
using EllipseEquipmentClassLibrary.EquipmentService;
using EllipseEquipmentClassLibrary.EquipTraceService;

namespace EllipseEquipmentClassLibrary
{
    public class Equipment
    {
        public string AccountCode;
        public string ActiveFlag;
        public string ActiveFlagFieldSpecified;
        public string AssocEquipmentItemSwitch;
        public string AssocEquipmentItemSwitchFieldSpecified;
        public string CompCode;
        public string ConAstSegEn;
        public string ConAstSegEnFieldSpecified;
        public string ConAstSegSt;
        public string ConAstSegStFieldSpecified;
        public string ConditionRating;
        public string ConditionRatingFieldSpecified;
        public string ConditionStandard;
        public string CopyEquipment;
        public string CopyNameplateValues;
        public string CopyNameplateValuesFieldSpecified;
        public string CostSegLgth;
        public string CostSegLgthFieldSpecified;
        public string CostingFlag;
        public string CtaxCode;
        public string Custodian;
        public string CustodianPosition;
        //public string CustomerName;
        public string CustomerNumber;
        public string DistrictCode;
        public string DrawingNo;
        public string EquipmentClass;
        public string EquipmentCriticality;
        public string EquipmentGrpId;
        public string EquipmentLocation;
        public string EquipmentNo;
        public string EquipmentNoDescription1;
        public string EquipmentNoDescription2;
        public string EquipmentRef;
        public string EquipmentStatus;
        public string EquipmentType;
        public string EquipmentTypeDescription;
        public string ExpElement;
        public string IaaAssetInd;
        public string IaaAssetIndFieldSpecified;
        public string InputBy;
        public string ItemNameCode;
        public string LatestConditionDate;
        //public string Location;
        public string Mnemonic;
        public string MsssFlag;
        public string OperatingStandard;
        public string OperatorId;
        public string OperatorPosition;
        public string OriginalDoc;
        public string ParentEquipment;
        public string ParentEquipmentRef;
        public string PartNo;
        public string PermitReqdSw;
        public string PermitReqdSwFieldSpecified;
        //public string PlantCode0;
        //public string PlantCode1;
        //public string PlantCode2;
        //public string PlantCode3;
        //public string PlantCode4;
        //public string PlantCode5;
        //public string PlantCodes;
        //public string PlantNames;
        public string PlantNo;
        public string PoNo;
        public string PrimaryFunction;
        public string ProdUnitItem;
        public string PurchaseDate;
        public string PurchasePrice;
        public string PurchasePriceFieldSpecified;
        public string RcmAnalysisSw;
        public string RcmAnalysisSwFieldSpecified;
        public string ReplaceValue;
        public string ReplaceValueFieldSpecified;
        public string SegmentUom;
        public string SerialNumber;
        public string ShutdownEquipment;
        public string StockCode;
        public string TaxCode;
        public string TraceableFlg;
        public string TraceableFlgFieldSpecified;
        public string ValuationDate;
        public string WarrStatType;
        public string WarrStatVal;
        public string WarrStatValFieldSpecified;
        public string WarrantyDate;

        public LinkOneBook LinkOne;
        public ClassificationCodes ClassCodes;
        public class LinkOneBook
        {
            public string Publisher;
            public string Book;
            public string PageReference;
            public string ItemId;
            public string WorkOrder;
        }

        public class ClassificationCodes
        {
            public string EquipmentClassif;
            public string EquipmentClassif0;
            public string EquipmentClassif1;
            public string EquipmentClassif2;
            public string EquipmentClassif3;
            public string EquipmentClassif4;
            public string EquipmentClassif5;
            public string EquipmentClassif6;
            public string EquipmentClassif7;
            public string EquipmentClassif8;
            public string EquipmentClassif9;
            public string EquipmentClassif10;
            public string EquipmentClassif11;
            public string EquipmentClassif12;
            public string EquipmentClassif13;
            public string EquipmentClassif14;
            public string EquipmentClassif15;
            public string EquipmentClassif16;
            public string EquipmentClassif17;
            public string EquipmentClassif18;
            public string EquipmentClassif19;
        }
    }

    public static class EquipmentActions
    {
        /// <summary>
        /// Obtiene el listado de equipos que coincidan con la referencia de equipo ingresada
        /// </summary>
        /// <param name="ef"></param>
        /// <param name="districtCode"></param>
        /// <param name="equipmentRef"></param>
        /// <returns></returns>
        public static List<string> GetEquipmentList(EllipseFunctions ef, string districtCode, string equipmentRef)
        {
            var equipmentList = new List<string>();

            var drEquipments = ef.GetQueryResult(Queries.GetEquipReferencesQuery(ef.dbReference, ef.dbLink, districtCode, equipmentRef));
            // ReSharper disable once InvertIf
            if (drEquipments != null && !drEquipments.IsClosed && drEquipments.HasRows)
                while (drEquipments.Read())
                    equipmentList.Add(drEquipments["EQUIP_NO"].ToString().Trim());
            return equipmentList;
        }

        /// <summary>
        /// Obtiene el listado de equipos pertenecientes a un EGI dado
        /// </summary>
        /// <param name="ef"></param>
        /// <param name="egi"></param>
        /// <returns></returns>
        public static List<string> GetEgiEquipments(EllipseFunctions ef, string egi)
        {
            var equipmentList = new List<string>();

            var drEquipments = ef.GetQueryResult(Queries.GetEgiEquipmentsQuery(ef.dbReference, ef.dbLink, egi));
            // ReSharper disable once InvertIf
            if (drEquipments != null && !drEquipments.IsClosed && drEquipments.HasRows)
                while (drEquipments.Read())
                    equipmentList.Add(drEquipments["EQUIP_NO"].ToString().Trim());
            return equipmentList;
        }

        /// <summary>
        /// Obtiene el listado de equipo pertenecientes a una lista (tipo lista, id lista) dada
        /// </summary>
        /// <param name="ef"></param>
        /// <param name="listType"></param>
        /// <param name="listId"></param>
        /// <returns></returns>
        public static List<string> GetListEquipments(EllipseFunctions ef, string listType, string listId)
        {
            var equipmentList = new List<string>();

            var drEquipments = ef.GetQueryResult(Queries.GetListEquipmentsQuery(ef.dbReference, ef.dbLink, listType, listId));
            // ReSharper disable once InvertIf
            if (drEquipments != null && !drEquipments.IsClosed && drEquipments.HasRows)
                while (drEquipments.Read())
                    equipmentList.Add(drEquipments["EQUIP_NO"].ToString().Trim());
            return equipmentList;
        }


        /// <summary>
        /// Obtiene el listado de equipos pertenecientes a una unidad productiva
        /// </summary>
        /// <param name="ef"></param>
        /// <param name="district"></param>
        /// <param name="productiveUnit"></param>
        /// <returns></returns>
        public static List<string> GetProductiveUnitEquipments(EllipseFunctions ef, string district, string productiveUnit)
        {
            var equipmentList = new List<string>();

            var drEquipments = ef.GetQueryResult(Queries.GetProductiveUnitEquipmentsQuery(ef.dbReference, ef.dbLink, district, productiveUnit));
            // ReSharper disable once InvertIf
            if (drEquipments != null && !drEquipments.IsClosed && drEquipments.HasRows)
                while (drEquipments.Read())
                    equipmentList.Add(drEquipments["EQUIP_NO"].ToString().Trim());
            return equipmentList;
        }
        /// <summary>
        /// Actualiza el estado del equipo a un estado especificado
        /// </summary>
        /// <param name="operationContext"></param>
        /// <param name="urlService"></param>
        /// <param name="equipmentNo"></param>
        /// <param name="status"></param>
        /// <returns></returns>
        public static string UpdateEquipmentStatus(EquipmentService.OperationContext operationContext, string urlService, string equipmentNo, string status)
        {
            var proxyEquip = new EquipmentService.EquipmentService();
            var request = new EquipmentServiceModifyRequestDTO
            {
                equipmentNo = equipmentNo,
                equipmentStatus = status
            };
            proxyEquip.Url = urlService + "/Equipment";

            var reply = proxyEquip.modify(operationContext, request);
            return reply.equipmentStatus;
        }
        public static List<Equipment> FetchEquipmentDataList(EllipseFunctions ef, string district, int primakeryKey, string primaryValue, int secondarykey, string secondaryValue, string eqStatus)
        {
            var sqlQuery = Queries.GetFetchEquipmentDataQuery(ef.dbReference, ef.dbLink, district, primakeryKey, primaryValue, secondarykey, secondaryValue, eqStatus);
            var drEquipments = ef.GetQueryResult(sqlQuery);
            var list = new List<Equipment>();

            if (drEquipments == null || drEquipments.IsClosed || !drEquipments.HasRows) return list;
            while (drEquipments.Read())
            {
                var equipment = new Equipment
                {
                    EquipmentNo = drEquipments["EQUIP_NO"].ToString().Trim(),
                    AccountCode = drEquipments["ACCOUNT_CODE"].ToString().Trim(),
                    ActiveFlag = drEquipments["ACTIVE_FLG"].ToString().Trim(),
                    AssocEquipmentItemSwitch = drEquipments["ASSOC_EQUIP_SW"].ToString().Trim(),
                    CompCode = drEquipments["COMP_CODE"].ToString().Trim(),
                    ConAstSegEn = drEquipments["CON_AST_SEG_EN"].ToString().Trim(),
                    ConAstSegSt = drEquipments["CON_AST_SEG_ST"].ToString().Trim(),
                    ConditionRating = drEquipments["COND_RATING"].ToString().Trim(),
                    ConditionStandard = drEquipments["COND_STANDARD"].ToString().Trim(),
                    CostSegLgth = drEquipments["COST_SEG_LGTH"].ToString().Trim(),
                    CostingFlag = drEquipments["COSTING_FLG"].ToString().Trim(),
                    CtaxCode = drEquipments["CTAX_CODE"].ToString().Trim(),
                    Custodian = drEquipments["CUSTODIAN"].ToString().Trim(),
                    CustodianPosition = drEquipments["CUSTODIAN_POSN"].ToString().Trim(),
                    CustomerNumber = drEquipments["CUST_NO"].ToString().Trim(),
                    DistrictCode = drEquipments["DSTRCT_CODE"].ToString().Trim(),
                    DrawingNo = drEquipments["DRAWING_NO"].ToString().Trim(),
                    EquipmentClass = drEquipments["EQUIP_CLASS"].ToString().Trim(),
                    EquipmentCriticality = drEquipments["EQUIP_CRITICALITY"].ToString().Trim(),
                    EquipmentGrpId = drEquipments["EQUIP_GRP_ID"].ToString().Trim(),
                    EquipmentLocation = drEquipments["EQUIP_LOCATION"].ToString().Trim(),
                    EquipmentNoDescription1 = drEquipments["ITEM_NAME_1"].ToString().Trim(),
                    EquipmentNoDescription2 = drEquipments["ITEM_NAME_2"].ToString().Trim(),
                    EquipmentRef = drEquipments["EQUIP_REF"].ToString().Trim(),
                    EquipmentStatus = drEquipments["EQUIP_STATUS"].ToString().Trim(),
                    EquipmentType = drEquipments["EQPT_TYPE"].ToString().Trim(),
                    ExpElement = drEquipments["EXP_ELEMENT"].ToString().Trim(),
                    IaaAssetInd = drEquipments["IAA_ASSET_IND"].ToString().Trim(),
                    InputBy = drEquipments["INPUT_BY"].ToString().Trim(),
                    ItemNameCode = drEquipments["ITEM_NAME_CODE"].ToString().Trim(),
                    LatestConditionDate = drEquipments["LATEST_COND_DATE"].ToString().Trim(),
                    //Location = drEquipments["LOCATION"].ToString().Trim(),
                    Mnemonic = drEquipments["MNEMONIC"].ToString().Trim(),
                    MsssFlag = drEquipments["MSSS_STATUS_IND"].ToString().Trim(),
                    OperatingStandard = drEquipments["OPERATING_STD"].ToString().Trim(),
                    OperatorId = drEquipments["OPERATOR_ID"].ToString().Trim(),
                    OperatorPosition = drEquipments["OPERATOR_POSN"].ToString().Trim(),
                    OriginalDoc = drEquipments["ORIGINAL_DOC"].ToString().Trim(),
                    ParentEquipment = drEquipments["PARENT_EQUIP"].ToString().Trim(),
                    ParentEquipmentRef = drEquipments["PARENT_EQUIP_REF"].ToString().Trim(),
                    PartNo = drEquipments["PART_NO"].ToString().Trim(),
                    PermitReqdSw = drEquipments["PERMIT_REQD_SW"].ToString().Trim(),
                    PlantNo = drEquipments["PLANT_NO"].ToString().Trim(),
                    PoNo = drEquipments["PO_NO"].ToString().Trim(),
                    PrimaryFunction = drEquipments["PRIMARY_FUNCTION"].ToString().Trim(),
                    ProdUnitItem = drEquipments["PROD_UNIT_ITEM"].ToString().Trim(),
                    PurchaseDate = drEquipments["PURCHASE_DATE"].ToString().Trim(),
                    PurchasePrice = drEquipments["PURCHASE_PRICE"].ToString().Trim(),
                    RcmAnalysisSw = drEquipments["RCM_ANALYSIS_SW"].ToString().Trim(),
                    ReplaceValue = drEquipments["REPLACE_VALUE"].ToString().Trim(),
                    ValuationDate = drEquipments["VALUATION_DATE"].ToString().Trim(),
                    WarrStatType = drEquipments["WARR_STAT_TYPE"].ToString().Trim(),
                    WarrStatVal = drEquipments["WARR_STAT_VAL"].ToString().Trim(),
                    ClassCodes = new Equipment.ClassificationCodes
                    {
                        EquipmentClassif = drEquipments["EQUIP_CLASS"].ToString().Trim(),
                        EquipmentClassif0 = drEquipments["EQUIP_CLASSIFX1"].ToString().Trim(),
                        EquipmentClassif1 = drEquipments["EQUIP_CLASSIFX2"].ToString().Trim(),
                        EquipmentClassif2 = drEquipments["EQUIP_CLASSIFX3"].ToString().Trim(),
                        EquipmentClassif3 = drEquipments["EQUIP_CLASSIFX4"].ToString().Trim(),
                        EquipmentClassif4 = drEquipments["EQUIP_CLASSIFX5"].ToString().Trim(),
                        EquipmentClassif5 = drEquipments["EQUIP_CLASSIFX6"].ToString().Trim(),
                        EquipmentClassif6 = drEquipments["EQUIP_CLASSIFX7"].ToString().Trim(),
                        EquipmentClassif7 = drEquipments["EQUIP_CLASSIFX8"].ToString().Trim(),
                        EquipmentClassif8 = drEquipments["EQUIP_CLASSIFX9"].ToString().Trim(),
                        EquipmentClassif9 = drEquipments["EQUIP_CLASSIFX10"].ToString().Trim(),
                        EquipmentClassif10 = drEquipments["EQUIP_CLASSIFX11"].ToString().Trim(),
                        EquipmentClassif11 = drEquipments["EQUIP_CLASSIFX12"].ToString().Trim(),
                        EquipmentClassif12 = drEquipments["EQUIP_CLASSIFX13"].ToString().Trim(),
                        EquipmentClassif13 = drEquipments["EQUIP_CLASSIFX14"].ToString().Trim(),
                        EquipmentClassif14 = drEquipments["EQUIP_CLASSIFX15"].ToString().Trim(),
                        EquipmentClassif15 = drEquipments["EQUIP_CLASSIFX16"].ToString().Trim(),
                        EquipmentClassif16 = drEquipments["EQUIP_CLASSIFX17"].ToString().Trim(),
                        EquipmentClassif17 = drEquipments["EQUIP_CLASSIFX18"].ToString().Trim(),
                        EquipmentClassif18 = drEquipments["EQUIP_CLASSIFX19"].ToString().Trim(),
                        EquipmentClassif19 = drEquipments["EQUIP_CLASSIFX20"].ToString().Trim()
                    },
                    LinkOne = new Equipment.LinkOneBook()

                };
                list.Add(equipment);
            }

            return list;
        }
        public static Equipment FetchEquipmentData(EllipseFunctions ef, string equipmentNo)
        {
            var sqlQuery = Queries.GetFetchEquipmentDataQuery(ef.dbReference, ef.dbLink, equipmentNo);
            var drEquipments = ef.GetQueryResult(sqlQuery);

            if (drEquipments == null || drEquipments.IsClosed || !drEquipments.HasRows) return null;
            drEquipments.Read();
            var equipment = new Equipment
            {
                EquipmentNo = drEquipments["EQUIP_NO"].ToString().Trim(),
                AccountCode = drEquipments["ACCOUNT_CODE"].ToString().Trim(),
                ActiveFlag = drEquipments["ACTIVE_FLG"].ToString().Trim(),
                AssocEquipmentItemSwitch = drEquipments["ASSOC_EQUIP_SW"].ToString().Trim(),
                CompCode = drEquipments["COMP_CODE"].ToString().Trim(),
                ConAstSegEn = drEquipments["CON_AST_SEG_EN"].ToString().Trim(),
                ConAstSegSt = drEquipments["CON_AST_SEG_ST"].ToString().Trim(),
                ConditionRating = drEquipments["COND_RATING"].ToString().Trim(),
                ConditionStandard = drEquipments["COND_STANDARD"].ToString().Trim(),
                CostSegLgth = drEquipments["COST_SEG_LGTH"].ToString().Trim(),
                CostingFlag = drEquipments["COSTING_FLG"].ToString().Trim(),
                CtaxCode = drEquipments["CTAX_CODE"].ToString().Trim(),
                Custodian = drEquipments["CUSTODIAN"].ToString().Trim(),
                CustodianPosition = drEquipments["CUSTODIAN_POSN"].ToString().Trim(),
                CustomerNumber = drEquipments["CUST_NO"].ToString().Trim(),
                DistrictCode = drEquipments["DSTRCT_CODE"].ToString().Trim(),
                DrawingNo = drEquipments["DRAWING_NO"].ToString().Trim(),
                EquipmentClass = drEquipments["EQUIP_CLASS"].ToString().Trim(),
                EquipmentCriticality = drEquipments["EQUIP_CRITICALITY"].ToString().Trim(),
                EquipmentGrpId = drEquipments["EQUIP_GRP_ID"].ToString().Trim(),
                EquipmentLocation = drEquipments["EQUIP_LOCATION"].ToString().Trim(),
                EquipmentNoDescription1 = drEquipments["ITEM_NAME_1"].ToString().Trim(),
                EquipmentNoDescription2 = drEquipments["ITEM_NAME_2"].ToString().Trim(),
                EquipmentRef = drEquipments["EQUIP_REF"].ToString().Trim(),
                EquipmentStatus = drEquipments["EQUIP_STATUS"].ToString().Trim(),
                EquipmentType = drEquipments["EQPT_TYPE"].ToString().Trim(),
                ExpElement = drEquipments["EXP_ELEMENT"].ToString().Trim(),
                IaaAssetInd = drEquipments["IAA_ASSET_IND"].ToString().Trim(),
                InputBy = drEquipments["INPUT_BY"].ToString().Trim(),
                ItemNameCode = drEquipments["ITEM_NAME_CODE"].ToString().Trim(),
                LatestConditionDate = drEquipments["LATEST_COND_DATE"].ToString().Trim(),
                //Location = drEquipments["LOCATION"].ToString().Trim(),
                Mnemonic = drEquipments["MNEMONIC"].ToString().Trim(),
                MsssFlag = drEquipments["MSSS_STATUS_IND"].ToString().Trim(),
                OperatingStandard = drEquipments["OPERATING_STD"].ToString().Trim(),
                OperatorId = drEquipments["OPERATOR_ID"].ToString().Trim(),
                OperatorPosition = drEquipments["OPERATOR_POSN"].ToString().Trim(),
                OriginalDoc = drEquipments["ORIGINAL_DOC"].ToString().Trim(),
                ParentEquipment = drEquipments["PARENT_EQUIP"].ToString().Trim(),
                ParentEquipmentRef = drEquipments["PARENT_EQUIP_REF"].ToString().Trim(),
                PartNo = drEquipments["PART_NO"].ToString().Trim(),
                PermitReqdSw = drEquipments["PERMIT_REQD_SW"].ToString().Trim(),
                PlantNo = drEquipments["PLANT_NO"].ToString().Trim(),
                PoNo = drEquipments["PO_NO"].ToString().Trim(),
                PrimaryFunction = drEquipments["PRIMARY_FUNCTION"].ToString().Trim(),
                ProdUnitItem = drEquipments["PROD_UNIT_ITEM"].ToString().Trim(),
                PurchaseDate = drEquipments["PURCHASE_DATE"].ToString().Trim(),
                PurchasePrice = drEquipments["PURCHASE_PRICE"].ToString().Trim(),
                RcmAnalysisSw = drEquipments["RCM_ANALYSIS_SW"].ToString().Trim(),
                ReplaceValue = drEquipments["REPLACE_VALUE"].ToString().Trim(),
                ValuationDate = drEquipments["VALUATION_DATE"].ToString().Trim(),
                WarrStatType = drEquipments["WARR_STAT_TYPE"].ToString().Trim(),
                WarrStatVal = drEquipments["WARR_STAT_VAL"].ToString().Trim(),
                ClassCodes = new Equipment.ClassificationCodes
                {
                    EquipmentClassif = drEquipments["EQUIP_CLASS"].ToString().Trim(),
                    EquipmentClassif0 = drEquipments["EQUIP_CLASSIFX1"].ToString().Trim(),
                    EquipmentClassif1 = drEquipments["EQUIP_CLASSIFX2"].ToString().Trim(),
                    EquipmentClassif2 = drEquipments["EQUIP_CLASSIFX3"].ToString().Trim(),
                    EquipmentClassif3 = drEquipments["EQUIP_CLASSIFX4"].ToString().Trim(),
                    EquipmentClassif4 = drEquipments["EQUIP_CLASSIFX5"].ToString().Trim(),
                    EquipmentClassif5 = drEquipments["EQUIP_CLASSIFX6"].ToString().Trim(),
                    EquipmentClassif6 = drEquipments["EQUIP_CLASSIFX7"].ToString().Trim(),
                    EquipmentClassif7 = drEquipments["EQUIP_CLASSIFX8"].ToString().Trim(),
                    EquipmentClassif8 = drEquipments["EQUIP_CLASSIFX9"].ToString().Trim(),
                    EquipmentClassif9 = drEquipments["EQUIP_CLASSIFX10"].ToString().Trim(),
                    EquipmentClassif10 = drEquipments["EQUIP_CLASSIFX11"].ToString().Trim(),
                    EquipmentClassif11 = drEquipments["EQUIP_CLASSIFX12"].ToString().Trim(),
                    EquipmentClassif12 = drEquipments["EQUIP_CLASSIFX13"].ToString().Trim(),
                    EquipmentClassif13 = drEquipments["EQUIP_CLASSIFX14"].ToString().Trim(),
                    EquipmentClassif14 = drEquipments["EQUIP_CLASSIFX15"].ToString().Trim(),
                    EquipmentClassif15 = drEquipments["EQUIP_CLASSIFX16"].ToString().Trim(),
                    EquipmentClassif16 = drEquipments["EQUIP_CLASSIFX17"].ToString().Trim(),
                    EquipmentClassif17 = drEquipments["EQUIP_CLASSIFX18"].ToString().Trim(),
                    EquipmentClassif18 = drEquipments["EQUIP_CLASSIFX19"].ToString().Trim(),
                    EquipmentClassif19 = drEquipments["EQUIP_CLASSIFX20"].ToString().Trim()
                },
                LinkOne = new Equipment.LinkOneBook()

            };

            return equipment;
        }
        /// <summary>
        /// Elimina un equipo especificado
        /// </summary>
        /// <param name="operationContext"></param>
        /// <param name="urlService"></param>
        /// <param name="equipment"></param>
        /// <returns></returns>
        public static EquipmentServiceDeleteReplyDTO DeleteEquipment(EquipmentService.OperationContext operationContext, string urlService, Equipment equipment)
        {
            var proxyEquip = new EquipmentService.EquipmentService();
            var request = new EquipmentServiceDeleteRequestDTO
            {
                equipmentNo = equipment.EquipmentNo,
                districtCode = equipment.DistrictCode
            };
            proxyEquip.Url = urlService + "/Equipment";

            return proxyEquip.delete(operationContext, request);
        }
        /// <summary>
        /// Crea un nuevo equipo con la información especificada
        /// </summary>
        /// <param name="operationContext"></param>
        /// <param name="urlService"></param>
        /// <param name="equipment"></param>
        /// <returns></returns>
        public static EquipmentServiceCreateReplyDTO CreateEquipment(EquipmentService.OperationContext operationContext, string urlService, Equipment equipment)
        {
            var proxyEquip = new EquipmentService.EquipmentService();
            var request = new EquipmentServiceCreateRequestDTO
            {
                accountCode = equipment.AccountCode,
                activeFlag = Utils.IsTrue(equipment.ActiveFlag),
                activeFlagSpecified = equipment.ActiveFlag != null,
                assocEquipmentItemSwitch = Utils.IsTrue(equipment.AssocEquipmentItemSwitch),
                assocEquipmentItemSwitchSpecified = equipment.AssocEquipmentItemSwitch != null,
                compCode = equipment.CompCode,
                conAstSegEn = !string.IsNullOrWhiteSpace(equipment.ConAstSegEn) ? Convert.ToDecimal(equipment.ConAstSegEn) : default(decimal),
                conAstSegEnSpecified = equipment.ConAstSegEn != null,
                conAstSegSt = !string.IsNullOrWhiteSpace(equipment.ConAstSegSt) ? Convert.ToDecimal(equipment.ConAstSegSt) : default(decimal),
                conAstSegStSpecified = equipment.ConAstSegSt != null,
                conditionRating = !string.IsNullOrWhiteSpace(equipment.ConditionRating) ? Convert.ToDecimal(equipment.ConditionRating) : default(decimal),
                conditionRatingSpecified = equipment.ConditionRating != null,
                conditionStandard = equipment.ConditionStandard,
                copyEquipment = equipment.CopyEquipment,
                copyNameplateValues = Utils.IsTrue(equipment.CopyNameplateValues),
                copyNameplateValuesSpecified = equipment.CopyNameplateValues != null,
                costSegLgth = !string.IsNullOrWhiteSpace(equipment.CostSegLgth) ? Convert.ToDecimal(equipment.CostSegLgth) : default(decimal),
                costSegLgthSpecified = equipment.CostSegLgth != null,
                costingFlag = equipment.CostingFlag,
                ctaxCode = equipment.CtaxCode,
                custodian = equipment.Custodian,
                custodianPosition = equipment.CustodianPosition,
                //customerName = equipment.CustomerName,
                customerNumber = equipment.CustomerNumber,
                districtCode = equipment.DistrictCode,
                drawingNo = equipment.DrawingNo,
                equipmentClass = equipment.EquipmentClass,
                equipmentCriticality = equipment.EquipmentCriticality,
                equipmentGrpId = equipment.EquipmentGrpId,
                equipmentLocation = equipment.EquipmentLocation,
                equipmentNo = equipment.EquipmentNo,
                equipmentNoDescription1 = equipment.EquipmentNoDescription1,
                equipmentNoDescription2 = equipment.EquipmentNoDescription2,
                equipmentRef = equipment.EquipmentRef,
                equipmentStatus = equipment.EquipmentStatus,
                equipmentType = equipment.EquipmentType,
                equipmentTypeDescription = equipment.EquipmentTypeDescription,
                expElement = equipment.ExpElement,
                iaaAssetInd = Utils.IsTrue(equipment.IaaAssetInd),
                iaaAssetIndSpecified = equipment.IaaAssetInd != null,
                inputBy = equipment.InputBy,
                itemNameCode = equipment.ItemNameCode,
                latestConditionDate = equipment.LatestConditionDate,
                //location = equipment.Location,
                mnemonic = equipment.Mnemonic,
                operatingStandard = equipment.OperatingStandard,
                operatorId = equipment.OperatorId,
                operatorPosition = equipment.OperatorPosition,
                originalDoc = equipment.OriginalDoc,
                parentEquipment = equipment.ParentEquipment,
                parentEquipmentRef = equipment.ParentEquipmentRef,
                partNo = equipment.PartNo,
                permitReqdSw = Utils.IsTrue(equipment.PermitReqdSw),
                permitReqdSwSpecified = equipment.PermitReqdSw != null,
                //plantCode0 = equipment.PlantCode0,
                //plantCode1 = equipment.PlantCode1,
                //plantCode2 = equipment.PlantCode2,
                //plantCode3 = equipment.PlantCode3,
                //plantCode4 = equipment.PlantCode4,
                //plantCode5 = equipment.PlantCode5,
                //plantCodes = equipment.PlantCodes,
                //plantNames = equipment.PlantNames,
                plantNo = equipment.PlantNo,
                poNo = equipment.PoNo,
                primaryFunction = equipment.PrimaryFunction,
                prodUnitItem = equipment.ProdUnitItem,
                purchaseDate = equipment.PurchaseDate,
                purchasePrice = !string.IsNullOrWhiteSpace(equipment.PurchasePrice) ? Convert.ToDecimal(equipment.PurchasePrice) : default(decimal),
                purchasePriceSpecified = equipment.PurchasePrice != null,
                rcmAnalysisSw = Utils.IsTrue(equipment.RcmAnalysisSw),
                rcmAnalysisSwSpecified = equipment.RcmAnalysisSw != null,
                replaceValue = !string.IsNullOrWhiteSpace(equipment.ReplaceValue) ? Convert.ToDecimal(equipment.ReplaceValue) : default(decimal),
                replaceValueSpecified = equipment.ReplaceValue != null,
                segmentUom = equipment.SegmentUom,
                serialNumber = equipment.SerialNumber,
                shutdownEquipment = equipment.ShutdownEquipment,
                stockCode = equipment.StockCode,
                taxCode = equipment.TaxCode,
                traceableFlg = Utils.IsTrue(equipment.TraceableFlg),
                traceableFlgSpecified = equipment.TraceableFlg != null,
                valuationDate = equipment.ValuationDate,
                warrStatType = equipment.WarrStatType,
                warrStatVal = !string.IsNullOrWhiteSpace(equipment.WarrStatVal) ? Convert.ToDecimal(equipment.WarrStatVal) : default(decimal),
                warrStatValSpecified = equipment.WarrStatVal != null,
                warrantyDate = equipment.WarrantyDate,
                equipmentClassif = equipment.ClassCodes.EquipmentClassif,
                equipmentClassif0 = equipment.ClassCodes.EquipmentClassif0,
                equipmentClassif1 = equipment.ClassCodes.EquipmentClassif1,
                equipmentClassif2 = equipment.ClassCodes.EquipmentClassif2,
                equipmentClassif3 = equipment.ClassCodes.EquipmentClassif3,
                equipmentClassif4 = equipment.ClassCodes.EquipmentClassif4,
                equipmentClassif5 = equipment.ClassCodes.EquipmentClassif5,
                equipmentClassif6 = equipment.ClassCodes.EquipmentClassif6,
                equipmentClassif7 = equipment.ClassCodes.EquipmentClassif7,
                equipmentClassif8 = equipment.ClassCodes.EquipmentClassif8,
                equipmentClassif9 = equipment.ClassCodes.EquipmentClassif9,
                equipmentClassif10 = equipment.ClassCodes.EquipmentClassif10,
                equipmentClassif11 = equipment.ClassCodes.EquipmentClassif11,
                equipmentClassif12 = equipment.ClassCodes.EquipmentClassif12,
                equipmentClassif13 = equipment.ClassCodes.EquipmentClassif13,
                equipmentClassif14 = equipment.ClassCodes.EquipmentClassif14,
                equipmentClassif15 = equipment.ClassCodes.EquipmentClassif15,
                equipmentClassif16 = equipment.ClassCodes.EquipmentClassif16,
                equipmentClassif17 = equipment.ClassCodes.EquipmentClassif17,
                equipmentClassif18 = equipment.ClassCodes.EquipmentClassif18,
                equipmentClassif19 = equipment.ClassCodes.EquipmentClassif19,
            };
            proxyEquip.Url = urlService + "/Equipment";
            return proxyEquip.create(operationContext, request);
        }

        /// <summary>
        /// Actualiza el estado del equipo a un estado especificado
        /// </summary>
        /// <param name="operationContext"></param>
        /// <param name="urlService"></param>
        /// <param name="equipment"></param>
        /// <returns></returns>
        public static void UpdateEquipmentData(EquipmentService.OperationContext operationContext, string urlService, Equipment equipment)
        {
            var proxyEquip = new EquipmentService.EquipmentService();

            //El servicio-método se modifica para hacer públicas las variables specifieds requeridas. En caso de actualización del servicio tener presente esta observación

            var request = new EquipmentServiceModifyRequestDTO
            {
                accountCode = equipment.AccountCode,
                activeFlag = Utils.IsTrue(equipment.ActiveFlag),
                activeFlagSpecified = equipment.ActiveFlag != null,
                assocEquipmentItemSwitch = Utils.IsTrue(equipment.AssocEquipmentItemSwitch),
                assocEquipmentItemSwitchSpecified = equipment.AssocEquipmentItemSwitch != null,
                compCode = equipment.CompCode,
                conAstSegEn = !string.IsNullOrWhiteSpace(equipment.ConAstSegEn) ? Convert.ToDecimal(equipment.ConAstSegEn) : default(decimal),
                conAstSegEnSpecified = equipment.ConAstSegEn != null,
                conAstSegSt = !string.IsNullOrWhiteSpace(equipment.ConAstSegSt) ? Convert.ToDecimal(equipment.ConAstSegSt) : default(decimal),
                conAstSegStSpecified = equipment.ConAstSegSt != null,
                conditionRating = !string.IsNullOrWhiteSpace(equipment.ConditionRating) ? Convert.ToDecimal(equipment.ConditionRating) : default(decimal),
                conditionRatingSpecified = equipment.ConditionRating != null,
                conditionStandard = equipment.ConditionStandard,
                //copyEquipment = equipment.CopyEquipment,
                //copyNameplateValues = Utils.IsTrue(equipment.CopyNameplateValues),
                //copyNameplateValuesSpecified = equipment.CopyNameplateValues != null,
                costSegLgth = !string.IsNullOrWhiteSpace(equipment.CostSegLgth) ? Convert.ToDecimal(equipment.CostSegLgth) : default(decimal),
                costSegLgthSpecified = equipment.CostSegLgth != null,
                costingFlag = equipment.CostingFlag,
                ctaxCode = equipment.CtaxCode,
                custodian = equipment.Custodian,
                custodianPosition = equipment.CustodianPosition,
                //customerName = equipment.CustomerName,
                customerNumber = equipment.CustomerNumber,
                districtCode = equipment.DistrictCode,
                drawingNo = equipment.DrawingNo,
                equipmentClass = equipment.EquipmentClass,
                equipmentCriticality = equipment.EquipmentCriticality,
                equipmentGrpId = equipment.EquipmentGrpId,
                equipmentLocation = equipment.EquipmentLocation,
                equipmentNo = equipment.EquipmentNo,
                equipmentNoDescription1 = equipment.EquipmentNoDescription1,
                equipmentNoDescription2 = equipment.EquipmentNoDescription2,
                //equipmentRef = equipment.EquipmentRef,
                equipmentStatus = equipment.EquipmentStatus,
                equipmentType = equipment.EquipmentType,
                equipmentTypeDescription = equipment.EquipmentTypeDescription,
                expElement = equipment.ExpElement,
                iaaAssetInd = Utils.IsTrue(equipment.IaaAssetInd),
                iaaAssetIndSpecified = equipment.IaaAssetInd != null,
                inputBy = equipment.InputBy,
                itemNameCode = equipment.ItemNameCode,
                latestConditionDate = equipment.LatestConditionDate,
                //location = equipment.Location,
                mnemonic = equipment.Mnemonic,
                operatingStandard = equipment.OperatingStandard,
                operatorId = equipment.OperatorId,
                operatorPosition = equipment.OperatorPosition,
                originalDoc = equipment.OriginalDoc,
                parentEquipment = equipment.ParentEquipment,
                parentEquipmentRef = equipment.ParentEquipmentRef,
                partNo = equipment.PartNo,
                permitReqdSw = Utils.IsTrue(equipment.PermitReqdSw),
                permitReqdSwSpecified = equipment.PermitReqdSw != null,
                //plantCode0 = equipment.PlantCode0,
                //plantCode1 = equipment.PlantCode1,
                //plantCode2 = equipment.PlantCode2,
                //plantCode3 = equipment.PlantCode3,
                //plantCode4 = equipment.PlantCode4,
                //plantCode5 = equipment.PlantCode5,
                //plantCodes = equipment.PlantCodes,
                //plantNames = equipment.PlantNames,
                plantNo = equipment.PlantNo,
                poNo = equipment.PoNo,
                primaryFunction = equipment.PrimaryFunction,
                prodUnitItem = equipment.ProdUnitItem,
                purchaseDate = equipment.PurchaseDate,
                purchasePrice = !string.IsNullOrWhiteSpace(equipment.PurchasePrice) ? Convert.ToDecimal(equipment.PurchasePrice) : default(decimal),
                purchasePriceSpecified = equipment.PurchasePrice != null,
                rcmAnalysisSw = Utils.IsTrue(equipment.RcmAnalysisSw),
                rcmAnalysisSwSpecified = equipment.RcmAnalysisSw != null,
                replaceValue = !string.IsNullOrWhiteSpace(equipment.ReplaceValue) ? Convert.ToDecimal(equipment.ReplaceValue) : default(decimal),
                replaceValueSpecified = equipment.ReplaceValue != null,
                segmentUom = equipment.SegmentUom,
                serialNumber = equipment.SerialNumber,
                shutdownEquipment = equipment.ShutdownEquipment,
                stockCode = equipment.StockCode,
                taxCode = equipment.TaxCode,
                traceableFlg = Utils.IsTrue(equipment.TraceableFlg),
                traceableFlgSpecified = equipment.TraceableFlg != null,
                valuationDate = equipment.ValuationDate,
                warrStatType = equipment.WarrStatType,
                warrStatVal = !string.IsNullOrWhiteSpace(equipment.WarrStatVal) ? Convert.ToDecimal(equipment.WarrStatVal) : default(decimal),
                warrStatValSpecified = equipment.WarrStatVal != null,
                warrantyDate = equipment.WarrantyDate,
                equipmentClassif = equipment.ClassCodes.EquipmentClassif,
                equipmentClassif0 = equipment.ClassCodes.EquipmentClassif0,
                equipmentClassif1 = equipment.ClassCodes.EquipmentClassif1,
                equipmentClassif2 = equipment.ClassCodes.EquipmentClassif2,
                equipmentClassif3 = equipment.ClassCodes.EquipmentClassif3,
                equipmentClassif4 = equipment.ClassCodes.EquipmentClassif4,
                equipmentClassif5 = equipment.ClassCodes.EquipmentClassif5,
                equipmentClassif6 = equipment.ClassCodes.EquipmentClassif6,
                equipmentClassif7 = equipment.ClassCodes.EquipmentClassif7,
                equipmentClassif8 = equipment.ClassCodes.EquipmentClassif8,
                equipmentClassif9 = equipment.ClassCodes.EquipmentClassif9,
                equipmentClassif10 = equipment.ClassCodes.EquipmentClassif10,
                equipmentClassif11 = equipment.ClassCodes.EquipmentClassif11,
                equipmentClassif12 = equipment.ClassCodes.EquipmentClassif12,
                equipmentClassif13 = equipment.ClassCodes.EquipmentClassif13,
                equipmentClassif14 = equipment.ClassCodes.EquipmentClassif14,
                equipmentClassif15 = equipment.ClassCodes.EquipmentClassif15,
                equipmentClassif16 = equipment.ClassCodes.EquipmentClassif16,
                equipmentClassif17 = equipment.ClassCodes.EquipmentClassif17,
                equipmentClassif18 = equipment.ClassCodes.EquipmentClassif18,
                equipmentClassif19 = equipment.ClassCodes.EquipmentClassif19,
            };
            proxyEquip.Url = urlService + "/Equipment";
            proxyEquip.modify(operationContext, request);
        }

        /// <summary>
        /// Obtiene el listado de códigos de estado de equipo
        /// </summary>
        /// <param name="ellipseFunctions"></param>
        /// <returns></returns>
        public static List<EllipseCodeItem> GetEquipmentStatusCodeList(EllipseFunctions ellipseFunctions)
        {
            return ellipseFunctions.GetItemCodes("ES");
        }

        public static string GetFetchLastInstallation(EllipseFunctions ef, string district, string equipmentNo, string component, string position)
        {
            var sqlQuery = Queries.GetFetchLastInstallationQuery(ef.dbReference, ef.dbLink, district, equipmentNo, component, position);
            var drLastInstallation = ef.GetQueryResult(sqlQuery);

            if (drLastInstallation == null || drLastInstallation.IsClosed || !drLastInstallation.HasRows) return null;
            var installedcomponent = "";
            while (drLastInstallation.Read())
            {
                installedcomponent = installedcomponent + " " + drLastInstallation["COMPONENTE"].ToString().Trim();
            }
            return installedcomponent;
        }

        public static class Queries
        {
            public static string GetEquipReferencesQuery(string dbReference, string dbLink, string districtCode, string equipmentRef)
            {
                string districtParam;
                //establecemos los parámetrode de distrito
                if (string.IsNullOrEmpty(districtCode))
                    districtParam = " IN (" + Utils.GetListInSeparator(DistrictConstants.GetDistrictList(), ",", "'") + ")";
                else
                    districtParam = " = '" + districtCode + "'";

                var query = "" +
                    " SELECT DISTINCT(EQUIP_NO)" +
                    " FROM (" +
                    "   SELECT TRIM(EQ.EQUIP_NO) EQUIP_NO " +
                    "     FROM  " + dbReference + ".MSF600" + dbLink + " EQ " +
                    "     WHERE TRIM(EQ.EQUIP_NO) = '" + equipmentRef + "' AND EQ.DSTRCT_CODE " + districtParam +
                    "   UNION ALL" +
                    "   SELECT TRIM(EQ.EQUIP_NO) EQUIP_NO" +
                    "     FROM " + dbReference + ".MSF600" + dbLink + " EQ" +
                    "     WHERE LPAD(TRIM(EQ.EQUIP_NO), 12, '0') = LPAD('" + equipmentRef + "',12,'0') AND EQ.DSTRCT_CODE " + districtParam +
                    "   UNION ALL" +
                    "   SELECT REQ.EQUIP_NO" +
                    "     FROM " + dbReference + ".MSF600" + dbLink + " REQ JOIN  " + dbReference + ".MSF601" + dbLink + " RAL ON REQ.EQUIP_NO = RAL.ALT_REF_CODE" +
                    "     WHERE REQ.DSTRCT_CODE " + districtParam + " AND TRIM(RAL.ALTERNATE_REF) = '" + equipmentRef + "'" +
                    " )";

                query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
                
                return query;
            }
            public static string GetEgiEquipmentsQuery(string dbReference, string dbLink, string egi)
            {
                var query = "" +
                    "SELECT TRIM(EQ.EQUIP_NO) EQUIP_NO FROM " + dbReference + ".MSF600" + dbLink + " EQ WHERE TRIM(EQ.EQUIP_GRP_ID) = '" + egi + "'";

                query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
                
                return query;
            }
            public static string GetListEquipmentsQuery(string dbReference, string dbLink, string listType, string listId)
            {
                var query = "" +
                    "SELECT LI.MEM_EQUIP_GRP EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + listType + "' AND TRIM(LI.LIST_ID) = '" + listId + "'";

                query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
                
                return query; 
            }
            public static string GetProductiveUnitEquipmentsQuery(string dbReference, string dbLink, string district, string productiveUnit)
            {
                var query = "" +
                    " SELECT DISTINCT(EQ.EQUIP_NO)" +
                    " FROM " + dbReference + ".MSF600" + dbLink + " EQ WHERE EQ.DSTRCT_CODE = '" + district + "'" +
                    "   START WITH EQ.EQUIP_NO       = '" + productiveUnit + "'" +
                    "   CONNECT BY PRIOR EQ.EQUIP_NO = EQ.PARENT_EQUIP";

                query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
                
                return query;
            }

            public static string GetFetchEquipmentDataQuery(string dbReference, string dbLink, string equipmentNo)
            {
                return GetFetchEquipmentDataQuery(dbReference, dbLink, null, SearchFieldCriteriaType.EquipmentNo.Key, equipmentNo, SearchFieldCriteriaType.None.Key, null);
            }

            public static string GetFetchEquipmentDataQuery(string dbReference, string dbLink, string districtCode, int searchCriteriaKey1, string searchCriteriaValue1, int searchCriteriaKey2, string searchCriteriaValue2, string eqStatus = null)
            {
                //establecemos los parámetrode de distrito
                string districtParam;
                if (string.IsNullOrEmpty(districtCode))
                    districtParam = " AND EQ.DSTRCT_CODE IN (" + Utils.GetListInSeparator(DistrictConstants.GetDistrictList(), ",", "'") + ")";
                else
                    districtParam = " AND EQ.DSTRCT_CODE = '" + districtCode + "'";

                //establecemos los parámetros de estado
                string statusRequirement;
                if (string.IsNullOrEmpty(eqStatus))
                    statusRequirement = "";
                else
                    statusRequirement = " AND EQ.EQUIP_STATUS = '" + Utils.GetCodeKey(eqStatus) + "'";

                var queryCriteria1 = "";
                //establecemos los parámetros del criterio 1
                if (searchCriteriaKey1 == SearchFieldCriteriaType.EquipmentNo.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EQ.EQUIP_NO = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                {
                    var equipParamsQuery = GetEquipReferencesQuery(dbReference, dbLink, districtCode, searchCriteriaValue1);
                    queryCriteria1 = " AND EQ.EQUIP_NO IN (" + equipParamsQuery + ")";
                }
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.ProductiveUnit.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EQ.PARENT_EQUIP = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.EquipmentDescription.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EQ.ITEM_NAME_1||EQ.ITEM_NAME_2 LIKE '%" + searchCriteriaValue1 + "%'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.CreationUser.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EQ.INPUT_BY = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.AccountCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EQ.ACCOUNT_CODE = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.Custodian.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EQ.CUSTODIAN = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.CustodianPosition.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EQ.CUSTODIAN_POSN = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                {
                    if (searchCriteriaKey2 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                        queryCriteria1 = "AND EQ.EQUIP_NO IN (SELECT DISTINCT LI.MEM_EQUIP_GRP EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue1 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "')";
                    else if (searchCriteriaKey2 != SearchFieldCriteriaType.ListId.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                        queryCriteria1 = "AND EQ.EQUIP_NO IN (SELECT DISTINCT LI.MEM_EQUIP_GRP EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue1 + "')";
                }
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                {
                    if (searchCriteriaKey2 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                        queryCriteria1 = "AND EQ.EQUIP_NO IN (SELECT DISTINCT LI.MEM_EQUIP_GRP EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue1 + "')";
                    else if (searchCriteriaKey2 != SearchFieldCriteriaType.ListType.Key || string.IsNullOrWhiteSpace(searchCriteriaValue2))
                        queryCriteria1 = "AND EQ.EQUIP_NO IN (SELECT DISTINCT LI.MEM_EQUIP_GRP EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_ID) = '" + searchCriteriaValue1 + "')";
                }
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.Egi.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EQ.EQUIP_GRP_ID = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.EquipmentClass.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EQ.EQUIP_CLASS = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.EquipmentType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EQ.EQPT_TYPE = '" + searchCriteriaValue1 + "'";
                else if (searchCriteriaKey1 == SearchFieldCriteriaType.EquipmentLocation.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                    queryCriteria1 = " AND EQ.LOCATION = '" + searchCriteriaValue1 + "'";

                var queryCriteria2 = "";
                //establecemos los parámetros del criterio 1
                if (searchCriteriaKey2 == SearchFieldCriteriaType.EquipmentReference.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                {
                    var equipParamsQuery = GetEquipReferencesQuery(dbReference, dbLink, districtCode, searchCriteriaValue2);
                    queryCriteria2 = " AND EQ.EQUIP_NO IN (" + equipParamsQuery + ")";
                }
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.ProductiveUnit.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EQ.PARENT_EQUIP = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.EquipmentDescription.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EQ.ITEM_NAME_1||EQ.ITEM_NAME_2 LIKE '%" + searchCriteriaValue2 + "%'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.CreationUser.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EQ.INPUT_BY = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.AccountCode.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EQ.ACCOUNT_CODE = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.Custodian.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EQ.CUSTODIAN = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.CustodianPosition.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EQ.CUSTODIAN_POSN = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                {
                    if (searchCriteriaKey1 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                        queryCriteria2 = "AND EQ.EQUIP_NO IN (SELECT DISTINCT LI.MEM_EQUIP_GRP EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue1 + "')";
                    else if (searchCriteriaKey1 != SearchFieldCriteriaType.ListId.Key || string.IsNullOrWhiteSpace(searchCriteriaValue1))
                        queryCriteria2 = "AND EQ.EQUIP_NO IN (SELECT DISTINCT LI.MEM_EQUIP_GRP EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue2 + "')";
                }
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.ListId.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                {
                    if (searchCriteriaKey1 == SearchFieldCriteriaType.ListType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue1))
                        queryCriteria2 = "AND EQ.EQUIP_NO IN (SELECT DISTINCT LI.MEM_EQUIP_GRP EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_TYP) = '" + searchCriteriaValue1 + "' AND TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "')";
                    else if (searchCriteriaKey1 != SearchFieldCriteriaType.ListType.Key || string.IsNullOrWhiteSpace(searchCriteriaValue1))
                        queryCriteria2 = "AND EQ.EQUIP_NO IN (SELECT DISTINCT LI.MEM_EQUIP_GRP EQUIP_NO FROM " + dbReference + ".MSF607" + dbLink + " LI WHERE TRIM(LI.LIST_ID) = '" + searchCriteriaValue2 + "')";
                }
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.Egi.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EQ.EQUIP_GRP_ID = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.EquipmentClass.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EQ.EQUIP_CLASS = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.EquipmentType.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EQ.EQPT_TYPE = '" + searchCriteriaValue2 + "'";
                else if (searchCriteriaKey2 == SearchFieldCriteriaType.EquipmentLocation.Key && !string.IsNullOrWhiteSpace(searchCriteriaValue2))
                    queryCriteria2 = " AND EQ.LOCATION = '" + searchCriteriaValue2 + "'";
                //


                var query = "" +
                    " SELECT " +
                    "   EQ.EQUIP_NO, EQ.ACCOUNT_CODE, EQ.ACTIVE_FLG, EQ.ASSOC_EQUIP_SW, EQ.COMP_CODE," +
                    "   EQ.CON_AST_SEG_EN, EQ.CON_AST_SEG_ST," +
                    "   EQ.COND_RATING, EQ.COND_STANDARD," +
                    "   EQ.COST_SEG_LGTH, EQ.COSTING_FLG, EQ.CTAX_CODE, EQ.CUSTODIAN, EQ.CUSTODIAN_POSN," +
                    "   EQ.CUST_NO," +
                    "   EQ.DSTRCT_CODE, EQ.DRAWING_NO, EQ.EQUIP_CLASS, EQ.EQUIP_CRITICALITY, EQ.EQUIP_GRP_ID, EQ.EQUIP_LOCATION," +
                    "   EQ.ITEM_NAME_1, EQ.ITEM_NAME_2, EQ.EQUIP_NO EQUIP_REF, EQ.EQUIP_STATUS, EQ.EQPT_TYPE, EQ.EXP_ELEMENT," +
                    "   EQ.IAA_ASSET_IND, EQ.INPUT_BY, EQ.ITEM_NAME_CODE, EQ.LATEST_COND_DATE, EQ.LOCATION, EQ.MNEMONIC, EQ.MSSS_STATUS_IND," +
                    "   EQ.OPERATING_STD, EQ.OPERATOR_ID, EQ.OPERATOR_POSN, EQ.ORIGINAL_DOC," +
                    "   EQ.PARENT_EQUIP, EQ.PARENT_EQUIP PARENT_EQUIP_REF, EQ.PART_NO, EQ.PERMIT_REQD_SW," +
                    "   EQ.PLANT_NO, EQ.PO_NO, EQ.PRIMARY_FUNCTION, EQ.PROD_UNIT_ITEM, EQ.PURCHASE_DATE, EQ.PURCHASE_PRICE, EQ.RCM_ANALYSIS_SW, EQ.REPLACE_VALUE," +
                    "   EQ.SEGMENT_UOM, EQ.SERIAL_NUMBER, EQ.SHUTDOWN_EQUIP, EQ.STOCK_CODE, EQ.TAX_CODE, EQ.TRACEABLE_FLG," +
                    "   EQ.VALUATION_DATE, EQ.WARR_STAT_TYPE, EQ.WARR_STAT_VAL," +
                    "   EQ.EQUIP_CLASSIFX1, EQ.EQUIP_CLASSIFX2, EQ.EQUIP_CLASSIFX3, EQ.EQUIP_CLASSIFX4, EQ.EQUIP_CLASSIFX5," +
                    "   EQ.EQUIP_CLASSIFX6, EQ.EQUIP_CLASSIFX7, EQ.EQUIP_CLASSIFX8, EQ.EQUIP_CLASSIFX9, EQ.EQUIP_CLASSIFX10," +
                    "   EQ.EQUIP_CLASSIFX11, EQ.EQUIP_CLASSIFX12, EQ.EQUIP_CLASSIFX13, EQ.EQUIP_CLASSIFX14, EQ.EQUIP_CLASSIFX15," +
                    "   EQ.EQUIP_CLASSIFX16, EQ.EQUIP_CLASSIFX17, EQ.EQUIP_CLASSIFX18, EQ.EQUIP_CLASSIFX19, EQ.EQUIP_CLASSIFX20" +
                    " FROM " + dbReference + ".MSF600" + dbLink + " EQ" +
                    " WHERE" +
                    " " + districtParam +
                    " " + queryCriteria1 +
                    " " + queryCriteria2 +
                    " " + statusRequirement +
                    "";
                query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
                
                return query;
            }

            public static string GetFetchLastInstallationQuery(string dbReference, string dbLink, string district, string equipmentNo, string component, string position)
            {
                var positionString = position == "" ? "    AND TRIM ( SUBSTR ( CO.INSTALL_POSN, 17, 2 ) ) is null" : "    AND TRIM ( SUBSTR ( CO.INSTALL_POSN, 17, 2 ) ) = '" + position + "'";

                var query = "" +
                    "WITH " +
                    "  MOV AS " +
                    "  ( " +
                    "    SELECT " +
                    "      TRIM ( SUBSTR ( CO.INSTALL_POSN, 1, 12 ) ) EQUIPO_PADRE, " +
                    "      TRIM ( SUBSTR ( CO.INSTALL_POSN, 13, 4 ) ) COMPONENTE, " +
                    "      TRIM ( SUBSTR ( CO.INSTALL_POSN, 17, 2 ) ) POSICION, " +
                    "      CO.INSTALL_POSN, " +
                    "      ( 99999999 - CO.REVSD_ET_DATE ) FECHA, " +
                    "      CO.DATE_SEQ, " +
                    "      LEAD ( ( 99999999 - CO.REVSD_ET_DATE ) ) OVER ( PARTITION BY CO.INSTALL_POSN ORDER BY CO.REVSD_ET_DATE DESC, CO.DATE_SEQ, CO.TRACING_ACTN DESC ) LEAD_FECHA, " +
                    "      LEAD ( CO.TRACING_ACTN ) OVER ( PARTITION BY CO.INSTALL_POSN ORDER BY CO.REVSD_ET_DATE DESC, CO.DATE_SEQ, CO.TRACING_ACTN DESC ) LEAD_ACTN, " +
                    "      MAX ( ( 99999999 - CO.REVSD_ET_DATE ) ) OVER ( PARTITION BY CO.INSTALL_POSN ) MAX_FECHA, " +
                    "      CO.TRACING_ACTN, " +
                    "      TRIM ( CO.FIT_EQUIP_NO ) EQUIPO " +
                    "    FROM " +
                    "      ELLIPSE.MSF650 CO " +
                    "    WHERE " +
                    "      CO.DSTRCT_CODE = '" + district + "' " +
                    "    AND CO.TRACING_ACTN IN ( 'B', 'C' ) " +
                    "    AND CO.REVSD_ET_DATE IS NOT NULL " +
                    "    AND CO.REVSD_ET_DATE <> '00000000' " +
                    "    AND TRIM ( SUBSTR ( CO.INSTALL_POSN, 1, 12 ) ) = '" + equipmentNo + "' " +
                    "    AND TRIM ( SUBSTR ( CO.INSTALL_POSN, 13, 4 ) ) = '" + component + "' " +
                    positionString +
                    "    ORDER BY " +
                    "      CO.REVSD_ET_DATE DESC, " +
                    "      CO.DATE_SEQ, " +
                    "      CO.TRACING_ACTN DESC " +
                    "  ) " +
                    "SELECT " +
                    "  DECODE ( TRACING_ACTN, 'B', DECODE(TRIM(EQUIPO), NULL, 'EQUIPADO', EQUIPO), 'C', '' ) COMPONENTE " +
                    "FROM " +
                    "  MOV " +
                    "WHERE " +
                    "  FECHA = MAX_FECHA";
                query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
                
                return query;
            }
        }

    }

    public static class TracingActions
    {
        public static bool Fitment(EquipTraceService.OperationContext operationContext, string urlService, string instEquipmentRef, string compCode, string modCode, string fitEquipmentRef, string fitDate, string refType, string refNumber)
        {
            var proxyEquip = new EquipTraceService.EquipTraceService();

            var request = new EquipTraceServiceFitRequestDTO()
            {
                installEquipmentRef = instEquipmentRef,
                compCode = compCode,
                modCode = modCode,
                fitEquipmentRef = fitEquipmentRef,
                ETDate = fitDate,
                refType = refType,
                refNum = refNumber
            };
            proxyEquip.Url = urlService + "/EquipTrace";

            var reply = proxyEquip.fit(operationContext, request);
            if (reply.installEquipmentRef.Equals(request.installEquipmentRef)
                && reply.compCode.Equals(request.compCode)
                && reply.modCode.Equals(request.modCode)
                && reply.fitEquipmentRef.Equals(request.fitEquipmentRef))
                return true;

            var information = reply.warningsAndInformation;

            throw new Exception(information.Aggregate("", (current, info) => current + (current + " " + info.fieldName + ". ")));
        }

        public static bool Defitment(EquipTraceService.OperationContext operationContext, string urlService, string instEquipmentRef, string compCode, string modCode, string fitEquipmentRef, string defitDate, string refType, string refNumber)
        {
            var proxyEquip = new EquipTraceService.EquipTraceService();

            var request = new EquipTraceServiceDefitRequestDTO()
            {
                installEquipmentRef = instEquipmentRef,
                compCode = compCode,
                modCode = modCode,
                fitEquipmentRef = fitEquipmentRef,
                ETDate = defitDate,
                refType = refType,
                refNum = refNumber
            };
            proxyEquip.Url = urlService + "/EquipTrace";

            var reply = proxyEquip.defit(operationContext, request);
            if (reply.installEquipmentRef.Equals(request.installEquipmentRef)
                && reply.compCode.Equals(request.compCode)
                && reply.modCode.Equals(request.modCode)
                && reply.fitEquipmentRef.Equals(request.fitEquipmentRef))
                return true;

            var information = reply.warningsAndInformation;

            throw new Exception(information.Aggregate("", (current, info) => current + (current + " " + info.fieldName + ". ")));
        }
    }

    public static class SearchFieldCriteriaType
    {
        public static KeyValuePair<int, string> None = new KeyValuePair<int, string>(0, "None");
        public static KeyValuePair<int, string> EquipmentNo = new KeyValuePair<int, string>(1, "Equipment No");
        public static KeyValuePair<int, string> EquipmentReference = new KeyValuePair<int, string>(2, "Equipment Ref");
        public static KeyValuePair<int, string> ProductiveUnit = new KeyValuePair<int, string>(3, "ProductiveUnit");
        public static KeyValuePair<int, string> EquipmentDescription = new KeyValuePair<int, string>(4, "Description");
        public static KeyValuePair<int, string> CreationUser = new KeyValuePair<int, string>(5, "Creation User");
        public static KeyValuePair<int, string> AccountCode = new KeyValuePair<int, string>(6, "AccountCode");
        public static KeyValuePair<int, string> Custodian = new KeyValuePair<int, string>(7, "Custodian");
        public static KeyValuePair<int, string> CustodianPosition = new KeyValuePair<int, string>(8, "Cust. Post.");
        public static KeyValuePair<int, string> ListType = new KeyValuePair<int, string>(9, "ListType");
        public static KeyValuePair<int, string> ListId = new KeyValuePair<int, string>(10, "ListId");
        public static KeyValuePair<int, string> Egi = new KeyValuePair<int, string>(11, "EGI");
        public static KeyValuePair<int, string> EquipmentClass = new KeyValuePair<int, string>(12, "Equipment Class");
        public static KeyValuePair<int, string> EquipmentType = new KeyValuePair<int, string>(13, "Equipment Type");
        public static KeyValuePair<int, string> EquipmentLocation = new KeyValuePair<int, string>(14, "Location");

        public static List<KeyValuePair<int, string>> GetSearchFieldCriteriaTypes(bool keyOrder = true)
        {
            var list = new List<KeyValuePair<int, string>>
            {
                None, EquipmentNo, EquipmentReference, ProductiveUnit, EquipmentDescription, CreationUser, AccountCode, Custodian,
                CustodianPosition, ListId, ListType, Egi, EquipmentClass, EquipmentType, EquipmentLocation
            };

            return keyOrder ? list.OrderBy(x => x.Key).ToList() : list.OrderBy(x => x.Value).ToList();
        }
    }
}
