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

        public class EquipmentReferenceCodes
        {
            public string EquipmentCapacity;
            public string RefrigerantType;
            public string FuelCostCenter;
            public string ReconstructedComponent;
            public string XerasModel;
        }
    }
}
