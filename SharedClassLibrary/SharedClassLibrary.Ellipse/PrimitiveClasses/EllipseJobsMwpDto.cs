using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using SharedClassLibrary.Utilities;

namespace SharedClassLibrary.Ellipse.PrimitiveClasses
{
    public class JobsMwpDto
    {
        [XmlElement("accountCode")] public string AccountCode;
        [XmlElement("actDurHrs")] public string ActDurHrs;
        [XmlElement("actEquipCost")] public string ActEquipCost;
        [XmlElement("actLabCost")] public string ActLabCost;
        [XmlElement("actMatCost")] public string ActMatCost;
        [XmlElement("actOtherCost")] public string ActOtherCost;
        [XmlElement("actualFinishDate")] public string ActualFinishDate;
        [XmlElement("actualFinishTime")] public string ActualFinishTime;
        [XmlElement("actualStartDate")] public string ActualStartDate;
        [XmlElement("actualStartTime")] public string ActualStartTime;
        [XmlElement("aptwExistsSw")] public string AptwExistsSw;
        [XmlElement("assignPerson")] public string AssignPerson;
        [XmlElement("assocEquipSw")] public string AssocEquipSw;
        [XmlElement("assumeFirstMSTI")] public string AssumeFirstMSTI;
        //[XmlElement("autoGroupProjection")] public JobLinkDTO autoGroupProjection;
        [XmlElement("calcEquipCost")] public string CalcEquipCost;
        [XmlElement("calcLabCost")] public string CalcLabCost;
        [XmlElement("calcLabHrs")] public string CalcLabHrs;
        [XmlElement("calcMatCost")] public string CalcMatCost;

        [XmlElement("calcOthCost")] public string CalcOthCost;

        //[XmlElement("childLinks")] public JobsJobLinkDTO[] childLinks;
        [XmlElement("closedDt")] public string ClosedDt;
        [XmlElement("closedTime")] public string ClosedTime;
        [XmlElement("compCode")] public string CompCode;
        [XmlElement("compModCode")] public string CompModCode;
        [XmlElement("completedCode")] public string CompletedCode;
        [XmlElement("conAstSegFr")] public string ConAstSegFr;
        [XmlElement("conAstSegTo")] public string ConAstSegTo;
        [XmlElement("countyShire")] public string CountyShire;
        [XmlElement("crew")] public string Crew;
        [XmlElement("crteInsitu")] public string CrteInsitu;

        [XmlElement("data1732")] public string Data1732;
        [XmlElement("dateStatus")] public string DateStatus;
        [XmlElement("dstrctAcctCode")] public string DstrctAcctCode;
        [XmlElement("dstrctCode")] public string DstrctCode;

        [XmlElement("emailAddress")] public string EmailAddress;
        [XmlElement("equipClass")] public string EquipClass;
        [XmlElement("equipClassifx1")] public string EquipClassifx1;
        [XmlElement("equipClassifx10")] public string EquipClassifx10;
        [XmlElement("equipClassifx11")] public string EquipClassifx11;
        [XmlElement("equipClassifx12")] public string EquipClassifx12;
        [XmlElement("equipClassifx13")] public string EquipClassifx13;
        [XmlElement("equipClassifx14")] public string EquipClassifx14;
        [XmlElement("equipClassifx15")] public string EquipClassifx15;
        [XmlElement("equipClassifx16")] public string EquipClassifx16;
        [XmlElement("equipClassifx17")] public string EquipClassifx17;
        [XmlElement("equipClassifx18")] public string EquipClassifx18;
        [XmlElement("equipClassifx19")] public string EquipClassifx19;
        [XmlElement("equipClassifx2")] public string EquipClassifx2;
        [XmlElement("equipClassifx20")] public string EquipClassifx20;
        [XmlElement("equipClassifx3")] public string EquipClassifx3;
        [XmlElement("equipClassifx4")] public string EquipClassifx4;
        [XmlElement("equipClassifx5")] public string EquipClassifx5;
        [XmlElement("equipClassifx6")] public string EquipClassifx6;
        [XmlElement("equipClassifx7")] public string EquipClassifx7;
        [XmlElement("equipClassifx8")] public string EquipClassifx8;
        [XmlElement("equipClassifx9")] public string EquipClassifx9;
        [XmlElement("equipGrpId")] public string EquipGrpId;
        [XmlElement("equipLocation")] public string EquipLocation;

        [XmlElement("equipNo")] public string EquipNo;
        [XmlElement("equipStatus")] public string EquipStatus;
        [XmlElement("equipUpdateFlag")] public string EquipUpdateFlag;
        [XmlElement("estDurHrs")] public string EstDurHrs;
        [XmlElement("estEquipCost")] public string EstEquipCost;
        [XmlElement("estLabCost")] public string EstLabCost;
        [XmlElement("estLabHrs")] public string EstLabHrs;
        [XmlElement("estMatCost")] public string EstMatCost;
        [XmlElement("estOtherCost")] public string EstOtherCost;

        [XmlElement("failurePart")] public string FailurePart;
        [XmlElement("faxNumber")] public string FaxNumber;
        [XmlElement("fromLink")] public string FromLink;

        [XmlElement("ganttFinishDateTime")] public string GanttFinishDateTime;//E9 deleted
        [XmlElement("ganttLinkId")] public string GanttLinkId;
        [XmlElement("ganttParentLinkId")] public string GanttParentLinkId;
        [XmlElement("ganttStartDateTime")] public string GanttStartDateTime;//E9 deleted

        [XmlElement("hasAssignTodo")] public string HasAssignTodo;//E9 deleted
        [XmlElement("id")] public string Id;
        [XmlElement("isControl")] public string IsControl;
        [XmlElement("isMSTParent")] public string IsMSTParent;
        [XmlElement("isParent")] public string IsParent;
        [XmlElement("itemName1")] public string ItemName1;
        [XmlElement("itemName2")] public string ItemName2;
        [XmlElement("jobId")] public string JobId;
        [XmlElement("jobParentId")] public string JobParentId;
        [XmlElement("jobType")] public string JobType;
        [XmlElement("lastPerformedDate")] public string LastPerformedDate;
        [XmlElement("lastPerformedDstrctCode")] public string LastPerformedDstrctCode;
        [XmlElement("lastPerformedWorkOrder")] public string LastPerformedWorkOrder;
        [XmlElement("linkDistrictCode")] public string LinkDistrictCode;
        [XmlElement("linkEntityId")] public string LinkEntityId;
        [XmlElement("linkEntityType")] public string LinkEntityType;
        [XmlElement("linkErpRef")] public string LinkErpRef;
        [XmlElement("linkErpTaskRef")] public string LinkErpTaskRef;
        [XmlElement("linkId")] public string LinkId;
        [XmlElement("linkJobLinkListId")] public string LinkJobLinkListId;
        [XmlElement("linkLagOrLead")] public string LinkLagOrLead;
        [XmlElement("linkOffset")] public string LinkOffset;
        [XmlElement("linkScale")] public string LinkScale;
        [XmlElement("linkScaledOffset")] public string LinkScaledOffset;
        [XmlElement("linkSequenceNum")] public string LinkSequenceNum;
        [XmlElement("linkStatType")] public string LinkStatType;
        [XmlElement("linkStatValue")] public string LinkStatValue;
        [XmlElement("linkState")] public string LinkState;
        [XmlElement("linkType")] public string LinkType;

        //[XmlElement("linkedChildren")] public JobsMWPDTO[] linkedChildren;
        [XmlElement("linkedInd")] public string LinkedInd;
        [XmlElement("location")] public string Location;
        [XmlElement("locationFr")] public string LocationFr;

        [XmlElement("maintSchTask")] public string MaintSchTask;
        [XmlElement("maintType")] public string MaintType;
        [XmlElement("matUpdateFlag")] public string MatUpdateFlag;
        [XmlElement("mnemonic")] public string Mnemonic;
        [XmlElement("mstReference")] public string MstReference;
        [XmlElement("multipleResourceRequirements")] public string MultipleResourceRequirements;

        [XmlElement("noGanttMove")] public string NoGanttMove;
        [XmlElement("noGanttResize")] public string NoGanttResize;
        [XmlElement("noOfTasks")] public string NoOfTasks;
        [XmlElement("numberOfMSTisRemaining")] public string NumberOfMSTisRemaining;

        [XmlElement("origPriority")] public string OrigPriority;
        [XmlElement("originalPlannedFinishDate")] public string OriginalPlannedFinishDate;
        [XmlElement("originalPlannedFinishTime")] public string OriginalPlannedFinishTime;
        [XmlElement("originalPlannedStartDate")] public string OriginalPlannedStartDate;
        [XmlElement("originalPlannedStartTime")] public string OriginalPlannedStartTime;
        [XmlElement("originatorId")] public string OriginatorId;
        [XmlElement("otherUpdateFlag")] public string OtherUpdateFlag;
        [XmlElement("parentEntityId")] public string ParentEntityId;
        [XmlElement("parentEquip")] public string ParentEquip;
        [XmlElement("parentJobType")] public string ParentJobType;
        [XmlElement("parentLinkId")] public string ParentLinkId;
        [XmlElement("parentPlantNo")] public string ParentPlantNo;
        [XmlElement("parentWo")] public string ParentWo;
        [XmlElement("partNo")] public string PartNo;
        [XmlElement("partialCacheKey")] public string PartialCacheKey;
        [XmlElement("pcComplete")] public string PcComplete;
        [XmlElement("planFinDate")] public string PlanFinDate;
        [XmlElement("planFinTime")] public string PlanFinTime;
        [XmlElement("planPriority")] public string PlanPriority;
        [XmlElement("planStatType")] public string PlanStatType;
        [XmlElement("planStatVal")] public string PlanStatVal;
        [XmlElement("planStrDate")] public string PlanStrDate;
        [XmlElement("planStrTime")] public string PlanStrTime;
        [XmlElement("plantNo")] public string PlantNo;
        [XmlElement("prefDate")] public string PrefDate;
        [XmlElement("prefTime")] public string PrefTime;
        [XmlElement("printerName")] public string PrinterName;
        [XmlElement("prodUnitItem")] public string ProdUnitItem;
        [XmlElement("projDesc")] public string ProjDesc;
        [XmlElement("projectNo")] public string ProjectNo;

        [XmlElement("raisedDate")] public string RaisedDate;
        [XmlElement("raisedTime")] public string RaisedTime;
        [XmlElement("recallTimeHrs")] public string RecallTimeHrs;
        [XmlElement("reference")] public string Reference;
        [XmlElement("reqByDate")] public string ReqByDate;
        [XmlElement("reqByTime")] public string ReqByTime;
        [XmlElement("reqStartDate")] public string ReqStartDate;
        [XmlElement("reqStartTime")] public string ReqStartTime;
        [XmlElement("requestId")] public string RequestId;
        [XmlElement("resUpdateFlag")] public string ResUpdateFlag;

        //[XmlElement("resourceRequirements")] public JobsResourceDTO[] resourceRequirements;

        [XmlElement("restartChildWODstrctCode")] public string RestartChildWODstrctCode;
        [XmlElement("restartChildWOPlanStrDate")] public string RestartChildWOPlanStrDate;
        [XmlElement("restartChildWOWorkOrder")] public string RestartChildWOWorkOrder;
        [XmlElement("restartMSTIFromLinkMstReference")] public string RestartMSTIFromLinkMstReference;
        [XmlElement("restartMSTIFromLinkPlanStrDate")] public string RestartMSTIFromLinkPlanStrDate;
        [XmlElement("restartMSTIMstReference")] public string RestartMSTIMstReference;
        [XmlElement("restartMSTIPlanStrDate")] public string RestartMSTIPlanStrDate;
        [XmlElement("restartParentWODstrctCode")] public string RestartParentWODstrctCode;
        [XmlElement("restartParentWOPlanStrDate")] public string RestartParentWOPlanStrDate;
        [XmlElement("restartParentWOWorkOrder")] public string RestartParentWOWorkOrder;

        //[XmlElement("restartPrevLinks")] public TaskKeyDTO[] restartPrevLinks;

        [XmlElement("schSegFr")] public string SchSegFr;
        [XmlElement("schSegTo")] public string SchSegTo;
        [XmlElement("schedDesc1")] public string SchedDesc1;
        [XmlElement("schedDesc2")] public string SchedDesc2;
        [XmlElement("sequenceNo")] public string SequenceNo;
        [XmlElement("serialNumber")] public string SerialNumber;
        [XmlElement("shortDesc1")] public string ShortDesc1;
        [XmlElement("shortDesc2")] public string ShortDesc2;
        [XmlElement("shutdownNo")] public string ShutdownNo;
        [XmlElement("shutdownType")] public string ShutdownType;
        [XmlElement("source")] public string Source;
        [XmlElement("state")] public string State;
        [XmlElement("statutoryFlg")] public string StatutoryFlg;
        [XmlElement("stdJobNo")] public string StdJobNo;
        [XmlElement("streetName")] public string StreetName;
        [XmlElement("streetNo")] public string StreetNo;
        [XmlElement("suburb")] public string Suburb;
        [XmlElement("suppressingMST")] public string SuppressingMST;

        [XmlElement("targtFinDate")] public string TargtFinDate;
        [XmlElement("targtStrDate")] public string TargtStrDate;
        [XmlElement("taskAptwSw")] public string TaskAptwSw;
        [XmlElement("townCity")] public string TownCity;

        [XmlElement("unitOfWork")] public string UnitOfWork;
        [XmlElement("unitsComplete")] public string UnitsComplete;
        [XmlElement("unitsRequired")] public string UnitsRequired;

        [XmlElement("woDesc")] public string WoDesc;
        [XmlElement("woJobCodex1")] public string WoJobCodex1;
        [XmlElement("woJobCodex10")] public string WoJobCodex10;
        [XmlElement("woJobCodex2")] public string WoJobCodex2;
        [XmlElement("woJobCodex3")] public string WoJobCodex3;
        [XmlElement("woJobCodex4")] public string WoJobCodex4;
        [XmlElement("woJobCodex5")] public string WoJobCodex5;
        [XmlElement("woJobCodex6")] public string WoJobCodex6;
        [XmlElement("woJobCodex7")] public string WoJobCodex7;
        [XmlElement("woJobCodex8")] public string WoJobCodex8;
        [XmlElement("woJobCodex9")] public string WoJobCodex9;
        [XmlElement("woStatusM")] public string WoStatusM;
        [XmlElement("woStatusU")] public string WoStatusU;
        [XmlElement("woTaskNo")] public string WoTaskNo;
        [XmlElement("woType")] public string WoType;
        [XmlElement("workGroup")] public string WorkGroup;
        [XmlElement("workOrder")] public string WorkOrder;
        [XmlElement("workRequestDescription")] public string WorkRequestDescription;

        [XmlElement("zipCode")] public string ZipCode;


        public JobsMwpDto()
        {

        }
    }
}
