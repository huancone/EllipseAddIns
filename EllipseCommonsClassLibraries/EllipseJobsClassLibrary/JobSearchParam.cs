using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Utilities;

namespace EllipseJobsClassLibrary
{
    public class JobSearchParam
    {
        public string AccountCode;
        public string AllDistricts;
        public string AssignPerson;
        public string AttachedToOutage;  
        public string BacklogToleranceDays;
        public string CompCode;
        public string CompModCode;
        public string ConAstSegFr;
        public string ConAstSegTo;
        public string[] Crews;

        public string Data1732;
        public string DateIncludes;
        public string DateIncrement;
        public string DateIncrementUnit;
        public string DatePreset;
        public string DisplayPrevious;
        public string DisplaySuppressed;
        public string DstrctAcctCode;
        public string DstrctAcctCodeSegment;
        public string DstrctCode;

        public string EnableSuppressedWithResourceBalancing;
        public string EquipClass;
        public string EquipGrpId;
        public string EquipLocation;
        public string EquipStatus;
        public string[] EquipmentHierarchy;
        public string ExcludeMaintenanceType;
        //public WorkOrderKeyDTO[] ExcludeWorkOrderKeys;
        public string ExcludeWorkOrderType;
        public string ExportComment;
        public string ExportView;
        public string GraphicalView;
        public string HierarchySearch;
        public string IncludeEquipmentHierarchy;
        public string IncludeOnlyProjectWorkOrders;
        public string IncludePreferedEgi;
        public string IncludeProjectHierarchy;
        public string IncludeSubLists;
        //public WorkOrderKeyDTO[] IncludeWorkOrderKeys;

        public string ListId;
        public string ListTyp;
        public string Location;
        public string LocationFr;

        public string[] MaintTypes;
        public string MatchOnChildren;
        public string OrigPriority;
        public string OriginatorId;
        public string OutageId;
        public string OverlappingDateSearch;

        public string ParentWo;
        public string PlanFinDate;
        public string PlanFinTime;
        public string PlanPriority;
        public string PlanStrDate;
        public string PlanStrTime;
        public string[] PreferedEgi;
        public string PrinterId;
        public string ProdUnitItem;
        //public ProjectRefDTO[] ProjectHierarchy;
        public string ProjectNo;

        public string RecallTimeHrs;
        public string ReportId;
        public string RequestId;
        public string ResourceCrewTotalsOnly;
        public string[] ResourceCrews;
        public string ResourceDisableAvailabilityCache;
        public string ResourceTotalsOnly;
        public string[] ResourceTypes;
        public string ResourceWorkGroupTotalsOnly;
        public string[] ResourceWorkGroups;
        public string RetrieveResourceRequirements;

        public string SchSegFr;
        public string SchSegTo;
        public string SearchEntity;
        public string SearchView;
        public string SeriesId;
        public string ShowResources;
        public string ShutdownNo;
        public string[] ShutdownTypes;
        public string StdJobNo;
        //public EquipListRefDTO[] SubLists;

        public string UseIncreasedForecastLimit;
        
        public string WoJobCode1;
        public string WoJobCode10;
        public string WoJobCode2;
        public string WoJobCode3;
        public string WoJobCode4;
        public string WoJobCode5;
        public string WoJobCode6;
        public string WoJobCode7;
        public string WoJobCode8;
        public string WoJobCode9;
        public string WoStatusMSearch;
        public string[] WoStatusUs;
        public string[] WoTypes;
        public string[] WorkGroups;
        public string WorkOrder;
        public string[] WorkReqClassif;
        public string[] WorkReqType;

        public JobSearchParam()
        {
            AllDistricts = "N";
        }
        public JobsMWPService.JobsMWPSearchParam ToDto()
        {
            var searchParam = new JobsMWPService.JobsMWPSearchParam();

            searchParam.accountCode = AccountCode;
            searchParam.allDistricts = MyUtilities.IsTrue(AllDistricts);
            searchParam.allDistrictsSpecified = AllDistricts != null;
            searchParam.assignPerson = AssignPerson;
            searchParam.attachedToOutage = MyUtilities.IsTrue(AttachedToOutage);
            searchParam.attachedToOutageSpecified = AttachedToOutage != null;
            
            searchParam.backlogToleranceDays = MyUtilities.ToDecimal(BacklogToleranceDays, MyUtilities.ConversionConstants.DEFAULT_NULL_AND_EMPTY);
            searchParam.backlogToleranceDaysSpecified = BacklogToleranceDays != null;
            
            searchParam.compCode = CompCode;
            searchParam.compModCode = CompModCode;
            searchParam.conAstSegFr = MyUtilities.ToDecimal(ConAstSegFr, MyUtilities.ConversionConstants.DEFAULT_NULL_AND_EMPTY);
            searchParam.conAstSegFrSpecified = ConAstSegFr != null;
            searchParam.conAstSegTo = MyUtilities.ToDecimal(ConAstSegTo, MyUtilities.ConversionConstants.DEFAULT_NULL_AND_EMPTY);
            searchParam.conAstSegToSpecified = ConAstSegTo != null;
            searchParam.crews = Crews;

            searchParam.data1732 = Data1732;
            searchParam.dateIncludes = DateIncludes;
            searchParam.dateIncrement = DateIncrement;
            searchParam.dateIncrementUnit = DateIncrementUnit;
            searchParam.datePreset = DatePreset;
            searchParam.displayPrevious = MyUtilities.IsTrue(DisplayPrevious);
            searchParam.displayPreviousSpecified = DisplayPrevious != null;
            searchParam.displaySuppressed = MyUtilities.IsTrue(DisplaySuppressed);
            searchParam.displaySuppressedSpecified = DisplaySuppressed != null;
            searchParam.dstrctAcctCode = DstrctAcctCode;
            searchParam.dstrctAcctCodeSegment = DstrctAcctCodeSegment;
            searchParam.dstrctCode = DstrctCode;
            
            searchParam.enableSuppressedWithResourceBalancing = MyUtilities.IsTrue(EnableSuppressedWithResourceBalancing);
            searchParam.enableSuppressedWithResourceBalancingSpecified = EnableSuppressedWithResourceBalancing != null;
            searchParam.equipClass = EquipClass;
            searchParam.equipGrpId = EquipGrpId;
            searchParam.equipLocation = EquipLocation;
            searchParam.equipStatus = EquipStatus;
            searchParam.equipmentHierarchy =  EquipmentHierarchy;
            searchParam.excludeMaintenanceType = MyUtilities.IsTrue(ExcludeMaintenanceType);
            searchParam.excludeMaintenanceTypeSpecified = ExcludeMaintenanceType != null;

            //searchParam.excludeWorkOrderKeys = ExcludeWorkOrderKeys;

            searchParam.excludeWorkOrderType = MyUtilities.IsTrue(ExcludeWorkOrderType);
            searchParam.excludeWorkOrderTypeSpecified = ExcludeWorkOrderType != null;
            searchParam.exportComment = ExportComment;
            searchParam.exportView = MyUtilities.IsTrue(ExportView);
            searchParam.exportViewSpecified = ExportView != null;

            searchParam.graphicalView = MyUtilities.IsTrue(GraphicalView);
            searchParam.graphicalViewSpecified = GraphicalView != null;
            searchParam.hierarchySearch = HierarchySearch;
            
            searchParam.includeEquipmentHierarchy = MyUtilities.IsTrue(IncludeEquipmentHierarchy);
            searchParam.includeEquipmentHierarchySpecified = IncludeEquipmentHierarchy != null;
            searchParam.includeOnlyProjectWorkOrders = MyUtilities.IsTrue(IncludeOnlyProjectWorkOrders);
            searchParam.includeOnlyProjectWorkOrdersSpecified = IncludeOnlyProjectWorkOrders != null;
            searchParam.includePreferedEGI = MyUtilities.IsTrue(IncludePreferedEgi);
            searchParam.includePreferedEGISpecified = IncludePreferedEgi != null;
            searchParam.includeProjectHierarchy = MyUtilities.IsTrue(IncludeProjectHierarchy);
            searchParam.includeProjectHierarchySpecified = IncludeProjectHierarchy != null;
            searchParam.includeSubLists = MyUtilities.IsTrue(IncludeSubLists);
            searchParam.includeSubListsSpecified = IncludeSubLists != null;

            //searchParam.includeWorkOrderKeys = IncludeWorkOrderKeys;
            
            searchParam.listId = ListId;
            searchParam.listTyp = ListTyp;
            searchParam.location = Location;
            searchParam.locationFr = LocationFr;

            searchParam.maintTypes = MaintTypes;
            searchParam.matchOnChildren = MyUtilities.IsTrue(MatchOnChildren);
            searchParam.matchOnChildrenSpecified = MatchOnChildren != null;
            
            searchParam.origPriority = OrigPriority;
            searchParam.originatorId = OriginatorId;
            searchParam.outageId = MyUtilities.ToDecimal(OutageId, MyUtilities.ConversionConstants.DEFAULT_NULL_AND_EMPTY);
            searchParam.outageIdSpecified = OutageId != null;
            searchParam.overlappingDateSearch = MyUtilities.IsTrue(OverlappingDateSearch);
            searchParam.overlappingDateSearchSpecified = OverlappingDateSearch != null;
            
            searchParam.parentWo = ParentWo;
            searchParam.planFinDate = MyUtilities.ToDateTime(PlanFinDate);
            searchParam.planFinDateSpecified = PlanFinDate != null;
            searchParam.planFinTime = PlanFinTime;
            searchParam.planPriority = PlanPriority;
            searchParam.planStrDate = MyUtilities.ToDateTime(PlanStrDate);
            searchParam.planStrDateSpecified = PlanStrDate != null;
            searchParam.planStrTime = PlanStrTime;
            searchParam.preferedEGI = PreferedEgi;
            searchParam.printerId = PrinterId;
            searchParam.prodUnitItem = MyUtilities.IsTrue(ProdUnitItem);
            searchParam.prodUnitItemSpecified = ProdUnitItem != null;
            //searchParam.projectHierarchy = ProjectHierarchy;
            searchParam.projectNo = ProjectNo;

            searchParam.recallTimeHrs = MyUtilities.ToDecimal(RecallTimeHrs, MyUtilities.ConversionConstants.DEFAULT_NULL_AND_EMPTY);
            searchParam.recallTimeHrsSpecified = RecallTimeHrs != null;
            searchParam.reportId = ReportId;
            searchParam.requestId = RequestId;
            searchParam.resourceCrewTotalsOnly = MyUtilities.IsTrue(ResourceCrewTotalsOnly);
            searchParam.resourceCrewTotalsOnlySpecified = ResourceCrewTotalsOnly != null;
            searchParam.resourceCrews = ResourceCrews;
            searchParam.resourceDisableAvailabilityCache = MyUtilities.IsTrue(ResourceDisableAvailabilityCache);
            searchParam.resourceDisableAvailabilityCacheSpecified = ResourceDisableAvailabilityCache != null;
            searchParam.resourceTotalsOnly = MyUtilities.IsTrue(ResourceTotalsOnly);
            searchParam.resourceTotalsOnlySpecified = ResourceTotalsOnly != null;
            searchParam.resourceTypes = ResourceTypes;
            searchParam.resourceWorkGroupTotalsOnly = MyUtilities.IsTrue(ResourceWorkGroupTotalsOnly);
            searchParam.resourceWorkGroupTotalsOnlySpecified = ResourceWorkGroupTotalsOnly != null;
            searchParam.resourceWorkGroups = ResourceWorkGroups;
            searchParam.retrieveResourceRequirements = MyUtilities.IsTrue(RetrieveResourceRequirements);
            searchParam.retrieveResourceRequirementsSpecified = RetrieveResourceRequirements != null;
            
            searchParam.schSegFr = MyUtilities.ToDecimal(SchSegFr, MyUtilities.ConversionConstants.DEFAULT_NULL_AND_EMPTY);
            searchParam.schSegFrSpecified = SchSegFr != null;
            searchParam.schSegTo = MyUtilities.ToDecimal(SchSegTo, MyUtilities.ConversionConstants.DEFAULT_NULL_AND_EMPTY);
            searchParam.schSegToSpecified = SchSegTo != null;
            searchParam.searchEntity = SearchEntity;
            searchParam.searchView = MyUtilities.IsTrue(SearchView);
            searchParam.searchViewSpecified = SearchView != null;
            searchParam.seriesID = SeriesId;
            searchParam.showResources = MyUtilities.IsTrue(ShowResources);
            searchParam.showResourcesSpecified = ShowResources != null;
            searchParam.shutdownNo = ShutdownNo;
            searchParam.shutdownTypes = ShutdownTypes;
            searchParam.stdJobNo = StdJobNo;
            //searchParam.subLists = SubLists;

            searchParam.useIncreasedForecastLimit = MyUtilities.IsTrue(UseIncreasedForecastLimit);
            searchParam.useIncreasedForecastLimitSpecified = UseIncreasedForecastLimit != null;
            
            searchParam.woJobCode1 = WoJobCode1;
            searchParam.woJobCode10 = WoJobCode10;
            searchParam.woJobCode2 = WoJobCode2;
            searchParam.woJobCode3 = WoJobCode3;
            searchParam.woJobCode4 = WoJobCode4;
            searchParam.woJobCode5 = WoJobCode5;
            searchParam.woJobCode6 = WoJobCode6;
            searchParam.woJobCode7 = WoJobCode7;
            searchParam.woJobCode8 = WoJobCode8;
            searchParam.woJobCode9 = WoJobCode9;
            searchParam.woStatusMSearch = WoStatusMSearch;
            searchParam.woStatusUs = WoStatusUs;
            searchParam.woTypes = WoTypes;
            searchParam.workGroups = WorkGroups;
            searchParam.workOrder = WorkOrder;
            searchParam.workReqClassif = WorkReqClassif;
            searchParam.workReqType = WorkReqType;
            
            return searchParam;
    }
}
}
