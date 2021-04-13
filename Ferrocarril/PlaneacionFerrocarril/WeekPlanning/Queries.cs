using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PlaneacionFerrocarril
{
    public partial class WeekPlanning
    {
        public class Queries
        {
            public static string GetTaskResourcesQuery(string workGroup, string startDate, string finishDate, string additional)
            {

                var paramDate = " AND WOT.PLAN_STR_DATE BETWEEN '" + startDate + "' AND '" + finishDate + "'";
                if (additional.Equals("BACKLOG"))
                    paramDate = " AND WOT.PLAN_STR_DATE <= '" + finishDate + "'"; 
                

                var query = "SELECT" +
                            " WO.DSTRCT_CODE                                          ," +
                            " WO.WORK_ORDER STD_WO," +
                            "     WOT.WO_TASK_NO TASK," +
                            "     WOT.WO_TASK_DESC DESCR," +
                            "     COALESCE(WOT.WORK_GROUP, WO.WORK_GROUP) WORK_GROUP      ," +
                            "     WO.EQUIP_NO                                             ," +
                            " EQ.ITEM_NAME_1," +
                            " WO.MAINT_SCH_TASK                                       ," +
                            " WOT.PLAN_STR_DATE NEXT_SCH," +
                            "     WOSUC.TABLE_DESC STATUS," +
                            "     RES.RESOURCE_TYPE  ," +
                            " RES.ACT_RESRCE_HRS ," +
                            " RES.EST_RESRCE_HRS ," +
                            " RES.EST_RESRCE_HRS - RES.ACT_RESRCE_HRS PND_RESRCE_HRS" +
                            "     FROM" +
                            " ELLIPSE.MSF620 WO LEFT JOIN ELLIPSE.MSF600 EQ ON WO.EQUIP_NO = EQ.EQUIP_NO" +
                            " LEFT JOIN ELLIPSE.MSF623 WOT ON WO.DSTRCT_CODE = WOT.DSTRCT_CODE AND WO.WORK_ORDER = WOT.WORK_ORDER" +
                            " LEFT JOIN ELLIPSE.MSF010 WOSUC ON WO.WO_STATUS_U = WOSUC.TABLE_CODE AND WOSUC.TABLE_TYPE = 'WS'" +
                            " LEFT JOIN ELLIPSE.MSF735 RES ON RES.KEY_735_ID = RPAD((WO.DSTRCT_CODE || WO.WORK_ORDER || WOT.WO_TASK_NO), 22, ' ')" +
                            " WHERE" +
                            " WO.DSTRCT_CODE = 'ICOR'" +
                            " AND COALESCE(WOT.WORK_GROUP, WO.WORK_GROUP) = '"+ workGroup+ "'" +
                            " AND WOT.TASK_STATUS_M <> 'C' AND TRIM(WOT.PLAN_STR_DATE) IS NOT NULL" +
                            paramDate;

                return query;
            }
            public static string GetWorkGroupResourcesQuery(string workGroup)
            {
                var query = " WITH" +
                            " PARAMS AS" +
                            " (" +
                            "   SELECT '" + workGroup + "' WORK_GROUP FROM DUAL" +
                            " )," +
                            " WGRES AS(" +
                            "   SELECT " +
                            "     REC.RESOURCE_TYPE,  REC.REQ_RESRC_NO , " +
                            "     (SELECT" +
                            "        SUM((SUBSTR((DEF_STOP_TIME - DEF_STR_TIME),1,2)) + (SUBSTR((DEF_STOP_TIME - DEF_STR_TIME),3,2)/60))" +
                            "      FROM     " +
                            " 	   ELLIPSE.MSF720 WGR JOIN PARAMS ON WGR.WORK_GROUP = PARAMS.WORK_GROUP" +
                            "     )  HORAS_GRUPO," +
                            "     (SELECT 1-((BDOWN_ALLOW_PC + ASSIGN_OTH_PC)/100) FROM ELLIPSE.MSF720 WHERE WORK_GROUP= REC.WORK_GROUP  )  BREAKDOWN " +
                            "      FROM  " +
                            " 	   ELLIPSE.MSF730_RESRC_REQ REC JOIN PARAMS ON REC.WORK_GROUP = PARAMS.WORK_GROUP" +
                            " )" +
                            " SELECT " +
                            " 	RESOURCE_TYPE," +
                            "     ROUND((REQ_RESRC_NO * HORAS_GRUPO)* BREAKDOWN * 7,0) EST_HRS," +
                            "     '0' ACT_HRS," +
                            "     TRIM(ELLIPSE.GET_DESC_010('TT', RESOURCE_TYPE)) as RES_DESC" +
                            " FROM WGRES" +
                            " WHERE " +
                            " 	ROUND((REQ_RESRC_NO * HORAS_GRUPO)* BREAKDOWN * 7,0)<>0" +
                            " ORDER BY RESOURCE_TYPE";

                return query;
            }
        }


    }
}
