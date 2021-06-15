using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using SharedClassLibrary.Cority;
using SharedClassLibrary.Utilities;

namespace SharedClassLibrary.Cority
{
    public static class Questionaries
    {
        public static bool? UpdateExistingRecords;
        public static bool? AutoInsertBaseTableValues;
        public static bool? InsertMultiple;
        public static bool? AlwaysInsert;
        public static bool? ContainsHeaderRow;
        public static bool? XmlFile;
        public static string DateFormat = "yyyy-mm-dd";

        private static void SetHttpSecurityProtocol()
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                                                   | SecurityProtocolType.Tls11
                                                   | SecurityProtocolType.Tls12
                                                   | SecurityProtocolType.Ssl3;
        }
        public static string CreateQuestionaryResponseHeader(MGIPService.MGIPService service, QuestionaryHeader header)
        {
            SetHttpSecurityProtocol();
            var qString = "" + header.Qrh + ", " + header.EmployeeId + ", " + header.DateOfResponse + ", " + header.QuestionaryCode + ", " + header.ToBeReviewed;

            byte[] data = Encoding.ASCII.GetBytes(qString);
            const string importFormat = "QRHId, Employee.EmployeNumber, DateOfResponse, Questionnarire.Code, ToBeReviewedBy.LoginName";
            const string currentModel = "Questionnaires.MedicalQuestionnaireResponseHeader";

            var updateExistingRecords = UpdateExistingRecords ?? false;
            var updateExistingRecordsSpecified = UpdateExistingRecords != null;
            var autoInsertBaseTableValues = AutoInsertBaseTableValues ?? false;
            var autoInsertBaseTableValuesSpecified = AutoInsertBaseTableValues != null;
            var insertMultiple = InsertMultiple ?? false;
            var insertMultipleSpecified = InsertMultiple != null;

            var alwaysInsert = AlwaysInsert ?? false;
            var alwaysInsertSpecified = AlwaysInsert != null;
            var containsHeaderRow = ContainsHeaderRow ?? false;
            var containsHeaderRowSpecified = ContainsHeaderRow != null;
            var xmlFile = XmlFile ?? false;
            var xmlFileSpecified = XmlFile != null;
            var dateFormat = DateFormat;
            
            object customLogicParam = null;
            var bypassBr = false;
            var byPassBrSpecified = false;
            string customLogicString = null;
            
            //var reply = service.DoImport(data, importFormat, currentModel, Authentication._username, Authentication._password, updateExistingRecords, updateExistingRecordsSpecified,
            //    autoInsertBaseTableValues, autoInsertBaseTableValuesSpecified, insertMultiple, insertMultipleSpecified,
            //    alwaysInsert, alwaysInsertSpecified, containsHeaderRow, containsHeaderRowSpecified,
            //    xmlFile, xmlFileSpecified, dateFormat, customLogicParam, bypassBr, byPassBrSpecified, customLogicString);

            var reply = service.DoImport2(data, importFormat, currentModel, Authentication._username, Authentication._password, updateExistingRecords, updateExistingRecordsSpecified,
                autoInsertBaseTableValues, autoInsertBaseTableValuesSpecified, insertMultiple, insertMultipleSpecified,
                alwaysInsert, alwaysInsertSpecified, containsHeaderRow, containsHeaderRowSpecified,
                xmlFile, xmlFileSpecified, dateFormat, customLogicParam, customLogicString);
                

            return reply;
        }

        public static string CreateQuestionaryResponse(MGIPService.MGIPService service, QuestionaryResponse response)
        {
            SetHttpSecurityProtocol();
            var qString = "" + response.Qrh + ", " + response.QuestionaryCode + ", " + response.QuestionaryCode + ", " + response.Response;

            byte[] data = Encoding.ASCII.GetBytes(qString);
            const string importFormat = "QuestionnaireResponseHeader.QR, Question.Questionnaire.Code, Question.Question.Code, Response";
            const string currentModel = "Questionnaires.QuestionnaireResponse";

            var updateExistingRecords = UpdateExistingRecords ?? false;
            var updateExistingRecordsSpecified = UpdateExistingRecords != null;
            var autoInsertBaseTableValues = AutoInsertBaseTableValues ?? false;
            var autoInsertBaseTableValuesSpecified = AutoInsertBaseTableValues != null;
            var insertMultiple = InsertMultiple ?? false;
            var insertMultipleSpecified = InsertMultiple != null;

            var alwaysInsert = AlwaysInsert ?? false;
            var alwaysInsertSpecified = AlwaysInsert != null;
            var containsHeaderRow = ContainsHeaderRow ?? false;
            var containsHeaderRowSpecified = ContainsHeaderRow != null;
            var xmlFile = XmlFile ?? false;
            var xmlFileSpecified = XmlFile != null;
            var dateFormat = DateFormat;

            object customLogicParam = null;
            var bypassBr = false;
            var byPassBrSpecified = false;
            string customLogicString = null;

            var reply = service.DoImport(data, importFormat, currentModel, Authentication._username, Authentication._password, updateExistingRecords, updateExistingRecordsSpecified,
                autoInsertBaseTableValues, autoInsertBaseTableValuesSpecified, insertMultiple, insertMultipleSpecified,
                alwaysInsert, alwaysInsertSpecified, containsHeaderRow, containsHeaderRowSpecified,
                xmlFile, xmlFileSpecified, dateFormat, customLogicParam, bypassBr, byPassBrSpecified, customLogicString);

            return reply;
        }
    }
}
