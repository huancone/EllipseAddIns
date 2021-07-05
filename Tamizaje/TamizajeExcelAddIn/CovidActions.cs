using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SharedClassLibrary.Cority;
using SharedClassLibrary.Cority.MGIPService;


namespace TamizajeExcelAddIn
{
    public static class CovidActions
    {
        public static CovidQuestionary CreateCovidQuestionary(MGIPService service, CovidQuestionary questionary)
        {
            var headerResult = DataIntegrationEngine.CreateQuestionaryResponseHeader(service, questionary.Header);

            var list = questionary.GetResponseList();
            var responseResult = "";
            foreach (var response in list)
            {
                if (!string.IsNullOrWhiteSpace(responseResult))
                    responseResult += ";";

                var qr = new QuestionaryResponse();
                qr.Qrh = questionary.Header.Qrh;
                qr.QuestionaryCode = questionary.QuestionaryCode;
                qr.QuestionCode = response.Code;
                qr.Response = response.Value;
                responseResult += DataIntegrationEngine.CreateQuestionaryResponse(service, qr);
            }

            questionary.ActionMessage = headerResult + responseResult;
            return questionary;
        }
    }
}
