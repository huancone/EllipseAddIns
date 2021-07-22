using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SharedClassLibrary.Cority;

namespace TamizajeExcelAddIn
{
    public class CovidQuestionary
    {
        public QuestionaryHeader Header;

        public string QuestionaryCode;

        public string EmployeeId;
        public string DocumentType;
        public string Sex;
        public string FirstName;
        public string SecondName;
        public string FirstLastName;
        public string SecondLastName;

        public QuestionResponse CodigoBarras { get; set; }
        public QuestionResponse Telefono { get; set; }
        public QuestionResponse Ciudad { get; set; }
        public QuestionResponse FechaRespuesta { get; set; }
        public QuestionResponse FechaResultado { get; set; }
        public QuestionResponse EtapaPrueba { get; set; }
        public QuestionResponse EstadoPrueba { get; set; }
        public QuestionResponse Laboratorio { get; set; }
        public QuestionResponse TipoCaso { get; set; }
        public QuestionResponse Conducta { get; set; }
        public QuestionResponse FuenteCaso { get; set; }
        public QuestionResponse Site { get; set; }
        public QuestionResponse CasoIndice { get; set; }
        public QuestionResponse Area { get; set; }
        public QuestionResponse Ubicacion { get; set; }
        public QuestionResponse Evolucion { get; set; }
        public QuestionResponse Severidad { get; set; }

        public string ActionMessage;
        public CovidQuestionary()
        {
            QuestionaryCode = "COVID";
            DocumentType = "CC";
            Header = new QuestionaryHeader();

            CodigoBarras = new QuestionResponse("COVID-14", null);
            Telefono = new QuestionResponse("COVID-11", null);
            Ciudad = new QuestionResponse("COVID-06", null);
            FechaRespuesta = new QuestionResponse("COVID-01", null);
            FechaResultado = new QuestionResponse("COVID-02", null);
            EtapaPrueba = new QuestionResponse("COVID-15", null);
            EstadoPrueba = new QuestionResponse("COVID-03", null);
            Laboratorio = new QuestionResponse("COVID-16", null);
            TipoCaso = new QuestionResponse("COVID-17", null);
            Conducta = new QuestionResponse("COVID-18", null);
            FuenteCaso = new QuestionResponse("COVID-04", null);
            Site = new QuestionResponse("COVID-12", null);
            CasoIndice = new QuestionResponse("COVID-05", null);
            Area = new QuestionResponse("COVID-07", null);
            Ubicacion = new QuestionResponse("COVID-08", null);
            Evolucion = new QuestionResponse("COVID-09", null);
            Severidad = new QuestionResponse("COVID-13", null);
        }

        public List<QuestionResponse> GetResponseList()
        {
            var list = new List<QuestionResponse>();
            list.Add(CodigoBarras);
            list.Add(Telefono);
            list.Add(Ciudad);
            list.Add(FechaRespuesta);
            list.Add(FechaResultado);
            list.Add(EtapaPrueba);
            list.Add(EstadoPrueba);
            list.Add(Laboratorio);
            list.Add(TipoCaso);
            list.Add(Conducta);
            list.Add(FuenteCaso);
            list.Add(Site);
            list.Add(CasoIndice);
            list.Add(Area);
            list.Add(Ubicacion);
            list.Add(Evolucion);
            list.Add(Severidad);

            return list;
        }
        
    }
}
