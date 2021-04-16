using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PlaneacionFerrocarril.TemperatureWagon
{
    public class LogCar
    {
        public string Order;//
        public string Owner;//
        public string Number;
        public string AxleNumber;
        public string Spacing;
        public string Ch1;
        public string Ch2;
        public string Ch3;
        public string Ch4;
        public string Alarms;

        public LogCar(string stringLine)
        {
            //Array(0, 1),
            //Array(5, 1),
            //Array(11, 1),
            //Array(20, 1),
            //Array(25, 1),
            //Array(35, 1),
            //Array(40, 1),
            //Array(44, 1),
            //Array(48, 1),
            //Array(52, 1),
            //Array(58, 1),
            //Array(64, 1))
            if (string.IsNullOrWhiteSpace(stringLine) || stringLine.Length < 52)
                throw new Exception("No se puede crear el registro de Log. La línea de información de Log no tiene las características requeridas. Por favor verifique el archivo de Log");
            Order = stringLine.Substring(0, 5).Trim();
            Owner = stringLine.Substring(5, 6).Trim();
            Number = stringLine.Substring(11, 9).Trim();
            AxleNumber = stringLine.Substring(20, 5).Trim();
            Spacing = stringLine.Substring(25, 10).Trim();
            Ch1 = stringLine.Substring(35, 5).Trim();
            Ch2 = stringLine.Substring(40, 4).Trim();
            Ch3 = stringLine.Substring(44, 4).Trim();
            Ch4 = stringLine.Substring(48, 4).Trim();
        }
    }
}
