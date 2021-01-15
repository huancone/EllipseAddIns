namespace EllipseDownLostExcelAddIn
{
    /// <summary>
    /// Una clase para almacenar los valores bases de un registro DOWN/LOST
    /// </summary>
    public class LostDownObject
    {
        public string EquipNo;
        public string CompCode;
        public string CompModCode;
        public string Date;
        public string StartTime;
        public string FinishTime;
        public string Elapsed;
        public string ShiftCode;
        public string EventCode;
        public string EventDescription;
        public string WoComment;
        public string WoEvent;//pbv
        public string SymptomId;//pbv
        public string AssetTypeId;//pbv
        public string StatusChangeid;//pbv

        /// <summary>
        /// Constructor de la clase
        /// </summary>
        /// <param name="equipNo">string: Número del Equipo</param>
        /// <param name="compCode">string: Código del Componente (N/A para Lost)</param>
        /// <param name="compModCode">string: Código de Modificador de componente (N/A para Lost)</param>
        /// <param name="date">string: Fecha del evento yyyyMMdd</param>
        /// <param name="startTime">string: Hora de inicio del evento hhmm</param>
        /// <param name="finishTime">string: Hora de finalización del evento hhmm</param>
        /// <param name="elapsed">string: Tiempo transcurrido (puede ser nulo)</param>
        /// <param name="shiftCode">string: código del turno</param>
        /// <param name="eventCode">string: código del evento</param>
        /// <param name="eventDescription">string: descripción del código del evento</param>
        /// <param name="woComment">string: WorkOrder para DownTime ó Texto de comentario para Lost</param>
        public LostDownObject(string equipNo, string compCode, string compModCode, string date, string startTime, string finishTime, string elapsed, string shiftCode, string eventCode, string eventDescription, string woComment)
        {
            EquipNo = equipNo;
            CompCode = compCode;
            CompModCode = compModCode;
            Date = date;
            StartTime = startTime;
            FinishTime = finishTime;
            Elapsed = elapsed;
            ShiftCode = shiftCode;
            EventCode = eventCode;
            EventDescription = eventDescription;
            WoComment = woComment;
        }

        /// <summary>
        /// Constructor de la clase con codigos de falla
        /// </summary>
        /// <param name="equipNo">string: Número del Equipo</param>
        /// <param name="compCode">string: Código del Componente (N/A para Lost)</param>
        /// <param name="compModCode">string: Código de Modificador de componente (N/A para Lost)</param>
        /// <param name="date">string: Fecha del evento yyyyMMdd</param>
        /// <param name="startTime">string: Hora de inicio del evento hhmm</param>
        /// <param name="finishTime">string: Hora de finalización del evento hhmm</param>
        /// <param name="elapsed">string: Tiempo transcurrido (puede ser nulo)</param>
        /// <param name="shiftCode">string: código del turno</param>
        /// <param name="eventCode">string: código del evento</param>
        /// <param name="eventDescription">string: descripción del código del evento</param>
        /// <param name="woComment">string: WorkOrder para DownTime ó Texto de comentario para Lost</param>
        /// <param name="woEvent">string: Orden que se genera por el evento</param>
        /// <param name="symptomId">string: codigo de falla, Sintoma</param>
        /// <param name="assetTypeId">string: codigo de falla, Componente</param>
        /// <param name="statusChangeid">string: codigo de falla, Causa</param>
        public LostDownObject(string equipNo, string compCode, string compModCode, string date, string startTime, string finishTime, string elapsed, string shiftCode, string eventCode, string eventDescription, string woComment, string woEvent, string symptomId, string assetTypeId, string statusChangeid)
        {
            EquipNo = equipNo;
            CompCode = compCode;
            CompModCode = compModCode;
            Date = date;
            StartTime = startTime;
            FinishTime = finishTime;
            Elapsed = elapsed;
            ShiftCode = shiftCode;
            EventCode = eventCode;
            EventDescription = eventDescription;
            WoComment = woComment;
            WoEvent = woEvent;
            SymptomId = symptomId;
            AssetTypeId = assetTypeId;
            StatusChangeid = statusChangeid;
        }
    }

}
