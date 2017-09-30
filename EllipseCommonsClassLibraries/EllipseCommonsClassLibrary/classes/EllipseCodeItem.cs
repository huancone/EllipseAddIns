using System.Diagnostics.CodeAnalysis;

namespace EllipseCommonsClassLibrary.Classes
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class EllipseCodeItem
    {
        public string code;
        public string description;
        public string table_type;
        public string assoc_rec;
        /// <summary>
        /// Inicializa el elemento con su código, descripción, tipo de tabla y registro asociado
        /// </summary>
        /// <param name="code"></param>
        /// <param name="description"></param>
        /// <param name="table_type"></param>
        /// <param name="assoc_rec"></param>
        public EllipseCodeItem(string code, string description, string table_type, string assoc_rec)
        {
            this.code = code;
            this.description = description;
            this.table_type = table_type;
            this.assoc_rec = assoc_rec;
        }
        /// <summary>
        /// Inicializa el elemento con su código, descripción y tipo de tabla
        /// </summary>
        /// <param name="code"></param>
        /// <param name="description"></param>
        /// <param name="table_type"></param>
        public EllipseCodeItem(string code, string description, string table_type)
        {
            this.code = code;
            this.description = description;
            this.table_type = table_type;
        }
        /// <summary>
        /// Inicializa el elemento con su código y descripción
        /// </summary>
        /// <param name="code"></param>
        /// <param name="description"></param>
        public EllipseCodeItem(string code, string description)
        {
            this.code = code;
            this.description = description;
        }

    }
}
