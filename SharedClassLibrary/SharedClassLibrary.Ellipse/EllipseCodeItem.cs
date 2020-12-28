using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;

namespace SharedClassLibrary.Ellipse
{
    public class EllipseCodeItem
    {
        public string Code;
        public string Description;
        public string TableType;
        public string AssocRec;
        public string Active;
        /// <summary>
        /// Inicializa el elemento con su código, descripción, tipo de tabla y registro asociado
        /// </summary>
        /// <param name="code"></param>
        /// <param name="description"></param>
        /// <param name="table_type"></param>
        /// <param name="assoc_rec"></param>
        /// <param name="active"></param>
        public EllipseCodeItem(string code, string description, string table_type, string assoc_rec, string active)
        {
            this.Code = code;
            this.Description = description;
            this.TableType = table_type;
            this.AssocRec = assoc_rec;
            this.Active = active;
        }
        /// <summary>
        /// Inicializa el elemento con su código, descripción y tipo de tabla
        /// </summary>
        /// <param name="code"></param>
        /// <param name="description"></param>
        /// <param name="table_type"></param>
        public EllipseCodeItem(string code, string description, string table_type)
        {
            this.Code = code;
            this.Description = description;
            this.TableType = table_type;
        }
        /// <summary>
        /// Inicializa el elemento con su código y descripción
        /// </summary>
        /// <param name="code"></param>
        /// <param name="description"></param>
        public EllipseCodeItem(string code, string description)
        {
            this.Code = code;
            this.Description = description;
        }

        public KeyValuePair<string, string> ToKeyValuePair()
        {
            var codeItem = new KeyValuePair<string, string>(Code, Description);
            return codeItem;
        }
    }
}
