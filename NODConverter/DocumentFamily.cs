using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
namespace NODConverter
{
    /// <summary>
    /// Enum-style class declaring the available document families (Text, Spreadsheet, Presentation).
    /// </summary>
    public class DocumentFamily
    {

        public static readonly DocumentFamily TEXT = new DocumentFamily("Text");
        public static readonly DocumentFamily SPREADSHEET = new DocumentFamily("Spreadsheet");
        public static readonly DocumentFamily PRESENTATION = new DocumentFamily("Presentation");
        public static readonly DocumentFamily DRAWING = new DocumentFamily("Drawing");

        private static IDictionary FAMILIES = new Hashtable();
        static DocumentFamily()
        {
            FAMILIES[TEXT.name] = TEXT;
            FAMILIES[SPREADSHEET.name] = SPREADSHEET;
            FAMILIES[PRESENTATION.name] = PRESENTATION;
            FAMILIES[DRAWING.name] = DRAWING;
        }

        private string name;

        private DocumentFamily(string name)
        {
            this.name = name;
        }

        public virtual string Name
        {
            get
            {
                return name;
            }
        }

        public static DocumentFamily getFamily(string name)
        {
            DocumentFamily family = (DocumentFamily)FAMILIES[name];
            if (family == null)
            {
                throw new System.ArgumentException("invalid DocumentFamily: " + name);
            }
            return family;
        }
    }
}
