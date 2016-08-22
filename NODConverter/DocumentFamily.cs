using System;
using System.Collections;

namespace NODConverter
{
    /// <summary>
    /// Enum-style class declaring the available document families (Text, Spreadsheet, Presentation).
    /// </summary>
    public class DocumentFamily
    {

        public static readonly DocumentFamily Text = new DocumentFamily("Text");
        public static readonly DocumentFamily Spreadsheet = new DocumentFamily("Spreadsheet");
        public static readonly DocumentFamily Presentation = new DocumentFamily("Presentation");
        public static readonly DocumentFamily Drawing = new DocumentFamily("Drawing");

        private static readonly IDictionary Families = new Hashtable();
        static DocumentFamily()
        {
            Families[Text._name] = Text;
            Families[Spreadsheet._name] = Spreadsheet;
            Families[Presentation._name] = Presentation;
            Families[Drawing._name] = Drawing;
        }

        private readonly string _name;

        private DocumentFamily(string name)
        {
            _name = name;
        }

        public virtual string Name
        {
            get
            {
                return _name;
            }
        }

        public static DocumentFamily GetFamily(string name)
        {
            DocumentFamily family = (DocumentFamily)Families[name];
            if (family == null)
            {
                throw new ArgumentException("invalid DocumentFamily: " + name);
            }
            return family;
        }
    }
}
