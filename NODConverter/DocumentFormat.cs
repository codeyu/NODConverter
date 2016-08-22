using System.Collections;

namespace NODConverter
{
    /// <summary>
    /// Represents a document format ("OpenDocument Text" or "PDF").
    /// Also contains its available export filters.
    /// </summary>
    public class DocumentFormat
    {

        private const string FilterName = "FilterName";

        private readonly string _name;
        private readonly DocumentFamily _family;
        private readonly string _mimeType;
        private readonly string _fileExtension;
        private readonly IDictionary _exportOptions = new Hashtable(); //<DocumentFamily,Map<String,Object>>
        private readonly IDictionary _importOptions = new Hashtable(); //<String,Object>

        public DocumentFormat()
        {
            // empty constructor needed for XStream deserialization
        }

        public DocumentFormat(string name, string mimeType, string extension)
        {
            _name = name;
            _mimeType = mimeType;
            _fileExtension = extension;
        }

        public DocumentFormat(string name, DocumentFamily family, string mimeType, string extension)
        {
            _name = name;
            _family = family;
            _mimeType = mimeType;
            _fileExtension = extension;
        }

        public virtual string Name
        {
            get
            {
                return _name;
            }
        }

        public virtual DocumentFamily Family
        {
            get
            {
                return _family;
            }
        }

        public virtual string MimeType
        {
            get
            {
                return _mimeType;
            }
        }

        public virtual string FileExtension
        {
            get
            {
                return _fileExtension;
            }
        }

        private string GetExportFilter(DocumentFamily family)
        {
            return (string)GetExportOptions(family)[FilterName];
        }

        public virtual bool Importable
        {
            get
            {
                return _family != null;
            }
        }

        public virtual bool ExportOnly
        {
            get
            {
                return !Importable;
            }
        }

        public virtual bool IsExportableTo(DocumentFormat otherFormat)
        {
            return otherFormat.IsExportableFrom(_family);
        }

        public virtual bool IsExportableFrom(DocumentFamily family)
        {
            return GetExportFilter(family) != null;
        }

        public virtual void SetExportFilter(DocumentFamily family, string filter)
        {
            GetExportOptions(family)[FilterName] = filter;
        }

        public virtual void SetExportOption(DocumentFamily family, string name, object value)
        {
            IDictionary options = (IDictionary)_exportOptions[family]; //<String,Object>
            if (options == null)
            {
                options = new Hashtable();
                _exportOptions[family] = options;
            }
            options[name] = value;
        }

        public virtual IDictionary GetExportOptions(DocumentFamily family) //<String,Object>
        {
            IDictionary options = (IDictionary)_exportOptions[family]; //<String,Object>
            if (options == null)
            {
                options = new Hashtable();
                _exportOptions[family] = options;
            }
            return options;
        }

        public virtual void SetImportOption(string name, object value)
        {
            _importOptions[name] = value;
        }

        public virtual IDictionary ImportOptions
        {
            get
            {
                if (_importOptions != null)
                {
                    return _importOptions;
                }
                else
                {
                    return new Hashtable();
                }
            }
        }
    }
}
