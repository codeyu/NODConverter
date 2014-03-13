using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
namespace NODConverter
{
    /// <summary>
    /// Represents a document format ("OpenDocument Text" or "PDF").
    /// Also contains its available export filters.
    /// </summary>
    public class DocumentFormat
    {

        private const string FILTER_NAME = "FilterName";

        private string name;
        private DocumentFamily family;
        private string mimeType;
        private string fileExtension;
        private IDictionary exportOptions = new Hashtable(); //<DocumentFamily,Map<String,Object>>
        private IDictionary importOptions = new Hashtable(); //<String,Object>

        public DocumentFormat()
        {
            // empty constructor needed for XStream deserialization
        }

        public DocumentFormat(string name, string mimeType, string extension)
        {
            this.name = name;
            this.mimeType = mimeType;
            this.fileExtension = extension;
        }

        public DocumentFormat(string name, DocumentFamily family, string mimeType, string extension)
        {
            this.name = name;
            this.family = family;
            this.mimeType = mimeType;
            this.fileExtension = extension;
        }

        public virtual string Name
        {
            get
            {
                return name;
            }
        }

        public virtual DocumentFamily Family
        {
            get
            {
                return family;
            }
        }

        public virtual string MimeType
        {
            get
            {
                return mimeType;
            }
        }

        public virtual string FileExtension
        {
            get
            {
                return fileExtension;
            }
        }

        private string getExportFilter(DocumentFamily family)
        {
            return (string)getExportOptions(family)[FILTER_NAME];
        }

        public virtual bool Importable
        {
            get
            {
                return family != null;
            }
        }

        public virtual bool ExportOnly
        {
            get
            {
                return !Importable;
            }
        }

        public virtual bool isExportableTo(DocumentFormat otherFormat)
        {
            return otherFormat.isExportableFrom(this.family);
        }

        public virtual bool isExportableFrom(DocumentFamily family)
        {
            return getExportFilter(family) != null;
        }

        public virtual void setExportFilter(DocumentFamily family, string filter)
        {
            getExportOptions(family)[FILTER_NAME] = filter;
        }

        public virtual void setExportOption(DocumentFamily family, string name, object value)
        {
            IDictionary options = (IDictionary)exportOptions[family]; //<String,Object>
            if (options == null)
            {
                options = new Hashtable();
                exportOptions[family] = options;
            }
            options[name] = value;
        }

        public virtual IDictionary getExportOptions(DocumentFamily family) //<String,Object>
        {
            IDictionary options = (IDictionary)exportOptions[family]; //<String,Object>
            if (options == null)
            {
                options = new Hashtable();
                exportOptions[family] = options;
            }
            return options;
        }

        public virtual void setImportOption(string name, object value)
        {
            importOptions[name] = value;
        }

        public virtual IDictionary ImportOptions
        {
            get
            {
                if (importOptions != null)
                {
                    return importOptions;
                }
                else
                {
                    return (System.Collections.IDictionary)new System.Collections.Hashtable();
                }
            }
        }
    }
}
