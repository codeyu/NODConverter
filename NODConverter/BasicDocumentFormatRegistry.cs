using System.Collections;

namespace NODConverter
{
    public class BasicDocumentFormatRegistry : IDocumentFormatRegistry
    {

        private readonly IList _documentFormats = new ArrayList(); //<DocumentFormat>

        public virtual void AddDocumentFormat(DocumentFormat documentFormat)
        {
            _documentFormats.Add(documentFormat);
        }

        protected internal virtual IList DocumentFormats
        {
            get
            {
                return _documentFormats;
            }
        }

        /// <param name="extension"> the file extension </param>
        /// <returns> the DocumentFormat for this extension, or null if the extension is not mapped </returns>
        public virtual DocumentFormat GetFormatByFileExtension(string extension)
        {
            if (extension == null)
            {
                return null;
            }
            string lowerExtension = extension.ToLower();
            for (IEnumerator it = _documentFormats.GetEnumerator(); it.MoveNext(); )
            {
                DocumentFormat format = (DocumentFormat)it.Current;
                if (format != null && format.FileExtension.Equals(lowerExtension))
                {
                    return format;
                }
            }
            return null;
        }

        public virtual DocumentFormat GetFormatByMimeType(string mimeType)
        {
            for (IEnumerator it = _documentFormats.GetEnumerator(); it.MoveNext(); )
            {
                DocumentFormat format = (DocumentFormat)it.Current;
                if (format != null && format.MimeType.Equals(mimeType))
                {
                    return format;
                }
            }
            return null;
        }
    }
}
