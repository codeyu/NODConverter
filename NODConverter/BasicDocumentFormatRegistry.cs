using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
namespace NODConverter
{
    public class BasicDocumentFormatRegistry : IDocumentFormatRegistry
    {

        private IList documentFormats = new ArrayList(); //<DocumentFormat>

        public virtual void addDocumentFormat(DocumentFormat documentFormat)
        {
            documentFormats.Add(documentFormat);
        }

        protected internal virtual IList DocumentFormats
        {
            get
            {
                return documentFormats;
            }
        }

        /// <param name="extension"> the file extension </param>
        /// <returns> the DocumentFormat for this extension, or null if the extension is not mapped </returns>
        public virtual DocumentFormat getFormatByFileExtension(string extension)
        {
            if (extension == null)
            {
                return null;
            }
            string lowerExtension = extension.ToLower();
            for (IEnumerator it = documentFormats.GetEnumerator(); it.MoveNext(); )
            {
                DocumentFormat format = (DocumentFormat)it.Current;
                if (format.FileExtension.Equals(lowerExtension))
                {
                    return format;
                }
            }
            return null;
        }

        public virtual DocumentFormat getFormatByMimeType(string mimeType)
        {
            for (IEnumerator it = documentFormats.GetEnumerator(); it.MoveNext(); )
            {
                DocumentFormat format = (DocumentFormat)it.Current;
                if (format.MimeType.Equals(mimeType))
                {
                    return format;
                }
            }
            return null;
        }
    }
}
