using System;
using System.Collections.Generic;
using System.Text;

namespace NODConverter
{
    public interface IDocumentFormatRegistry
    {

        DocumentFormat getFormatByFileExtension(string extension);

        DocumentFormat getFormatByMimeType(string extension);

    }
}
