using System;
using System.Collections.Generic;
using System.Text;

namespace NODConverter.OpenOffice.Connection
{
    public class OpenOfficeException : Exception
    {

        private const long serialVersionUID = 1L;

        public OpenOfficeException(string message)
            : base(message)
        {
        }

        public OpenOfficeException(string message, Exception cause)
            : base(message, cause)
        {
        }
    }
}
