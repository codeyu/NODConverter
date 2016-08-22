using System;

namespace NODConverter.OpenOffice.Connection
{
    public class OpenOfficeException : Exception
    {

        //private const long SerialVersionUid = 1L;

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
