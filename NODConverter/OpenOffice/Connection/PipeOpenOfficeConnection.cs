using System;
using System.Collections.Generic;
using System.Text;

namespace NODConverter.OpenOffice.Connection
{
    /// <summary>
    /// OpenOffice connection using a named pipe
    /// <p>
    /// <b>Warning!</b> This requires the <i>sal3</i> native library shipped with OpenOffice.org;
    /// it must be made available via the <i>java.library.path</i> parameter, e.g.
    /// <pre>
    ///   java -Djava.library.path=/opt/openoffice.org/program my.App
    /// </pre>  
    /// </summary>
    public class PipeOpenOfficeConnection : AbstractOpenOfficeConnection
    {

        public const string DEFAULT_PIPE_NAME = "nodconverter";

        public PipeOpenOfficeConnection()
            : this(DEFAULT_PIPE_NAME)
        {
        }

        public PipeOpenOfficeConnection(string pipeName)
            : base("pipe,name=" + pipeName)
        {
        }
    }
}
