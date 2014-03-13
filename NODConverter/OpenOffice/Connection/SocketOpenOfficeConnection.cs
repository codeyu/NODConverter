using System;
using System.Collections.Generic;
using System.Text;

namespace NODConverter.OpenOffice.Connection
{
    public class SocketOpenOfficeConnection : AbstractOpenOfficeConnection
    {

        public const string DEFAULT_HOST = "127.0.0.1";
        public const int DEFAULT_PORT = 8100;

        public SocketOpenOfficeConnection()
            : this(DEFAULT_HOST, DEFAULT_PORT)
        {
        }

        public SocketOpenOfficeConnection(int port)
            : this(DEFAULT_HOST, port)
        {
        }

        public SocketOpenOfficeConnection(string host, int port)
            : base("socket,host=" + host + ",port=" + port + ",tcpNoDelay=1")//用 pipe代替 socket试试。
        {
        }
    }
}
