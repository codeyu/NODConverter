namespace NODConverter.OpenOffice.Connection
{
    public class SocketOpenOfficeConnection : AbstractOpenOfficeConnection
    {

        public const string DefaultHost = "127.0.0.1";
        public const int DefaultPort = 8100;

        public SocketOpenOfficeConnection()
            : this(DefaultHost, DefaultPort)
        {
        }

        public SocketOpenOfficeConnection(int port)
            : this(DefaultHost, port)
        {
        }

        public SocketOpenOfficeConnection(string host, int port)
            : base("socket,host=" + host + ",port=" + port + ",tcpNoDelay=1")//用 pipe代替 socket试试。
        {
        }
    }
}
