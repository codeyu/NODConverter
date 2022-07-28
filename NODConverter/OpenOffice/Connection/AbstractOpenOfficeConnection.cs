using System;
using System.Net.Sockets;
using System.Runtime.CompilerServices;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.bridge;
using unoidl.com.sun.star.connection;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.ucb;
using unoidl.com.sun.star.uno;
using Exception = System.Exception;
namespace NODConverter.OpenOffice.Connection
{
    public abstract class AbstractOpenOfficeConnection : IOpenOfficeConnection, XEventListener
    {
        private readonly string _connectionString;
        private XComponent _bridgeComponent;
        private XMultiComponentFactory _serviceManager;
        private XComponentContext _componentContext;
        private XBridge _bridge;
        private bool _connected;
        private bool _expectingDisconnection;
        protected internal AbstractOpenOfficeConnection(string connectionString)
        {
            _connectionString = connectionString;
        }
        [MethodImpl(MethodImplOptions.Synchronized)]
        public virtual void Connect()
        {
            Console.WriteLine("connecting");
            try
            {
                EnvUtils.InitUno();
                SocketUtils.Connect();
                //var sock = new Socket(AddressFamily.InterNetwork, SocketType.Raw, ProtocolType.IP);
                XComponentContext localContext = Bootstrap.bootstrap();
                //XComponentContext localContext = Bootstrap.defaultBootstrap_InitialComponentContext();
                XMultiComponentFactory localServiceManager = localContext.getServiceManager();
                XConnector connector = (XConnector)localServiceManager.createInstanceWithContext("com.sun.star.connection.Connector", localContext);
                XConnection connection = connector.connect(_connectionString);
                XBridgeFactory bridgeFactory = (XBridgeFactory)localServiceManager.createInstanceWithContext("com.sun.star.bridge.BridgeFactory", localContext);
                _bridge = bridgeFactory.createBridge("", "urp", connection, null);
                _bridgeComponent = (XComponent)_bridge;
                _bridgeComponent.addEventListener(this);
                _serviceManager = (XMultiComponentFactory)_bridge.getInstance("StarOffice.ServiceManager");
                XPropertySet properties = (XPropertySet)_serviceManager;
                // Get the default context from the office server. 
                var oDefaultContext = properties.getPropertyValue("DefaultContext");

                _componentContext = (XComponentContext)oDefaultContext.Value;
                _connected = true;
                Console.WriteLine("connected");
            }
            catch (NoConnectException connectException)
            {
                throw new OpenOfficeException("connection failed: " + _connectionString + ": " + connectException.Message);
            }
            catch (Exception exception)
            {
                throw new OpenOfficeException("connection failed: " + _connectionString, exception);
            }
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        public virtual void Disconnect()
        {
            SocketUtils.Disconnect();
            Console.WriteLine("disconnecting");
            _expectingDisconnection = true;
            _bridgeComponent.dispose();
        }

        public virtual bool Connected
        {
            get
            {
                return _connected;
            }
        }

        public virtual void disposing(EventObject @event)
        {
            _connected = false;
            if (_expectingDisconnection)
            {
                Console.WriteLine("disconnected");
            }
            else
            {
                Console.WriteLine("disconnected unexpectedly");
            }
            _expectingDisconnection = false;
        }

        // for unit tests only
        internal virtual void SimulateUnexpectedDisconnection()
        {
            disposing(null);
            _bridgeComponent.dispose();
        }

        private object GetService(string className)
        {
            try
            {
                if (!_connected)
                {
                    Console.WriteLine("trying to (re)connect");
                    Connect();
                }
                return _serviceManager.createInstanceWithContext(className, _componentContext);
            }
            catch (Exception exception)
            {
                throw new OpenOfficeException("could not obtain service: " + className, exception);
            }
        }

        public virtual XComponentLoader Desktop
        {
            get
            {
                return (XComponentLoader)GetService("com.sun.star.frame.Desktop");
            }
        }

        public virtual XFileIdentifierConverter FileContentProvider
        {
            get
            {
                return (XFileIdentifierConverter)GetService("com.sun.star.ucb.FileContentProvider");
            }
        }

        public virtual XBridge Bridge
        {
            get
            {
                return _bridge;
            }
        }

        public virtual XMultiComponentFactory RemoteServiceManager
        {
            get
            {
                return _serviceManager;
            }
        }

        public virtual XComponentContext ComponentContext
        {
            get
            {
                return _componentContext;
            }
        }

        
    }
}
