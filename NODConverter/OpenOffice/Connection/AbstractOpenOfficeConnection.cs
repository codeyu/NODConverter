using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.CompilerServices;
using System.Net;
using System.Net.Sockets;
namespace NODConverter.OpenOffice.Connection
{
    using Logger = Slf.ILogger;

    using XPropertySet = unoidl.com.sun.star.beans.XPropertySet;
    using XBridge = unoidl.com.sun.star.bridge.XBridge;
    using XBridgeFactory = unoidl.com.sun.star.bridge.XBridgeFactory;
    using Bootstrap = uno.util.Bootstrap;
    using NoConnectException = unoidl.com.sun.star.connection.NoConnectException;
    using XConnection = unoidl.com.sun.star.connection.XConnection;
    using XConnector = unoidl.com.sun.star.connection.XConnector;
    using XComponentLoader = unoidl.com.sun.star.frame.XComponentLoader;
    using EventObject = unoidl.com.sun.star.lang.EventObject;
    using XComponent = unoidl.com.sun.star.lang.XComponent;
    using XEventListener = unoidl.com.sun.star.lang.XEventListener;
    using XMultiComponentFactory = unoidl.com.sun.star.lang.XMultiComponentFactory;
    using XFileIdentifierConverter = unoidl.com.sun.star.ucb.XFileIdentifierConverter;
    using XComponentContext = unoidl.com.sun.star.uno.XComponentContext;
    using XUnoUrlResolver = unoidl.com.sun.star.bridge.XUnoUrlResolver;
    using XMultiServiceFactory = unoidl.com.sun.star.lang.XMultiServiceFactory;
    
    public abstract class AbstractOpenOfficeConnection : IOpenOfficeConnection, unoidl.com.sun.star.lang.XEventListener
    {

        protected internal readonly Logger logger = LoggerFactory.GetLogger();

        private string connectionString;
        private XComponent bridgeComponent;
        private XMultiComponentFactory serviceManager;
        private XComponentContext componentContext;
        private XBridge bridge;
        private bool connected = false;
        private bool expectingDisconnection = false;
        protected internal AbstractOpenOfficeConnection(string connectionString)
        {
            this.connectionString = connectionString;
        }
        [MethodImpl(MethodImplOptions.Synchronized)]
        public virtual void connect()
        {
            logger.Debug("connecting");
            try
            {
                Socket sock = new Socket(AddressFamily.InterNetwork, SocketType.Raw, ProtocolType.IP);
                XComponentContext localContext = uno.util.Bootstrap.defaultBootstrap_InitialComponentContext();
                XMultiComponentFactory localServiceManager = localContext.getServiceManager();
                XConnector connector = (XConnector)localServiceManager.createInstanceWithContext("com.sun.star.connection.Connector", localContext);
                XConnection connection = connector.connect(connectionString);
                XBridgeFactory bridgeFactory = (XBridgeFactory)localServiceManager.createInstanceWithContext("com.sun.star.bridge.BridgeFactory", localContext);
                bridge = bridgeFactory.createBridge("", "urp", connection, null);
                bridgeComponent = (XComponent)bridge;
                bridgeComponent.addEventListener(this);
                serviceManager = (XMultiComponentFactory)bridge.getInstance("StarOffice.ServiceManager");
                //另一种连接方式------------begin
                //XUnoUrlResolver xUrlResolver = (XUnoUrlResolver)localServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext);

                //// able to connect successfully 
                //XMultiServiceFactory multiServiceFactory = (XMultiServiceFactory)xUrlResolver.resolve("uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager"); 
                //---------------------------------------end


                XPropertySet properties = (XPropertySet)serviceManager;
                // Get the default context from the office server. 
                uno.Any oDefaultContext = properties.getPropertyValue("DefaultContext");

                componentContext = (XComponentContext)oDefaultContext.Value;
                connected = true;
                logger.Info("connected");
            }
            catch (NoConnectException connectException)
            {
                throw new OpenOfficeException("connection failed: " + connectionString + ": " + connectException.Message);
            }
            catch (Exception exception)
            {
                throw new OpenOfficeException("connection failed: " + connectionString, exception);
            }
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        public virtual void disconnect()
        {
            logger.Debug("disconnecting");
            expectingDisconnection = true;
            bridgeComponent.dispose();
        }

        public virtual bool Connected
        {
            get
            {
                return connected;
            }
        }

        public virtual void disposing(EventObject @event)
        {
            connected = false;
            if (expectingDisconnection)
            {
                logger.Info("disconnected");
            }
            else
            {
                logger.Error("disconnected unexpectedly");
            }
            expectingDisconnection = false;
        }

        // for unit tests only
        internal virtual void simulateUnexpectedDisconnection()
        {
            disposing(null);
            bridgeComponent.dispose();
        }

        private object getService(string className)
        {
            try
            {
                if (!connected)
                {
                    logger.Info("trying to (re)connect");
                    connect();
                }
                return serviceManager.createInstanceWithContext(className, componentContext);
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
                return (XComponentLoader)getService("com.sun.star.frame.Desktop");
            }
        }

        public virtual XFileIdentifierConverter FileContentProvider
        {
            get
            {
                return (XFileIdentifierConverter)getService("com.sun.star.ucb.FileContentProvider");
            }
        }

        public virtual XBridge Bridge
        {
            get
            {
                return bridge;
            }
        }

        public virtual XMultiComponentFactory RemoteServiceManager
        {
            get
            {
                return serviceManager;
            }
        }

        public virtual XComponentContext ComponentContext
        {
            get
            {
                return componentContext;
            }
        }
    }
}
