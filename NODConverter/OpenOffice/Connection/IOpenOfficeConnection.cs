using System;
using System.Collections.Generic;
using System.Text;

namespace NODConverter.OpenOffice.Connection
{
    using XBridge = unoidl.com.sun.star.bridge.XBridge;
    using XComponentLoader = unoidl.com.sun.star.frame.XComponentLoader;
    using XMultiComponentFactory = unoidl.com.sun.star.lang.XMultiComponentFactory;
    using XFileIdentifierConverter = unoidl.com.sun.star.ucb.XFileIdentifierConverter;
    using XComponentContext = unoidl.com.sun.star.uno.XComponentContext;

    public interface IOpenOfficeConnection
    {
        void connect();

        void disconnect();

        bool Connected { get; }

        /// <returns> the com.sun.star.frame.Desktop service </returns>
        XComponentLoader Desktop { get; }

        /// <returns> the com.sun.star.ucb.FileContentProvider service </returns>
        XFileIdentifierConverter FileContentProvider { get; }

        XBridge Bridge { get; }

        XMultiComponentFactory RemoteServiceManager { get; }

        XComponentContext ComponentContext { get; }
    }
}
