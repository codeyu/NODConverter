using unoidl.com.sun.star.bridge;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.ucb;
using unoidl.com.sun.star.uno;
namespace NODConverter.OpenOffice.Connection
{
    public interface IOpenOfficeConnection
    {
        void Connect();

        void Disconnect();

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
