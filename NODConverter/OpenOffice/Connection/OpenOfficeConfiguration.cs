using System;
using uno;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.lang;

namespace NODConverter.OpenOffice.Connection
{
    /// <summary>
    /// Utility class to access OpenOffice.org configuration properties at runtime
    /// </summary>
    public class OpenOfficeConfiguration
    {

        public const string NodeL10N = "org.openoffice.Setup/L10N";
        public const string NodeProduct = "org.openoffice.Setup/Product";

        private readonly IOpenOfficeConnection _connection;

        public OpenOfficeConfiguration(IOpenOfficeConnection connection)
        {
            _connection = connection;
        }

        public virtual string GetOpenOfficeProperty(string nodePath, string node)
        {
            if (!nodePath.StartsWith("/"))
            {
                nodePath = "/" + nodePath;
            }
            Any property;
            // create the provider and remember it as a XMultiServiceFactory
            try
            {
                const string sProviderService = "com.sun.star.configuration.ConfigurationProvider";
                object configProvider = _connection.RemoteServiceManager.createInstanceWithContext(sProviderService, _connection.ComponentContext);
                XMultiServiceFactory xConfigProvider = (XMultiServiceFactory)configProvider;

                // The service name: Need only read access:
                const string sReadOnlyView = "com.sun.star.configuration.ConfigurationAccess";
                // creation arguments: nodepath
                PropertyValue aPathArgument = new PropertyValue
                {
                    Name = "nodepath",
                    Value = new Any(nodePath)
                };
                Any[] aArguments = new Any[1];
                aArguments[0] = aPathArgument.Value;

                // create the view
                //XInterface xElement = (XInterface) xConfigProvider.createInstanceWithArguments(sReadOnlyView, aArguments);
                XNameAccess xChildAccess = (XNameAccess)xConfigProvider.createInstanceWithArguments(sReadOnlyView, aArguments);

                // get the value
                property = xChildAccess.getByName(node);
            }
            catch (Exception exception)
            {
                throw new OpenOfficeException("Could not retrieve property", exception);
            }
            return property.ToString();
        }

        public virtual string OpenOfficeVersion
        {
            get
            {
                try
                {
                    // OOo >= 2.2 returns major.minor.micro
                    return GetOpenOfficeProperty(NodeProduct, "ooSetupVersionAboutBox");
                }
                catch (OpenOfficeException)
                {
                    // OOo < 2.2 only returns major.minor
                    return GetOpenOfficeProperty(NodeProduct, "ooSetupVersion");
                }
            }
        }

        public virtual string OpenOfficeLocale
        {
            get
            {
                return GetOpenOfficeProperty(NodeL10N, "ooLocale");
            }
        }

    }
}
