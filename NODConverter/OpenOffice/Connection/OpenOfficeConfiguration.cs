using System;
using System.Collections.Generic;
using System.Text;

namespace NODConverter.OpenOffice.Connection
{
    using PropertyValue = unoidl.com.sun.star.beans.PropertyValue;
    using XNameAccess = unoidl.com.sun.star.container.XNameAccess;
    using XMultiServiceFactory = unoidl.com.sun.star.lang.XMultiServiceFactory;

    /// <summary>
    /// Utility class to access OpenOffice.org configuration properties at runtime
    /// </summary>
    public class OpenOfficeConfiguration
    {

        public const string NODE_L10N = "org.openoffice.Setup/L10N";
        public const string NODE_PRODUCT = "org.openoffice.Setup/Product";

        private IOpenOfficeConnection connection;

        public OpenOfficeConfiguration(IOpenOfficeConnection connection)
        {
            this.connection = connection;
        }

        public virtual string getOpenOfficeProperty(string nodePath, string node)
        {
            if (!nodePath.StartsWith("/"))
            {
                nodePath = "/" + nodePath;
            }
            uno.Any property;
            // create the provider and remember it as a XMultiServiceFactory
            try
            {
                const string sProviderService = "com.sun.star.configuration.ConfigurationProvider";
                object configProvider = connection.RemoteServiceManager.createInstanceWithContext(sProviderService, connection.ComponentContext);
                XMultiServiceFactory xConfigProvider = (XMultiServiceFactory)configProvider;

                // The service name: Need only read access:
                const string sReadOnlyView = "com.sun.star.configuration.ConfigurationAccess";
                // creation arguments: nodepath
                PropertyValue aPathArgument = new PropertyValue();
                aPathArgument.Name = "nodepath";
                aPathArgument.Value = new uno.Any(nodePath);
                uno.Any[] aArguments = new uno.Any[1];
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
                    return getOpenOfficeProperty(NODE_PRODUCT, "ooSetupVersionAboutBox");
                }
                catch (OpenOfficeException noSuchElementException)
                {
                    // OOo < 2.2 only returns major.minor
                    return getOpenOfficeProperty(NODE_PRODUCT, "ooSetupVersion");
                }
            }
        }

        public virtual string OpenOfficeLocale
        {
            get
            {
                return getOpenOfficeProperty(NODE_L10N, "ooLocale");
            }
        }

    }
}
