using System;
using System.Collections.Generic;
using System.Text;
using NODConverter.OpenOffice.Connection;
namespace NODConverter.OpenOffice.Converter
{
    using XComponentLoader = unoidl.com.sun.star.frame.XComponentLoader;
    using XStorable = unoidl.com.sun.star.frame.XStorable;
    using XComponent = unoidl.com.sun.star.lang.XComponent;
    
    
    /// <summary>
    /// Alternative stream-based <seealso cref="DocumentConverter"/> implementation.
    /// <p>
    /// This implementation passes document data to and from the OpenOffice.org
    /// service as streams.
    /// <p>
    /// Stream-based conversions are slower than the default file-based ones (provided
    /// by <seealso cref="OpenOfficeDocumentConverter"/>) but they allow to run the OpenOffice.org
    /// service on a different machine, or under a different system user on the same
    /// machine without file permission problems.
    /// <p>
    /// <b>Warning!</b> There is a <a href="http://www.openoffice.org/issues/show_bug.cgi?id=75519">bug</a>
    /// in OpenOffice.org 2.2.x that causes some input formats, including Word and Excel, not to work when
    /// using stream-base conversions.
    /// <p>
    /// Use this implementation only if <seealso cref="OpenOfficeDocumentConverter"/> is not possible.
    /// </summary>
    /// <seealso cref= OpenOfficeDocumentConverter </seealso>
    public class StreamOpenOfficeDocumentConverter : AbstractOpenOfficeDocumentConverter
    {

        public StreamOpenOfficeDocumentConverter(IOpenOfficeConnection connection)
            : base(connection)
        {
            
        }

        public StreamOpenOfficeDocumentConverter(IOpenOfficeConnection connection, IDocumentFormatRegistry formatRegistry)
            : base(connection, formatRegistry)
        {
        }

        protected internal override void convertInternal(System.IO.FileInfo inputFile, DocumentFormat inputFormat, System.IO.FileInfo outputFile, DocumentFormat outputFormat)
        {
            System.IO.Stream inputStream = null;
            System.IO.Stream outputStream = null;
            try
            {
                
                inputStream = new System.IO.FileStream(inputFile.FullName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                
                outputStream = new System.IO.FileStream(outputFile.FullName, System.IO.FileMode.Create);
                convert(inputStream, inputFormat, outputStream, outputFormat);
            }
            catch (System.IO.FileNotFoundException fileNotFoundException)
            {
               
                throw new System.ArgumentException(fileNotFoundException.Message);
            }
            finally
            {
                inputStream.Close();
                outputStream.Close();
            }
        }

        protected internal override void convertInternal(System.IO.Stream inputStream, DocumentFormat inputFormat, System.IO.Stream outputStream, DocumentFormat outputFormat)
        {
            System.Collections.IDictionary exportOptions = outputFormat.getExportOptions(inputFormat.Family);

            try
            {
                lock (openOfficeConnection)
                {
                    loadAndExport(inputStream, inputFormat.ImportOptions, outputStream, exportOptions);
                }
            }
            catch (OpenOfficeException openOfficeException)
            {
                throw openOfficeException;
            }
           
            catch (System.Exception throwable)
            {
                throw new OpenOfficeException("conversion failed", throwable);
            }
        }

        private void loadAndExport(System.IO.Stream inputStream, System.Collections.IDictionary importOptions, System.IO.Stream outputStream, System.Collections.IDictionary exportOptions)
        {
            XComponentLoader desktop = openOfficeConnection.Desktop;

            
            System.Collections.IDictionary loadProperties = new System.Collections.Hashtable();
            SupportClass.MapSupport.PutAll(loadProperties, DefaultLoadProperties);
            SupportClass.MapSupport.PutAll(loadProperties, importOptions);
            // doesn't work using InputStreamToXInputStreamAdapter; probably because it's not XSeekable 
            //property("InputStream", new InputStreamToXInputStreamAdapter(inputStream))
            loadProperties["InputStream"] = new uno.Any(typeof(unoidl.com.sun.star.io.XInputStream), new XInputStreamWrapper(inputStream)); 
            
            XComponent document = desktop.loadComponentFromURL("private:stream", "_blank", 0, toPropertyValues(loadProperties));
            if (document == null)
            {
                throw new OpenOfficeException("conversion failed: input document is null after loading");
            }

            refreshDocument(document);

            
            System.Collections.IDictionary storeProperties = new System.Collections.Hashtable();
            SupportClass.MapSupport.PutAll(storeProperties, exportOptions);
            storeProperties["OutputStream"] = new uno.Any(typeof(unoidl.com.sun.star.io.XOutputStream), new XOutputStreamWrapper(outputStream)); 

            try
            {
                XStorable storable = (XStorable)document;
                storable.storeToURL("private:stream", toPropertyValues(storeProperties));
            }
            finally
            {
                document.dispose();
            }
        }
    }
}
