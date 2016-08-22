using System;
using System.Collections;
using System.IO;
using NODConverter.OpenOffice.Connection;
using uno;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.io;
using unoidl.com.sun.star.lang;

namespace NODConverter.OpenOffice.Converter
{
    /// <summary>
    /// Alternative stream-based <seealso cref="DocumentConverter"/> implementation.
    /// <p/>
    /// This implementation passes document data to and from the OpenOffice.org
    /// service as streams.
    /// <p/>
    /// Stream-based conversions are slower than the default file-based ones (provided
    /// by <seealso cref="OpenOfficeDocumentConverter"/>) but they allow to run the OpenOffice.org
    /// service on a different machine, or under a different system user on the same
    /// machine without file permission problems.
    /// <p/>
    /// <b>Warning!</b> There is a <a href="http://www.openoffice.org/issues/show_bug.cgi?id=75519">bug</a>
    /// in OpenOffice.org 2.2.x that causes some input formats, including Word and Excel, not to work when
    /// using stream-base conversions.
    /// <p/>
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

        protected internal override void ConvertInternal(FileInfo inputFile, DocumentFormat inputFormat, FileInfo outputFile, DocumentFormat outputFormat)
        {
            Stream inputStream = null;
            Stream outputStream = null;
            try
            {
                
                inputStream = new FileStream(inputFile.FullName, FileMode.Open, FileAccess.Read);
                
                outputStream = new FileStream(outputFile.FullName, FileMode.Create);
                Convert(inputStream, inputFormat, outputStream, outputFormat);
            }
            catch (FileNotFoundException fileNotFoundException)
            {
               
                throw new ArgumentException(fileNotFoundException.Message);
            }
            finally
            {
                if (inputStream != null) inputStream.Close();
                if (outputStream != null) outputStream.Close();
            }
        }

        protected internal override void ConvertInternal(Stream inputStream, DocumentFormat inputFormat, Stream outputStream, DocumentFormat outputFormat)
        {
            IDictionary exportOptions = outputFormat.GetExportOptions(inputFormat.Family);

            try
            {
                lock (OpenOfficeConnection)
                {
                    LoadAndExport(inputStream, inputFormat.ImportOptions, outputStream, exportOptions);
                }
            }
            catch (OpenOfficeException)
            {
                throw;
            }
           
            catch (Exception throwable)
            {
                throw new OpenOfficeException("conversion failed", throwable);
            }
        }

        private void LoadAndExport(Stream inputStream, IDictionary importOptions, Stream outputStream, IDictionary exportOptions)
        {
            XComponentLoader desktop = OpenOfficeConnection.Desktop;

            
            IDictionary loadProperties = new Hashtable();
            SupportClass.MapSupport.PutAll(loadProperties, DefaultLoadProperties);
            SupportClass.MapSupport.PutAll(loadProperties, importOptions);
            // doesn't work using InputStreamToXInputStreamAdapter; probably because it's not XSeekable 
            //property("InputStream", new InputStreamToXInputStreamAdapter(inputStream))
            loadProperties["InputStream"] = new Any(typeof(XInputStream), new XInputStreamWrapper(inputStream)); 
            
            XComponent document = desktop.loadComponentFromURL("private:stream", "_blank", 0, ToPropertyValues(loadProperties));
            if (document == null)
            {
                throw new OpenOfficeException("conversion failed: input document is null after loading");
            }

            RefreshDocument(document);

            
            IDictionary storeProperties = new Hashtable();
            SupportClass.MapSupport.PutAll(storeProperties, exportOptions);
            storeProperties["OutputStream"] = new Any(typeof(XOutputStream), new XOutputStreamWrapper(outputStream)); 

            try
            {
                XStorable storable = (XStorable)document;
                storable.storeToURL("private:stream", ToPropertyValues(storeProperties));
            }
            finally
            {
                document.dispose();
            }
        }
    }
}
