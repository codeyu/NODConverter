using System;
using System.Collections;
using System.IO;
using Dotnet.Commons.IO;
using NODConverter.OpenOffice.Connection;
using Slf;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.task;
using unoidl.com.sun.star.ucb;
using unoidl.com.sun.star.util;

namespace NODConverter.OpenOffice.Converter
{
    using IOUtils = StreamUtils;
    using Logger = ILogger;

    /// <summary>
    /// Default file-based <seealso cref="AbstractOpenOfficeDocumentConverter"/> implementation.
    /// <p/>
    /// This implementation passes document data to and from the OpenOffice.org
    /// service as file URLs.
    /// <p/>
    /// File-based conversions are faster than stream-based ones (provided by
    /// <seealso cref="StreamOpenOfficeDocumentConverter"/>) but they require the
    /// OpenOffice.org service to be running locally and have the correct
    /// permissions to the files.
    /// </summary>
    /// <seealso cref= "StreamOpenOfficeDocumentConverter"> </seealso>
    public class OpenOfficeDocumentConverter : AbstractOpenOfficeDocumentConverter
    {

        
        private static readonly Logger Logger;

        public OpenOfficeDocumentConverter(IOpenOfficeConnection connection)
            : base(connection)
        {
        }

        public OpenOfficeDocumentConverter(IOpenOfficeConnection connection, IDocumentFormatRegistry formatRegistry)
            : base(connection, formatRegistry)
        {
        }

        /// <summary> In this file-based implementation, streams are emulated using temporary files.</summary>
        protected internal override void ConvertInternal(Stream inputStream, DocumentFormat inputFormat, Stream outputStream, DocumentFormat outputFormat)
        {
            FileInfo inputFile = null;
            FileInfo outputFile = null;
            try
            {

                inputFile = new FileInfo(string.Format("document.{0}", inputFormat.FileExtension))
                {
                    Attributes = FileAttributes.Temporary
                };
               
                Stream inputFileStream = new FileStream(inputFile.FullName, FileMode.Create);
                IOUtils.Copy(inputStream, inputFileStream);

                
                outputFile = new FileInfo(string.Format("document.{0}", outputFormat.FileExtension))
                {
                    Attributes = FileAttributes.Temporary
                };
                Convert(inputFile, inputFormat, outputFile, outputFormat);
                Stream outputFileStream = null;
                try
                {
                    
                    outputFileStream = new FileStream(outputFile.FullName, FileMode.Open, FileAccess.Read);
                    IOUtils.Copy(outputFileStream, outputStream);
                }
                finally
                {
                    //IOUtils.closeQuietly(outputFileStream);
                    if (outputFileStream != null) outputFileStream.Close();
                }
            }
            catch (IOException ioException)
            {
                throw new OpenOfficeException("conversion failed", ioException);
            }
            finally
            {
                if (inputFile != null)
                {
                    if (File.Exists(inputFile.FullName))
                    {
                        File.Delete(inputFile.FullName);
                    }
                    else if (Directory.Exists(inputFile.FullName))
                    {
                        Directory.Delete(inputFile.FullName);
                    }
                }
                if (outputFile != null)
                {
                    if (File.Exists(outputFile.FullName))
                    {
                        File.Delete(outputFile.FullName);
                    }
                    else if (Directory.Exists(outputFile.FullName))
                    {
                        Directory.Delete(outputFile.FullName);
                    }
                }
            }
        }

        protected internal override void ConvertInternal(FileInfo inputFile, DocumentFormat inputFormat, FileInfo outputFile, DocumentFormat outputFormat)
        {
            
            IDictionary loadProperties = new Hashtable();
            SupportClass.MapSupport.PutAll(loadProperties, DefaultLoadProperties);
            SupportClass.MapSupport.PutAll(loadProperties, inputFormat.ImportOptions);

            IDictionary storeProperties = outputFormat.GetExportOptions(inputFormat.Family);

            lock (OpenOfficeConnection)
            {
                XFileIdentifierConverter fileContentProvider = OpenOfficeConnection.FileContentProvider;
                String inputUrl = fileContentProvider.getFileURLFromSystemPath("", inputFile.FullName);
                String outputUrl = fileContentProvider.getFileURLFromSystemPath("", outputFile.FullName);

                LoadAndExport(inputUrl, loadProperties, outputUrl, storeProperties);
            }
        }

        private void LoadAndExport(String inputUrl, IDictionary loadProperties, String outputUrl, IDictionary storeProperties)
        {
            XComponent document;
            try
            {
                document = LoadDocument(inputUrl, loadProperties);
            }
            catch (ErrorCodeIOException errorCodeIoException)
            {
                throw new OpenOfficeException("conversion failed: could not load input document; OOo errorCode: " + errorCodeIoException.ErrCode, errorCodeIoException);
            }
            catch (Exception otherException)
            {
                throw new OpenOfficeException("conversion failed: could not load input document", otherException);
            }
            if (document == null)
            {
                throw new OpenOfficeException("conversion failed: could not load input document");
            }

            try
            {
                RefreshDocument(document);//使用此函数后Excel2pdf 提示异常：无法将透明代理强制转换为类型XRefreshable
            }
            catch
            {
                // ignored
            }

            try
            {
                storeDocument(document, outputUrl, storeProperties);
            }
            catch (ErrorCodeIOException errorCodeIoException)
            {
                throw new OpenOfficeException("conversion failed: could not save output document; OOo errorCode: " + errorCodeIoException.ErrCode, errorCodeIoException);
            }
            catch (Exception otherException)
            {
                throw new OpenOfficeException("conversion failed: could not save output document", otherException);
            }
        }

        private XComponent LoadDocument(String inputUrl, IDictionary loadProperties)
        {
            XComponentLoader desktop = OpenOfficeConnection.Desktop;
            //PropertyValue[] propertyValue2 = new PropertyValue[1];
            //PropertyValue aProperty = new PropertyValue();
            //aProperty.Name = "Hidden";
            //aProperty.Value = new uno.Any(true);
            //propertyValue2[0] = aProperty;
            ToPropertyValues(loadProperties);
            return desktop.loadComponentFromURL(inputUrl, "_blank", 0, ToPropertyValues(loadProperties));//toPropertyValues(loadProperties)
        }

        private void storeDocument(XComponent document, String outputUrl, IDictionary storeProperties)
        {
            try
            {
                XStorable storable = (XStorable)document;
                storable.storeToURL(outputUrl, ToPropertyValues(storeProperties));
            }
            finally
            {
                XCloseable closeable = (XCloseable)document;
                if (closeable != null)
                {
                    try
                    {
                        closeable.close(true);
                    }
                    catch (CloseVetoException closeVetoException)
                    {
                        Logger.Warn("document.close() vetoed:"+closeVetoException.Message);
                    }
                }
                else
                {
                    document.dispose();
                }
            }
        }
        static OpenOfficeDocumentConverter()
        {
            Logger = LoggerFactory.GetLogger();
        }
    }
}
