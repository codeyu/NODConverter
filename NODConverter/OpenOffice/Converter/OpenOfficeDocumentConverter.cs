using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using NODConverter.OpenOffice.Connection;
namespace NODConverter.OpenOffice.Converter
{
    using PropertyValue = unoidl.com.sun.star.beans.PropertyValue;
    using IOUtils = Dotnet.Commons.IO.StreamUtils;
    using Logger = Slf.ILogger;
    using XComponentLoader = unoidl.com.sun.star.frame.XComponentLoader;
    using XStorable = unoidl.com.sun.star.frame.XStorable;
    using IllegalArgumentException = unoidl.com.sun.star.lang.IllegalArgumentException;
    using XComponent = unoidl.com.sun.star.lang.XComponent;
    using ErrorCodeIOException = unoidl.com.sun.star.task.ErrorCodeIOException;
    using XFileIdentifierConverter = unoidl.com.sun.star.ucb.XFileIdentifierConverter;

    using CloseVetoException = unoidl.com.sun.star.util.CloseVetoException;
    using XCloseable = unoidl.com.sun.star.util.XCloseable;
    /// <summary>
    /// Default file-based <seealso cref="DocumentConverter"/> implementation.
    /// <p>
    /// This implementation passes document data to and from the OpenOffice.org
    /// service as file URLs.
    /// <p>
    /// File-based conversions are faster than stream-based ones (provided by
    /// <seealso cref="StreamOpenOfficeDocumentConverter"/>) but they require the
    /// OpenOffice.org service to be running locally and have the correct
    /// permissions to the files.
    /// </summary>
    /// <seealso cref= StreamOpenOfficeDocumentConverter </seealso>
    public class OpenOfficeDocumentConverter : AbstractOpenOfficeDocumentConverter
    {

        //UPGRADE_NOTE: Final 已从“logger ”的声明中移除。 "ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?index='!DefaultContextWindowIndex'&keyword='jlca1003'"
        //UPGRADE_NOTE: “logger”的初始化已移动到 static method“com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter”。 "ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?index='!DefaultContextWindowIndex'&keyword='jlca1005'"
        private static readonly Logger logger;

        public OpenOfficeDocumentConverter(IOpenOfficeConnection connection)
            : base(connection)
        {
        }

        public OpenOfficeDocumentConverter(IOpenOfficeConnection connection, IDocumentFormatRegistry formatRegistry)
            : base(connection, formatRegistry)
        {
        }

        /// <summary> In this file-based implementation, streams are emulated using temporary files.</summary>
        protected internal override void convertInternal(System.IO.Stream inputStream, DocumentFormat inputFormat, System.IO.Stream outputStream, DocumentFormat outputFormat)
        {
            System.IO.FileInfo inputFile = null;
            System.IO.FileInfo outputFile = null;
            try
            {
                //UPGRADE_ISSUE: 未转换 方法“java.io.File.createTempFile”。 "ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?index='!DefaultContextWindowIndex'&keyword='jlca1000_javaioFilecreateTempFile_javalangString_javalangString'"
                inputFile = new System.IO.FileInfo(string.Format("document.{0}", inputFormat.FileExtension));
                inputFile.Attributes = System.IO.FileAttributes.Temporary;
                System.IO.Stream inputFileStream = null;
                try
                {
                    //UPGRADE_TODO: 构造函数“java.io.FileOutputStream.FileOutputStream”被转换为具有不同行为的 'System.IO.FileStream.FileStream'。 "ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?index='!DefaultContextWindowIndex'&keyword='jlca1073_javaioFileOutputStreamFileOutputStream_javaioFile'"
                    inputFileStream = new System.IO.FileStream(inputFile.FullName, System.IO.FileMode.Create);
                    IOUtils.Copy(inputStream, inputFileStream);
                }
                finally
                {
                    //IOUtils.closeQuietly(inputFileStream);
                }

                //UPGRADE_ISSUE: 未转换 方法“java.io.File.createTempFile”。 "ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?index='!DefaultContextWindowIndex'&keyword='jlca1000_javaioFilecreateTempFile_javalangString_javalangString'"

                outputFile = new System.IO.FileInfo(string.Format("document.{0}", outputFormat.FileExtension));
                outputFile.Attributes = System.IO.FileAttributes.Temporary;
                convert(inputFile, inputFormat, outputFile, outputFormat);
                System.IO.Stream outputFileStream = null;
                try
                {
                    //UPGRADE_TODO: 构造函数“java.io.FileInputStream.FileInputStream”被转换为具有不同行为的 'System.IO.FileStream.FileStream'。 "ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?index='!DefaultContextWindowIndex'&keyword='jlca1073_javaioFileInputStreamFileInputStream_javaioFile'"
                    outputFileStream = new System.IO.FileStream(outputFile.FullName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                    IOUtils.Copy(outputFileStream, outputStream);
                }
                finally
                {
                    //IOUtils.closeQuietly(outputFileStream);
                    outputFileStream.Close();
                }
            }
            catch (System.IO.IOException ioException)
            {
                throw new OpenOfficeException("conversion failed", ioException);
            }
            finally
            {
                if (inputFile != null)
                {
                    bool tmpBool;
                    if (System.IO.File.Exists(inputFile.FullName))
                    {
                        System.IO.File.Delete(inputFile.FullName);
                        tmpBool = true;
                    }
                    else if (System.IO.Directory.Exists(inputFile.FullName))
                    {
                        System.IO.Directory.Delete(inputFile.FullName);
                        tmpBool = true;
                    }
                    else
                        tmpBool = false;
                    bool generatedAux = tmpBool;
                }
                if (outputFile != null)
                {
                    bool tmpBool2;
                    if (System.IO.File.Exists(outputFile.FullName))
                    {
                        System.IO.File.Delete(outputFile.FullName);
                        tmpBool2 = true;
                    }
                    else if (System.IO.Directory.Exists(outputFile.FullName))
                    {
                        System.IO.Directory.Delete(outputFile.FullName);
                        tmpBool2 = true;
                    }
                    else
                        tmpBool2 = false;
                    bool generatedAux2 = tmpBool2;
                }
            }
        }

        protected internal override void convertInternal(System.IO.FileInfo inputFile, DocumentFormat inputFormat, System.IO.FileInfo outputFile, DocumentFormat outputFormat)
        {
            //UPGRADE_TODO: Class“java.util.HashMap”被转换为具有不同行为的 'System.Collections.Hashtable'。 "ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?index='!DefaultContextWindowIndex'&keyword='jlca1073_javautilHashMap'"
            System.Collections.IDictionary loadProperties = new System.Collections.Hashtable();
            SupportClass.MapSupport.PutAll(loadProperties, DefaultLoadProperties);
            SupportClass.MapSupport.PutAll(loadProperties, inputFormat.ImportOptions);

            System.Collections.IDictionary storeProperties = outputFormat.getExportOptions(inputFormat.Family);

            lock (openOfficeConnection)
            {
                XFileIdentifierConverter fileContentProvider = openOfficeConnection.FileContentProvider;
                System.String inputUrl = fileContentProvider.getFileURLFromSystemPath("", inputFile.FullName);
                System.String outputUrl = fileContentProvider.getFileURLFromSystemPath("", outputFile.FullName);

                loadAndExport(inputUrl, loadProperties, outputUrl, storeProperties);
            }
        }

        private void loadAndExport(System.String inputUrl, System.Collections.IDictionary loadProperties, System.String outputUrl, System.Collections.IDictionary storeProperties)
        {
            XComponent document;
            try
            {
                document = loadDocument(inputUrl, loadProperties);
            }
            catch (ErrorCodeIOException errorCodeIOException)
            {
                throw new OpenOfficeException("conversion failed: could not load input document; OOo errorCode: " + errorCodeIOException.ErrCode, errorCodeIOException);
            }
            catch (System.Exception otherException)
            {
                throw new OpenOfficeException("conversion failed: could not load input document", otherException);
            }
            if (document == null)
            {
                throw new OpenOfficeException("conversion failed: could not load input document");
            }

            try
            {
                refreshDocument(document);//使用此函数后Excel2pdf 提示异常：无法将透明代理强制转换为类型XRefreshable
            }
            catch
            {

            }

            try
            {
                storeDocument(document, outputUrl, storeProperties);
            }
            catch (ErrorCodeIOException errorCodeIOException)
            {
                throw new OpenOfficeException("conversion failed: could not save output document; OOo errorCode: " + errorCodeIOException.ErrCode, errorCodeIOException);
            }
            catch (System.Exception otherException)
            {
                throw new OpenOfficeException("conversion failed: could not save output document", otherException);
            }
        }

        private XComponent loadDocument(System.String inputUrl, System.Collections.IDictionary loadProperties)
        {
            XComponentLoader desktop = openOfficeConnection.Desktop;
            //PropertyValue[] propertyValue2 = new PropertyValue[1];
            //PropertyValue aProperty = new PropertyValue();
            //aProperty.Name = "Hidden";
            //aProperty.Value = new uno.Any(true);
            //propertyValue2[0] = aProperty;
            PropertyValue[] propertyValue = toPropertyValues(loadProperties);
            return desktop.loadComponentFromURL(inputUrl, "_blank", 0, toPropertyValues(loadProperties));//toPropertyValues(loadProperties)
        }

        private void storeDocument(XComponent document, System.String outputUrl, System.Collections.IDictionary storeProperties)
        {
            try
            {
                XStorable storable = (XStorable)document;
                storable.storeToURL(outputUrl, toPropertyValues(storeProperties));
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
                        logger.Warn("document.close() vetoed:"+closeVetoException.Message);
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
            logger = LoggerFactory.GetLogger();
        }
    }
}
