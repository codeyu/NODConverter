using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using net.sf.dotnetcli;
using NODConverter.OpenOffice.Connection;
namespace NODConverter.OpenOffice.Converter
{
    using PropertyValue = unoidl.com.sun.star.beans.PropertyValue;
    using XComponent = unoidl.com.sun.star.lang.XComponent;
    using XRefreshable = unoidl.com.sun.star.util.XRefreshable;

    public abstract class AbstractOpenOfficeDocumentConverter : IDocumentConverter
    {
        virtual protected internal System.Collections.IDictionary DefaultLoadProperties
        {
            get
            {
                return defaultLoadProperties;
            }

        }
        virtual protected internal IDocumentFormatRegistry DocumentFormatRegistry
        {
            get
            {
                return documentFormatRegistry;
            }

        }

        
        private System.Collections.IDictionary defaultLoadProperties;

        protected internal IOpenOfficeConnection openOfficeConnection;
        private IDocumentFormatRegistry documentFormatRegistry;

        public AbstractOpenOfficeDocumentConverter(IOpenOfficeConnection connection)
            : this(connection, new DefaultDocumentFormatRegistry())
        {
        }

        public AbstractOpenOfficeDocumentConverter(IOpenOfficeConnection openOfficeConnection, IDocumentFormatRegistry documentFormatRegistry)
        {
            this.openOfficeConnection = openOfficeConnection;
            this.documentFormatRegistry = documentFormatRegistry;

            //UPGRADE_TODO: Class“java.util.HashMap”被转换为具有不同行为的 'System.Collections.Hashtable'。 "ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?index='!DefaultContextWindowIndex'&keyword='jlca1073_javautilHashMap'"
            defaultLoadProperties = new System.Collections.Hashtable();
            defaultLoadProperties["Hidden"] = true;
            defaultLoadProperties["ReadOnly"] = true;
        }

        public virtual void setDefaultLoadProperty(System.String name, System.Object value_Renamed)
        {
            defaultLoadProperties[name] = value_Renamed;
        }

        public virtual void convert(System.IO.FileInfo inputFile, System.IO.FileInfo outputFile)
        {
            convert(inputFile, outputFile, null);
        }

        public virtual void convert(System.IO.FileInfo inputFile, System.IO.FileInfo outputFile, DocumentFormat outputFormat)
        {
            convert(inputFile, null, outputFile, outputFormat);
        }

        public virtual void convert(System.IO.Stream inputStream, DocumentFormat inputFormat, System.IO.Stream outputStream, DocumentFormat outputFormat)
        {
            ensureNotNull("inputStream", inputStream);
            ensureNotNull("inputFormat", inputFormat);
            ensureNotNull("outputStream", outputStream);
            ensureNotNull("outputFormat", outputFormat);
            convertInternal(inputStream, inputFormat, outputStream, outputFormat);
        }

        public virtual void convert(System.IO.FileInfo inputFile, DocumentFormat inputFormat, System.IO.FileInfo outputFile, DocumentFormat outputFormat)
        {
            ensureNotNull("inputFile", inputFile);
            ensureNotNull("outputFile", outputFile);

            bool tmpBool;
            if (System.IO.File.Exists(inputFile.FullName))
                tmpBool = true;
            else
                tmpBool = System.IO.Directory.Exists(inputFile.FullName);
            if (!tmpBool)
            {
                throw new System.ArgumentException("inputFile doesn't exist: " + inputFile);
            }
            if (inputFormat == null)
            {
                inputFormat = guessDocumentFormat(inputFile);
            }
            if (outputFormat == null)
            {
                outputFormat = guessDocumentFormat(outputFile);
            }
            if (!inputFormat.Importable)
            {
                throw new System.ArgumentException("unsupported input format: " + inputFormat.Name);
            }
            if (!inputFormat.isExportableTo(outputFormat))
            {
                throw new System.ArgumentException("unsupported conversion: from " + inputFormat.Name + " to " + outputFormat.Name);
            }
            convertInternal(inputFile, inputFormat, outputFile, outputFormat);
        }

        protected internal abstract void convertInternal(System.IO.Stream inputStream, DocumentFormat inputFormat, System.IO.Stream outputStream, DocumentFormat outputFormat);

        protected internal abstract void convertInternal(System.IO.FileInfo inputFile, DocumentFormat inputFormat, System.IO.FileInfo outputFile, DocumentFormat outputFormat);

        private void ensureNotNull(System.String argumentName, System.Object argumentValue)
        {
            if (argumentValue == null)
            {
                throw new System.ArgumentException(argumentName + " is null");
            }
        }

        private DocumentFormat guessDocumentFormat(System.IO.FileInfo file)
        {
            //System.String extension = FilenameUtils.getExtension(file.Name);
            //System.String extension = System.IO.Path.GetExtension(file.Name);
            System.String extension = file.Extension.Substring(1);
            DocumentFormat format = DocumentFormatRegistry.getFormatByFileExtension(extension);
            if (format == null)
            {
                throw new System.ArgumentException("unknown document format for file: " + file);
            }
            return format;
        }

        protected internal virtual void refreshDocument(XComponent document)
        {
            XRefreshable refreshable = (XRefreshable)document;
            if (refreshable != null)
            {
                refreshable.refresh();
            }
        }

        protected internal static PropertyValue property(System.String name, System.Object value_Renamed)
        {
            PropertyValue property = new PropertyValue();
            property.Name = name;
            //property.Value = new uno.Any("writer_pdf_Export"); value_Renamed;
            property.Value = new uno.Any(value_Renamed.GetType(), value_Renamed);
            return property;
        }

        protected internal static PropertyValue[] toPropertyValues(System.Collections.IDictionary properties)
        {
            PropertyValue[] propertyValues = new PropertyValue[properties.Count];
            int i = 0;
            //UPGRADE_TODO: 方法“java.util.Map.entrySet”被转换为具有不同行为的 'SupportClass.HashSetSupport'。 "ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?index='!DefaultContextWindowIndex'&keyword='jlca1073_javautilMapentrySet'"
            //UPGRADE_TODO: 方法“java.util.Iterator.hasNext”被转换为具有不同行为的 'System.Collections.IEnumerator.MoveNext'。 "ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?index='!DefaultContextWindowIndex'&keyword='jlca1073_javautilIteratorhasNext'"
            for (System.Collections.IEnumerator iter = new SupportClass.HashSetSupport(properties).GetEnumerator(); iter.MoveNext(); )
            {
                //UPGRADE_TODO: 方法“java.util.Iterator.next”被转换为具有不同行为的 'System.Collections.IEnumerator.Current'。 "ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?index='!DefaultContextWindowIndex'&keyword='jlca1073_javautilIteratornext'"
                System.Collections.DictionaryEntry entry = (System.Collections.DictionaryEntry)iter.Current;
                System.Object value_Renamed = entry.Value;
                if (value_Renamed is System.Collections.IDictionary)
                {
                    // recursively convert nested Map to PropertyValue[]
                    System.Collections.IDictionary subProperties = (System.Collections.IDictionary)value_Renamed;
                    value_Renamed = toPropertyValues(subProperties);
                }
                propertyValues[i++] = property((System.String)entry.Key, value_Renamed);
            }
            return propertyValues;
        }
    }
}
