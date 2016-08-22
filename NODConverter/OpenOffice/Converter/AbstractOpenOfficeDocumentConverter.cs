using System;
using System.Collections;
using System.IO;
using NODConverter.OpenOffice.Connection;
using uno;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.util;

namespace NODConverter.OpenOffice.Converter
{
    public abstract class AbstractOpenOfficeDocumentConverter : IDocumentConverter
    {
        virtual protected internal IDictionary DefaultLoadProperties
        {
            get
            {
                return _defaultLoadProperties;
            }

        }
        virtual protected internal IDocumentFormatRegistry DocumentFormatRegistry
        {
            get
            {
                return _documentFormatRegistry;
            }

        }

        
        private readonly IDictionary _defaultLoadProperties;

        protected internal IOpenOfficeConnection OpenOfficeConnection;
        private readonly IDocumentFormatRegistry _documentFormatRegistry;

        protected AbstractOpenOfficeDocumentConverter(IOpenOfficeConnection connection)
            : this(connection, new DefaultDocumentFormatRegistry())
        {
        }

        protected AbstractOpenOfficeDocumentConverter(IOpenOfficeConnection openOfficeConnection, IDocumentFormatRegistry documentFormatRegistry)
        {
            OpenOfficeConnection = openOfficeConnection;
            _documentFormatRegistry = documentFormatRegistry;

            
            _defaultLoadProperties = new Hashtable();
            _defaultLoadProperties["Hidden"] = true;
            _defaultLoadProperties["ReadOnly"] = true;
        }

        public virtual void SetDefaultLoadProperty(String name, Object valueRenamed)
        {
            _defaultLoadProperties[name] = valueRenamed;
        }

        public virtual void Convert(FileInfo inputFile, FileInfo outputFile)
        {
            Convert(inputFile, outputFile, null);
        }

        public virtual void Convert(FileInfo inputFile, FileInfo outputFile, DocumentFormat outputFormat)
        {
            Convert(inputFile, null, outputFile, outputFormat);
        }

        public virtual void Convert(Stream inputStream, DocumentFormat inputFormat, Stream outputStream, DocumentFormat outputFormat)
        {
            EnsureNotNull("inputStream", inputStream);
            EnsureNotNull("inputFormat", inputFormat);
            EnsureNotNull("outputStream", outputStream);
            EnsureNotNull("outputFormat", outputFormat);
            ConvertInternal(inputStream, inputFormat, outputStream, outputFormat);
        }

        public virtual void Convert(FileInfo inputFile, DocumentFormat inputFormat, FileInfo outputFile, DocumentFormat outputFormat)
        {
            EnsureNotNull("inputFile", inputFile);
            EnsureNotNull("outputFile", outputFile);

            var tmpBool = File.Exists(inputFile.FullName) || Directory.Exists(inputFile.FullName);
            if (!tmpBool)
            {
                throw new ArgumentException("inputFile doesn't exist: " + inputFile);
            }
            if (inputFormat == null)
            {
                inputFormat = GuessDocumentFormat(inputFile);
            }
            if (outputFormat == null)
            {
                outputFormat = GuessDocumentFormat(outputFile);
            }
            if (!inputFormat.Importable)
            {
                throw new ArgumentException("unsupported input format: " + inputFormat.Name);
            }
            if (!inputFormat.IsExportableTo(outputFormat))
            {
                throw new ArgumentException("unsupported conversion: from " + inputFormat.Name + " to " + outputFormat.Name);
            }
            ConvertInternal(inputFile, inputFormat, outputFile, outputFormat);
        }

        protected internal abstract void ConvertInternal(Stream inputStream, DocumentFormat inputFormat, Stream outputStream, DocumentFormat outputFormat);

        protected internal abstract void ConvertInternal(FileInfo inputFile, DocumentFormat inputFormat, FileInfo outputFile, DocumentFormat outputFormat);

        private static void EnsureNotNull(string argumentName, object argumentValue)
        {
            if (argumentValue == null)
            {
                throw new ArgumentException(argumentName + " is null");
            }
        }

        private DocumentFormat GuessDocumentFormat(FileInfo file)
        {
            //System.String extension = FilenameUtils.getExtension(file.Name);
            //System.String extension = System.IO.Path.GetExtension(file.Name);
            String extension = file.Extension.Substring(1);
            DocumentFormat format = DocumentFormatRegistry.GetFormatByFileExtension(extension);
            if (format == null)
            {
                throw new ArgumentException("unknown document format for file: " + file);
            }
            return format;
        }

        protected internal virtual void RefreshDocument(XComponent document)
        {
            XRefreshable refreshable = (XRefreshable)document;
            if (refreshable != null)
            {
                refreshable.refresh();
            }
        }

        protected internal static PropertyValue Property(String name, Object valueRenamed)
        {
            PropertyValue property = new PropertyValue
            {
                Name = name,
                Value = new Any(valueRenamed.GetType(), valueRenamed)
            };
            //property.Value = new uno.Any("writer_pdf_Export"); value_Renamed;
            return property;
        }

        protected internal static PropertyValue[] ToPropertyValues(IDictionary properties)
        {
            PropertyValue[] propertyValues = new PropertyValue[properties.Count];
            int i = 0;
            
            for (IEnumerator iter = new SupportClass.HashSetSupport(properties).GetEnumerator(); iter.MoveNext(); )
            {
                if (iter.Current != null)
                {
                    DictionaryEntry entry = (DictionaryEntry)iter.Current;
                    Object valueRenamed = entry.Value;
                    var renamed = valueRenamed as IDictionary;
                    if (renamed != null)
                    {
                        // recursively Convert nested Map to PropertyValue[]
                        IDictionary subProperties = renamed;
                        valueRenamed = ToPropertyValues(subProperties);
                    }
                    propertyValues[i++] = Property((String)entry.Key, valueRenamed);
                }
            }
            return propertyValues;
        }
    }
}
