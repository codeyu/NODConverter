using System;
using System.Collections.Generic;
using System.Text;

namespace NODConverter
{
    public interface IDocumentConverter
    {

        /// <summary> Convert a document.
        /// <p>
        /// Note that this method does not close <tt>inputStream</tt> and <tt>outputStream</tt>.
        /// 
        /// </summary>
        /// <param name="inputStream">
        /// </param>
        /// <param name="inputFormat">
        /// </param>
        /// <param name="outputStream">
        /// </param>
        /// <param name="outputFormat">
        /// </param>
        void convert(System.IO.Stream inputStream, DocumentFormat inputFormat, System.IO.Stream outputStream, DocumentFormat outputFormat);

        /// <summary> Convert a document.
        /// 
        /// </summary>
        /// <param name="inputFile">
        /// </param>
        /// <param name="inputFormat">
        /// </param>
        /// <param name="outputFile">
        /// </param>
        /// <param name="outputFormat">
        /// </param>
        void convert(System.IO.FileInfo inputFile, DocumentFormat inputFormat, System.IO.FileInfo outputFile, DocumentFormat outputFormat);


        /// <summary> Convert a document. The input format is guessed from the file extension.
        /// 
        /// </summary>
        /// <param name="inputDocument">
        /// </param>
        /// <param name="outputDocument">
        /// </param>
        /// <param name="outputFormat">
        /// </param>
        void convert(System.IO.FileInfo inputDocument, System.IO.FileInfo outputDocument, DocumentFormat outputFormat);

        /// <summary> Convert a document. Both input and output formats are guessed from the file extension.
        /// 
        /// </summary>
        /// <param name="inputDocument">
        /// </param>
        /// <param name="outputDocument">
        /// </param>
        void convert(System.IO.FileInfo inputDocument, System.IO.FileInfo outputDocument);

    }
}
