using System;
using System.Collections.Generic;
using System.Text;

namespace NODConverter
{
    public class DefaultDocumentFormatRegistry : BasicDocumentFormatRegistry
    {

        public DefaultDocumentFormatRegistry()
        {
            
            DocumentFormat pdf = new DocumentFormat("Portable Document Format", "application/pdf", "pdf");
            pdf.setExportFilter(DocumentFamily.DRAWING, "draw_pdf_Export");
            pdf.setExportFilter(DocumentFamily.PRESENTATION, "impress_pdf_Export");
            pdf.setExportFilter(DocumentFamily.SPREADSHEET, "calc_pdf_Export");
            pdf.setExportFilter(DocumentFamily.TEXT, "writer_pdf_Export");
            addDocumentFormat(pdf);

            
            DocumentFormat swf = new DocumentFormat("Macromedia Flash", "application/x-shockwave-flash", "swf");
            swf.setExportFilter(DocumentFamily.DRAWING, "draw_flash_Export");
            swf.setExportFilter(DocumentFamily.PRESENTATION, "impress_flash_Export");
            addDocumentFormat(swf);

            
            DocumentFormat xhtml = new DocumentFormat("XHTML", "application/xhtml+xml", "xhtml");
            xhtml.setExportFilter(DocumentFamily.PRESENTATION, "XHTML Impress File");
            xhtml.setExportFilter(DocumentFamily.SPREADSHEET, "XHTML Calc File");
            xhtml.setExportFilter(DocumentFamily.TEXT, "XHTML Writer File");
            addDocumentFormat(xhtml);

            // HTML is treated as Text when supplied as input, but as an output it is also
            // available for exporting Spreadsheet and Presentation formats
            
            DocumentFormat html = new DocumentFormat("HTML", DocumentFamily.TEXT, "text/html", "html");
            html.setExportFilter(DocumentFamily.PRESENTATION, "impress_html_Export");
            html.setExportFilter(DocumentFamily.SPREADSHEET, "HTML (StarCalc)");
            html.setExportFilter(DocumentFamily.TEXT, "HTML (StarWriter)");
            addDocumentFormat(html);

            
            DocumentFormat odt = new DocumentFormat("OpenDocument Text", DocumentFamily.TEXT, "application/vnd.oasis.opendocument.text", "odt");
            odt.setExportFilter(DocumentFamily.TEXT, "writer8");
            addDocumentFormat(odt);

            
            DocumentFormat sxw = new DocumentFormat("OpenOffice.org 1.0 Text Document", DocumentFamily.TEXT, "application/vnd.sun.xml.writer", "sxw");
            sxw.setExportFilter(DocumentFamily.TEXT, "StarOffice XML (Writer)");
            addDocumentFormat(sxw);

            
            DocumentFormat doc = new DocumentFormat("Microsoft Word", DocumentFamily.TEXT, "application/msword", "doc");
            doc.setExportFilter(DocumentFamily.TEXT, "MS Word 97");
            addDocumentFormat(doc);

           
            DocumentFormat docx = new DocumentFormat("Microsoft Word 2007 XML", DocumentFamily.TEXT, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "docx");
            addDocumentFormat(docx);

            
            DocumentFormat rtf = new DocumentFormat("Rich Text Format", DocumentFamily.TEXT, "text/rtf", "rtf");
            rtf.setExportFilter(DocumentFamily.TEXT, "Rich Text Format");
            addDocumentFormat(rtf);

           
            DocumentFormat wpd = new DocumentFormat("WordPerfect", DocumentFamily.TEXT, "application/wordperfect", "wpd");
            addDocumentFormat(wpd);

            
            DocumentFormat txt = new DocumentFormat("Plain Text", DocumentFamily.TEXT, "text/plain", "txt");
            // default to "Text (encoded)" UTF8,CRLF to prevent OOo from trying to display the "ASCII Filter Options" dialog
            txt.setImportOption("FilterName", "Text (encoded)");
            txt.setImportOption("FilterOptions", "UTF8,CRLF");
            txt.setExportFilter(DocumentFamily.TEXT, "Text (encoded)");
            txt.setExportOption(DocumentFamily.TEXT, "FilterOptions", "UTF8,CRLF");
            addDocumentFormat(txt);

            
            DocumentFormat wikitext = new DocumentFormat("MediaWiki wikitext", "text/x-wiki", "wiki");
            wikitext.setExportFilter(DocumentFamily.TEXT, "MediaWiki");
            addDocumentFormat(wikitext);

            
            DocumentFormat ods = new DocumentFormat("OpenDocument Spreadsheet", DocumentFamily.SPREADSHEET, "application/vnd.oasis.opendocument.spreadsheet", "ods");
            ods.setExportFilter(DocumentFamily.SPREADSHEET, "calc8");
            addDocumentFormat(ods);

            
            DocumentFormat sxc = new DocumentFormat("OpenOffice.org 1.0 Spreadsheet", DocumentFamily.SPREADSHEET, "application/vnd.sun.xml.calc", "sxc");
            sxc.setExportFilter(DocumentFamily.SPREADSHEET, "StarOffice XML (Calc)");
            addDocumentFormat(sxc);

            
            DocumentFormat xls = new DocumentFormat("Microsoft Excel", DocumentFamily.SPREADSHEET, "application/vnd.ms-excel", "xls");
            xls.setExportFilter(DocumentFamily.SPREADSHEET, "MS Excel 97");
            addDocumentFormat(xls);

           
            DocumentFormat xlsx = new DocumentFormat("Microsoft Excel 2007 XML", DocumentFamily.SPREADSHEET, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx");
            addDocumentFormat(xlsx);

            
            DocumentFormat csv = new DocumentFormat("CSV", DocumentFamily.SPREADSHEET, "text/csv", "csv");
            csv.setImportOption("FilterName", "Text - txt - csv (StarCalc)");
            csv.setImportOption("FilterOptions", "44,34,0"); // Field Separator: ','; Text Delimiter: '"'
            csv.setExportFilter(DocumentFamily.SPREADSHEET, "Text - txt - csv (StarCalc)");
            csv.setExportOption(DocumentFamily.SPREADSHEET, "FilterOptions", "44,34,0");
            addDocumentFormat(csv);

            
            DocumentFormat tsv = new DocumentFormat("Tab-separated Values", DocumentFamily.SPREADSHEET, "text/tab-separated-values", "tsv");
            tsv.setImportOption("FilterName", "Text - txt - csv (StarCalc)");
            tsv.setImportOption("FilterOptions", "9,34,0"); // Field Separator: '\t'; Text Delimiter: '"'
            tsv.setExportFilter(DocumentFamily.SPREADSHEET, "Text - txt - csv (StarCalc)");
            tsv.setExportOption(DocumentFamily.SPREADSHEET, "FilterOptions", "9,34,0");
            addDocumentFormat(tsv);

            
            DocumentFormat odp = new DocumentFormat("OpenDocument Presentation", DocumentFamily.PRESENTATION, "application/vnd.oasis.opendocument.presentation", "odp");
            odp.setExportFilter(DocumentFamily.PRESENTATION, "impress8");
            addDocumentFormat(odp);

            
            DocumentFormat sxi = new DocumentFormat("OpenOffice.org 1.0 Presentation", DocumentFamily.PRESENTATION, "application/vnd.sun.xml.impress", "sxi");
            sxi.setExportFilter(DocumentFamily.PRESENTATION, "StarOffice XML (Impress)");
            addDocumentFormat(sxi);

            
            DocumentFormat ppt = new DocumentFormat("Microsoft PowerPoint", DocumentFamily.PRESENTATION, "application/vnd.ms-powerpoint", "ppt");
            ppt.setExportFilter(DocumentFamily.PRESENTATION, "MS PowerPoint 97");
            addDocumentFormat(ppt);

            
            DocumentFormat pptx = new DocumentFormat("Microsoft PowerPoint 2007 XML", DocumentFamily.PRESENTATION, "application/vnd.openxmlformats-officedocument.presentationml.presentation", "pptx");
            addDocumentFormat(pptx);

            
            DocumentFormat odg = new DocumentFormat("OpenDocument Drawing", DocumentFamily.DRAWING, "application/vnd.oasis.opendocument.graphics", "odg");
            odg.setExportFilter(DocumentFamily.DRAWING, "draw8");
            addDocumentFormat(odg);

            
            DocumentFormat svg = new DocumentFormat("Scalable Vector Graphics", "image/svg+xml", "svg");
            svg.setExportFilter(DocumentFamily.DRAWING, "draw_svg_Export");
            addDocumentFormat(svg);
        }
    }
}
