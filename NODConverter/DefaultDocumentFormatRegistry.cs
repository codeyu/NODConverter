namespace NODConverter
{
    public sealed class DefaultDocumentFormatRegistry : BasicDocumentFormatRegistry
    {

        public DefaultDocumentFormatRegistry()
        {
            
            DocumentFormat pdf = new DocumentFormat("Portable Document Format", "application/pdf", "pdf");
            pdf.SetExportFilter(DocumentFamily.Drawing, "draw_pdf_Export");
            pdf.SetExportFilter(DocumentFamily.Presentation, "impress_pdf_Export");
            pdf.SetExportFilter(DocumentFamily.Spreadsheet, "calc_pdf_Export");
            pdf.SetExportFilter(DocumentFamily.Text, "writer_pdf_Export");
            AddDocumentFormat(pdf);

            
            DocumentFormat swf = new DocumentFormat("Macromedia Flash", "application/x-shockwave-flash", "swf");
            swf.SetExportFilter(DocumentFamily.Drawing, "draw_flash_Export");
            swf.SetExportFilter(DocumentFamily.Presentation, "impress_flash_Export");
            AddDocumentFormat(swf);

            
            DocumentFormat xhtml = new DocumentFormat("XHTML", "application/xhtml+xml", "xhtml");
            xhtml.SetExportFilter(DocumentFamily.Presentation, "XHTML Impress File");
            xhtml.SetExportFilter(DocumentFamily.Spreadsheet, "XHTML Calc File");
            xhtml.SetExportFilter(DocumentFamily.Text, "XHTML Writer File");
            AddDocumentFormat(xhtml);

            // HTML is treated as Text when supplied as input, but as an output it is also
            // available for exporting Spreadsheet and Presentation formats
            
            DocumentFormat html = new DocumentFormat("HTML", DocumentFamily.Text, "text/html", "html");
            html.SetExportFilter(DocumentFamily.Presentation, "impress_html_Export");
            html.SetExportFilter(DocumentFamily.Spreadsheet, "HTML (StarCalc)");
            html.SetExportFilter(DocumentFamily.Text, "HTML (StarWriter)");
            AddDocumentFormat(html);

            
            DocumentFormat odt = new DocumentFormat("OpenDocument Text", DocumentFamily.Text, "application/vnd.oasis.opendocument.text", "odt");
            odt.SetExportFilter(DocumentFamily.Text, "writer8");
            AddDocumentFormat(odt);

            
            DocumentFormat sxw = new DocumentFormat("OpenOffice.org 1.0 Text Document", DocumentFamily.Text, "application/vnd.sun.xml.writer", "sxw");
            sxw.SetExportFilter(DocumentFamily.Text, "StarOffice XML (Writer)");
            AddDocumentFormat(sxw);

            
            DocumentFormat doc = new DocumentFormat("Microsoft Word", DocumentFamily.Text, "application/msword", "doc");
            doc.SetExportFilter(DocumentFamily.Text, "MS Word 97");
            AddDocumentFormat(doc);

           
            DocumentFormat docx = new DocumentFormat("Microsoft Word 2007 XML", DocumentFamily.Text, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "docx");
            AddDocumentFormat(docx);

            
            DocumentFormat rtf = new DocumentFormat("Rich Text Format", DocumentFamily.Text, "text/rtf", "rtf");
            rtf.SetExportFilter(DocumentFamily.Text, "Rich Text Format");
            AddDocumentFormat(rtf);

           
            DocumentFormat wpd = new DocumentFormat("WordPerfect", DocumentFamily.Text, "application/wordperfect", "wpd");
            AddDocumentFormat(wpd);

            
            DocumentFormat txt = new DocumentFormat("Plain Text", DocumentFamily.Text, "text/plain", "txt");
            // default to "Text (encoded)" UTF8,CRLF to prevent OOo from trying to display the "ASCII Filter Options" dialog
            txt.SetImportOption("FilterName", "Text (encoded)");
            txt.SetImportOption("FilterOptions", "UTF8,CRLF");
            txt.SetExportFilter(DocumentFamily.Text, "Text (encoded)");
            txt.SetExportOption(DocumentFamily.Text, "FilterOptions", "UTF8,CRLF");
            AddDocumentFormat(txt);

            
            DocumentFormat wikitext = new DocumentFormat("MediaWiki wikitext", "text/x-wiki", "wiki");
            wikitext.SetExportFilter(DocumentFamily.Text, "MediaWiki");
            AddDocumentFormat(wikitext);

            
            DocumentFormat ods = new DocumentFormat("OpenDocument Spreadsheet", DocumentFamily.Spreadsheet, "application/vnd.oasis.opendocument.spreadsheet", "ods");
            ods.SetExportFilter(DocumentFamily.Spreadsheet, "calc8");
            AddDocumentFormat(ods);

            
            DocumentFormat sxc = new DocumentFormat("OpenOffice.org 1.0 Spreadsheet", DocumentFamily.Spreadsheet, "application/vnd.sun.xml.calc", "sxc");
            sxc.SetExportFilter(DocumentFamily.Spreadsheet, "StarOffice XML (Calc)");
            AddDocumentFormat(sxc);

            
            DocumentFormat xls = new DocumentFormat("Microsoft Excel", DocumentFamily.Spreadsheet, "application/vnd.ms-excel", "xls");
            xls.SetExportFilter(DocumentFamily.Spreadsheet, "MS Excel 97");
            AddDocumentFormat(xls);

           
            DocumentFormat xlsx = new DocumentFormat("Microsoft Excel 2007 XML", DocumentFamily.Spreadsheet, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx");
            AddDocumentFormat(xlsx);

            
            DocumentFormat csv = new DocumentFormat("CSV", DocumentFamily.Spreadsheet, "text/csv", "csv");
            csv.SetImportOption("FilterName", "Text - txt - csv (StarCalc)");
            csv.SetImportOption("FilterOptions", "44,34,0"); // Field Separator: ','; Text Delimiter: '"'
            csv.SetExportFilter(DocumentFamily.Spreadsheet, "Text - txt - csv (StarCalc)");
            csv.SetExportOption(DocumentFamily.Spreadsheet, "FilterOptions", "44,34,0");
            AddDocumentFormat(csv);

            
            DocumentFormat tsv = new DocumentFormat("Tab-separated Values", DocumentFamily.Spreadsheet, "text/tab-separated-values", "tsv");
            tsv.SetImportOption("FilterName", "Text - txt - csv (StarCalc)");
            tsv.SetImportOption("FilterOptions", "9,34,0"); // Field Separator: '\t'; Text Delimiter: '"'
            tsv.SetExportFilter(DocumentFamily.Spreadsheet, "Text - txt - csv (StarCalc)");
            tsv.SetExportOption(DocumentFamily.Spreadsheet, "FilterOptions", "9,34,0");
            AddDocumentFormat(tsv);

            
            DocumentFormat odp = new DocumentFormat("OpenDocument Presentation", DocumentFamily.Presentation, "application/vnd.oasis.opendocument.presentation", "odp");
            odp.SetExportFilter(DocumentFamily.Presentation, "impress8");
            AddDocumentFormat(odp);

            
            DocumentFormat sxi = new DocumentFormat("OpenOffice.org 1.0 Presentation", DocumentFamily.Presentation, "application/vnd.sun.xml.impress", "sxi");
            sxi.SetExportFilter(DocumentFamily.Presentation, "StarOffice XML (Impress)");
            AddDocumentFormat(sxi);

            
            DocumentFormat ppt = new DocumentFormat("Microsoft PowerPoint", DocumentFamily.Presentation, "application/vnd.ms-powerpoint", "ppt");
            ppt.SetExportFilter(DocumentFamily.Presentation, "MS PowerPoint 97");
            AddDocumentFormat(ppt);

            
            DocumentFormat pptx = new DocumentFormat("Microsoft PowerPoint 2007 XML", DocumentFamily.Presentation, "application/vnd.openxmlformats-officedocument.presentationml.presentation", "pptx");
            AddDocumentFormat(pptx);

            
            DocumentFormat odg = new DocumentFormat("OpenDocument Drawing", DocumentFamily.Drawing, "application/vnd.oasis.opendocument.graphics", "odg");
            odg.SetExportFilter(DocumentFamily.Drawing, "draw8");
            AddDocumentFormat(odg);

            
            DocumentFormat svg = new DocumentFormat("Scalable Vector Graphics", "image/svg+xml", "svg");
            svg.SetExportFilter(DocumentFamily.Drawing, "draw_svg_Export");
            AddDocumentFormat(svg);
        }
    }
}
