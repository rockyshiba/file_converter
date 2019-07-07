using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace File_Converter
{
    public class Word_Transform
    {
        public void Doc_To_Pdf(string docpath, string outputpath)
        {
            // possible source: https://www.ryadel.com/en/programmatically-convert-ms-word-files-pdf-asp-net-c-doc-docx/

            Application app = new Application();
            Document doc = app.Documents.Open(docpath);
            // doc.SaveAs2(outputpath, WdSaveFormat.wdFormatPDF);
            doc.ExportAsFixedFormat(outputpath, WdExportFormat.wdExportFormatPDF);

            doc.Close();
            app.Quit();
        }
    }
}
