using Microsoft.Office.Interop.Excel;

namespace File_Converter
{
    public class Excel_Transform
    {
        public void Workbook_To_Pdf(string workbookpath, string outputpath)
        {
            // source: https://stackoverflow.com/questions/20874412/create-excel-file-and-save-as-pdf
            Application app = new Application();
            Workbook wkb = app.Workbooks.Open(workbookpath);
            wkb.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputpath);

            wkb.Close();
            app.Quit();
        }
    }
}
