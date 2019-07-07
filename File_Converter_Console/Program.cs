using System;
using File_Converter;

namespace File_Converter_Console
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel_Transform et = new Excel_Transform();
            et.Workbook_To_Pdf(@"D:\Coffee Code\C#\File_Converter\toSaveAsPdf.xlsx", @"D:\Coffee Code\C#\File_Converter\toSaveAsPdf.pdf");
            
            // NEED TO HAVE MICROSOFT OFFICE INSTALLED
            // Word_Transform wt = new Word_Transform();
            // wt.Doc_To_Pdf(@"D:\Coffee Code\C#\File_Converter\toSaveAsPdf.docx", @"D:\Coffee Code\C#\File_Converter\toSaveAsPdf.pdf");
        }
    }
}
