using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.WebSockets;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataRewriting
{
    class Program
    {
        static void Main(string[] args)
        {
            // https://stackoverflow.com/questions/18993735/how-to-read-single-excel-cell-value?rq=1

            // Ouverture application
            Excel.Application xlapp = new Excel.Application();
            xlapp.Visible = true;
            Excel.Workbook xlWorkBook = xlapp.Workbooks.Open(@"C:\Users\tgaidechevronnay\Documents\DP\Archives Prestations & Productivités.xls");
            Excel.Sheets xlWorkSheet = xlWorkBook.Worksheets;
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)xlWorkSheet.get_Item("Enreg.des prod.");
            
            // Fermeture Application 

            Console.WriteLine(cellValue.ToString());
            Console.ReadKey();
        }

        static void Readrang(Excel.Worksheet ws,int lig)
        {
            int i = 1;
            var cellValue = (ws.Cells[lig, i] as Excel.Range).Value;
            i++;
            if (ws.Cells[lig,col]
        }
    }
}
