using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Dynamic;
using System.Linq;
using System.Net.WebSockets;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using ModelProductivityDB;

namespace DataRewriting
{
    class Program
    {
        static void Main(string[] args)
        {
            // https://stackoverflow.com/questions/18993735/how-to-read-single-excel-cell-value?rq=1

            // Opening app
            Excel.Application xlapp = new Excel.Application();
            xlapp.Visible = true;
            Excel.Workbook xlWorkBook =
                xlapp.Workbooks.Open(
                    @"C:\Users\tgaidechevronnay\Documents\DP\Archives Prestations & Productivités.xls");
            Excel.Sheets xlWorkSheet = xlWorkBook.Worksheets;
            Excel.Worksheet excelWorksheet = (Excel.Worksheet) xlWorkSheet.get_Item("Enreg.des prod.");

            ReadrangAndWriteDB(excelWorksheet, 2);

            // closing app 
            xlapp.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp);
            Console.WriteLine("Press a key to quit");
            Console.ReadKey();
        }

        static void ReadrangAndWriteDB(Excel.Worksheet ws, int lig)
        {
            int i = 1;
            int ValidScrapWIPExtrusion = 0;
            Collection<Machine> Machines = new Collection<Machine>();

            Collection<Production> productionstmp = new Collection<Production>();            // creating  production
            while((ws.Cells[lig, 1] as Excel.Range).Value != null || (ws.Cells[lig+1, 1] as Excel.Range).Value != null )         
                //Reading data range by range, checking if there is a value Or 
            {

                bool MachineFounded = false;
                // Creating machine used
                Machine machinetmp = new Machine()
                {
                    Name = (string) (ws.Cells[lig, 1] as Excel.Range).Value
                };
                // Adding the machine

                var machine = Machines
                    .Where(m => m.Name == machinetmp.Name)
                    .FirstOrDefault<Machine>();
                if (machine != null)
                {
                    Machines.Add(machine);
                }

                //foreach (Machine machine in Machines)
                //{
                //    // If we have a machine in the park, it's found
                //    if (machine.Name == machinetmp.Name)
                //    {
                //        MachineFounded = true;
                //        break;
                //    }
                //}
                if (!MachineFounded)
                    Machines.Add(machinetmp);
                    
                if (!Int32.TryParse((ws.Cells[lig, 7] as Excel.Range).Text, out ValidScrapWIPExtrusion))
                {
                    // If the value is an INT32, it's ok => use ValidScrapWIPExtrusion
                    // Else, it's not an INT32, NOK => set ValidScrapWIPExtrusion to 0

                    ValidScrapWIPExtrusion = 0;
                }

                productionstmp.Add(new Production()
                {           
                    
                    Machine = machinetmp, // adding the machine 
                    WorkingDate = (DateTime?) (ws.Cells[lig, 2] as Excel.Range).Value,
                    HourNeeded = (float?) (ws.Cells[lig, 4] as Excel.Range).Value,
                    HourReallyDone = (float?) (ws.Cells[lig, 5] as Excel.Range).Value,
                    QuantityProduced = (int?) (ws.Cells[lig, 6] as Excel.Range).Value,
                    ScrapWIPExtrusion = ValidScrapWIPExtrusion,
                    ScrapPFExtrusion = (int?) (ws.Cells[lig, 8] as Excel.Range).Value,
                    Comments = (string) (ws.Cells[lig, 9] as Excel.Range).Text.ToString()
                    //Comments = (string) (ws.Cells[lig, 9] as Excel.Range).Value
                });
                if (lig == 50)
                    break;
                lig++;
                Console.WriteLine(lig);
            }

            using (var ctx = new ProductivityDBContext())
            {
                // Adding machines 

                Console.WriteLine(Machines.Count.ToString());
                //ctx.Productions.AddRange(productionstmp);                
                //ctx.SaveChanges();
                Console.WriteLine("Written in the DB");
            };
        }

    }
}
