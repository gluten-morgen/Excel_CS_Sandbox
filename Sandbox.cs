using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

class Sandbox
{
    Excel.Application app = new Excel.Application();
    Excel.Workbook wb;
    Excel.Worksheet ws;
    Excel.Range range;
    public Sandbox(string path)
    {
        this.app = new Excel.Application();
        wb = app.Workbooks.Open(path);
        ws = wb.Sheets[1];
        range = ws.UsedRange;
    }

    public void readFromExcel()
    {
        for (int i = 1; i <= range.Rows.Count; i++)
        {
            for (int j = 1; j <= range.Columns.Count; j++)
            {
                Console.Write(range.Cells[i, j].Value);
                Console.Write('\t');
            }
            Console.WriteLine('\n');
        }
    }

    public void cleanup()
    {
        //cleanup
        GC.Collect();
        GC.WaitForPendingFinalizers();

        Marshal.ReleaseComObject(range);
        Marshal.ReleaseComObject(ws);

        wb.Close();
        Marshal.ReleaseComObject(wb);
        app.Quit();
        Marshal.ReleaseComObject(app);
    }



    static void Main(string[] args)
    {
        string path = @"C:\Users\Avi\Documents\Visual Studio 2022\Excel_CS\Example.xlsx";

        Sandbox s = new Sandbox(path);
        s.readFromExcel();
        s.cleanup();
    }
}