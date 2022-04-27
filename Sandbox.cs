using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace Excel_CS
{
    public class Parameters
    {
        public int reg_num { get; set; }
        public double reg_numValue { get; set; }

        //public string? Reg_flagValue { get; set; }
        //public int Group { get; set; }

        public List<Parameters> param_list = new();
    }



    public class Sandbox : Parameters
    {
        Excel.Application app = new Excel.Application();
        Excel.Workbook wb;
        Excel.Worksheet ws;
        Excel.Range range;

        //List<Parameters> parameters = new List<Parameters>();
        public Sandbox(string path= @"C:\Users\Avi\Documents\Visual Studio 2022\Excel_CS\Excel_CS_sandbox\Example.xlsx")
        {
            this.app = new Excel.Application();
            wb = app.Workbooks.Open(path);
            ws = wb.Sheets[1];
            range = ws.UsedRange;
        }

        public void readFromExcel()
        {
            for (int i = 2; i <= range.Rows.Count; i++)
            {
                Parameters param = new Parameters();
                for (int j = 1; j <= range.Columns.Count; j++)
                {
                    switch (j)
                    {
                        case 1: param.reg_num = Convert.ToInt32(range.Cells[i, j].Value); break;

                        case 2: param.reg_numValue = Convert.ToDouble(range.Cells[i, j].Value); break;
                    }
                }

                param_list.Add(param);
            }
        }

        public void displayList()
        {
            for (int i = 0; i < param_list.Count; i++)
            {
                Console.Write(param_list[i].reg_num);
                Console.Write('\t');
                Console.WriteLine(param_list[i].reg_numValue);
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
    }
}

/*class Program_Main
{
    static void Main(string[] args)
    {
        string path = @"C:\Users\Avi\Documents\Visual Studio 2022\Excel_CS\Excel_CS_sandbox\Example.xlsx";

        Excel_CS.Sandbox s = new Excel_CS.Sandbox(path);
        s.readFromExcel();
        s.displayList();
        s.cleanup();

    }

}*/