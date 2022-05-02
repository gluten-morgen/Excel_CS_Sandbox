using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace Excel_CS
{
    public class Register
    {
        public string ID { get; set; } = string.Empty;
        public string value { get; set; } = string.Empty;
        public bool isFlag { get; set; }
        public int section { get; set; }

        public List<Register> reg_list = new();
    }



    public class Sandbox : Register
    {
        Excel.Application app = new Excel.Application();
        Excel.Workbook wb;
        Excel.Worksheet ws;
        Excel.Range range;

        public Sandbox(string path = @"C:\Users\Avi\Documents\Visual Studio 2022\Excel_CS\Excel_CS_sandbox\Example.xlsx")
        {
            this.app = new Excel.Application();

            wb = app.Workbooks.Open(path);
            ws = wb.Sheets[1];
            range = ws.UsedRange;
        }

        public void readFromExcel()
        {
            int row_count = range.Rows.Count;
            int col_count = range.Columns.Count;
            string temp;

            for (int i = 2; i <= row_count; i++)
            {
                Register reg = new Register();
                for (int j = 1; j <= col_count; j++)
                {
                    switch (j)
                    {
                        case 1: reg.ID = Convert.ToString(range.Cells[i, j].Value); break;

                        case 2: reg.value = Convert.ToString(range.Cells[i, j].Value); break;

                        case 3:
                            temp = Convert.ToString(range.Cells[i, j].Value);
                            temp = temp.ToLower();
                            if (temp.Equals("true"))
                                {
                                    reg.isFlag = true;
                                }
                            else
                                {
                                    reg.isFlag = false;
                                }
                            break;

                        case 4: 
                            reg.section = Convert.ToInt16(range.Cells[i, j].Value); break;
                    }
                }

                reg_list.Add(reg);
            }
        }

        public void displayList()
        {
            for (int i = 0; i < reg_list.Count; i++)
            {
                Console.Write(reg_list[i].ID);
                Console.Write('\t');
                Console.Write(reg_list[i].value);
                Console.Write('\t');
                Console.Write(reg_list[i].isFlag);
                Console.Write('\t');
                Console.WriteLine(reg_list[i].section);
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

class Program_Main
{
    static void Main(string[] args)
    {
        string path = @"C:\Users\Avi\Documents\Visual Studio 2022\Excel_CS\Excel_CS_sandbox\Example.xlsx";

        Excel_CS.Sandbox s = new Excel_CS.Sandbox(path);
        s.readFromExcel();
        s.displayList();
        s.cleanup();
    }

}