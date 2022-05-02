using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Excel_CS_Sandbox;
using Excel_CS;


namespace Excel_CS_Sandbox
{
    public partial class Sandbox_Template 
    {
        private Excel_CS.Sandbox register;
        public Sandbox_Template(Excel_CS.Sandbox reg) { this.register = reg; }
    }
}

class Template_Main
{
    public static void Main(String[] args)
    {
        string path = @"C:\Users\Avi\Documents\Visual Studio 2022\Excel_CS\Excel_CS_sandbox\Example.xlsx";

        Excel_CS.Sandbox registerObj = new Excel_CS.Sandbox(path);
        registerObj.readFromExcel();


        var templateObj = new Excel_CS_Sandbox.Sandbox_Template(registerObj);
        String pageContent = templateObj.TransformText();
        System.IO.File.WriteAllText(@"C:\Users\Avi\Documents\Visual Studio 2022\Excel_CS\Excel_CS_sandbox\outputPage.ls", pageContent);

        registerObj.cleanup();
    }
}
    
