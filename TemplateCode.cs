using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Excel_CS_Sandbox;


public partial class Sandbox_Template : Excel_CS_Sandbox.Sandbox_Template
{
    public Sandbox_Template()
    {

    }
}



class Template_Main
{
    public static void Main(String[] args)
    {
        Sandbox_Template templateObj = new Sandbox_Template();
        String pageContent = templateObj.TransformText();
        System.IO.File.WriteAllText(@"C:\Users\Avi\Documents\Visual Studio 2022\Excel_CS\Excel_CS_sandbox\outputPage.txt", pageContent);
    }
}
    
