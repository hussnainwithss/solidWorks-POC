// using System;
// using System.IO;
// using System.Runtime.InteropServices;
// using Xarial.XCad.Documents;
// using Xarial.XCad.Documents.Enums;
// using Xarial.XCad.Documents.Structures;
// using Xarial.XCad.Enums;
// using Xarial.XCad.SolidWorks;
// using Xarial.XCad.SolidWorks.Documents;
// using Xarial.XCad.SolidWorks.Enums;
// using SolidWorks.Interop.sldworks;

// namespace keychain
// {
//     class Program
//     { static void Main(string[] args)
//         {
//             using(var app = SwApplicationFactory.Create(SwVersion_e.Sw2022, ApplicationState_e.Silent))
//             {
//                 ModelDoc2 swModel = default(ModelDoc2);
//                 DesignTable swDesTable = default(DesignTable);
//                 bool bRet = false;

//                 swDesTable = (DesignTable)swModel.GetDesignTable(); 
//                 swDesTable = (DesignTable)swModel.GetDesignTable(); 
//                 bRet = swDesTable.Attach();
//                 Console.WriteLine("Opened %d", bRet);
//             }
//         } }
// }

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;
using System.Diagnostics;
 
namespace GetTitleDesignTableCSharp.csproj
{
 
    partial class SolidWorksMacro
    {
 
        public void Main()
        {
            ModelDoc2 swModel = default(ModelDoc2);
            DesignTable swDesTable = default(DesignTable);
            int nTotRow = 0;
            int nTotCol = 0;
            string sRowStr = null;
            int i = 0;
            int j = 0;
            bool bRet = false;
 
            swModel = (ModelDoc2)swApp.ActiveDoc;
 
            swDesTable = (DesignTable)swModel.GetDesignTable();
            bRet = swDesTable.Attach();
 
            nTotRow = swDesTable.GetTotalRowCount();
            nTotCol = swDesTable.GetTotalColumnCount();
            Debug.Print("File = " + swModel.GetPathName());
            Debug.Print("  Title        = " + swDesTable.GetTitle());
            Debug.Print("  Row          = " + swDesTable.GetRowCount());
            Debug.Print("  Col          = " + swDesTable.GetColumnCount());
            Debug.Print("  TotRow       = " + nTotRow);
            Debug.Print("  TotCol       = " + nTotCol);
            Debug.Print("  VisRow       = " + swDesTable.GetVisibleRowCount());
            Debug.Print("  VisCol       = " + swDesTable.GetVisibleColumnCount());
            Debug.Print("");
 
            for (i = 0; i <= nTotRow; i++)
            {
                sRowStr = "  |";
                for (j = 0; j <= nTotCol; j++)
                {
                    sRowStr = sRowStr + swDesTable.GetEntryText(i, j) + "|";
                }
                Debug.Print(sRowStr);
            }
            swDesTable.Detach();
        }
 
        /// <summary>
        /// The SldWorks swApp variable is pre-assigned for you.
        /// </summary>
 
        public SldWorks swApp;
 
    }
}

