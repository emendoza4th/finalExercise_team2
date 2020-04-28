using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationTraining_M7.Base_Files
{
    class technologiesHeader
    {
        public static string[] fnTechnologiesHeader()
        {
            string pstrFilePath = @"C:\Users\daniel.luna\Documents\Automation\Exc 2 Mod 8\Autom_TrngM7\AutomTrng_M7\technologies.txt";

            string[] arrLines = System.IO.File.ReadAllLines(pstrFilePath);

            return arrLines;

        }

    }
}
