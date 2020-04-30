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
            string pstrFilePath = @"C:\technologies\technologies.txt";

            string[] arrStrLines = System.IO.File.ReadAllLines(pstrFilePath);

            return arrStrLines;

        }

    }
}
