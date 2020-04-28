using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutomationTraining_M7.Reporting;
using AutomationTraining_M7.Data_Model;
using Syncfusion.XlsIO;
using System.Collections;

namespace AutomationTraining_M7.Base_Files
{
    class ExportDataCsv:clsReportManager
    {
        public string FileName { get; set; }
        public string Header { get; set; }
        public List<Candidates> Member { get; set; }

        public ExportDataCsv(string pstrTechnologies)
        {
            FileName = $"{pstrTechnologies}_{DateTime.Now.ToString("MMddyyyy_HHmmss")}.xlsx";
           // Header = "Actor Name|Profile/Role|LinkedIn URL|Last Job|Experience|Skills and Validations|Tools and technologies";
            Member = new List<Candidates>();
        }
        
        public void fnCreateFile(List<Candidates> plstCan)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //Create a new workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                int i = 2;

                sheet["A1"].Text = "Actor Name";
                sheet["A1"].CellStyle.Font.Bold = true;
                sheet["A1"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                sheet["B1"].Text = "Profile Role";
                sheet["B1"].CellStyle.Font.Bold = true;
                sheet["B1"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                sheet["C1"].Text = "LinkedIn Url";
                sheet["C1"].CellStyle.Font.Bold = true;
                sheet["C1"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                sheet["D1"].Text = "Last Job";
                sheet["D1"].CellStyle.Font.Bold = true;
                sheet["D1"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                sheet["E1"].Text = "Experience";
                sheet["E1"].CellStyle.Font.Bold = true;
                sheet["E1}"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                sheet["F1"].Text = "Skill Validations";
                sheet["F1"].CellStyle.Font.Bold = true;
                sheet["F1"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                
                sheet["G1"].Text = "Tools Technologies";
                sheet["G1"].CellStyle.Font.Bold = true;
                sheet["G1"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                foreach (Candidates candidate in plstCan)
                {
                    
                    sheet[$"A{i}"].Text = candidate.ActorName;
                    sheet[$"A{i}"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                    sheet[$"B{i}"].Text = candidate.ProfileRole;
                    sheet[$"B{i}"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                    sheet[$"C{i}"].Text = candidate.LinkedInUrl;
                    sheet[$"C{i}"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                    sheet[$"D{i}"].Text = candidate.LastJob;
                    sheet[$"D{i}"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                    sheet[$"E{i}"].Text = candidate.Experience;
                    sheet[$"E{i}"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                    sheet[$"F{i}"].Text = candidate.SkillsValidations;
                    sheet[$"F{i}"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                    sheet[$"G{i}"].Text = candidate.ToolsTechnologies;
                    sheet[$"G{i}"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                    sheet.UsedRange.AutofitColumns();

                    i += 1;
                }

                Stream excelStream = File.Create(Path.GetFullPath(fnGetCSVPath() + FileName));
                workbook.SaveAs(excelStream);
                excelStream.Dispose();
            }
        }
    }
}