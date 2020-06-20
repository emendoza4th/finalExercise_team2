using LinqToExcel.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Model
{
    public class LinkedIn
    {
        [ExcelColumn("Technology")]
        public string Technology { get; set; }
        [ExcelColumn("Languajes")]
        public string Languajes { get; set; }
        [ExcelColumn("Location")]
        public string Location { get; set; }
        [ExcelColumn("YearExperience")]
        public string YearExperience { get; set; }
        [ExcelColumn("ValidFilters")]
        public string ValidFilters { get; set; }
        [ExcelColumn("NotValidFilter")]
        public string NotValidFilter { get; set; }
        [ExcelColumn("ToolsAndTechnologies")]
        public string ToolsAndTechnologies { get; set; }
        [ExcelColumn("Skills")]
        public string Skills { get; set; }
    }
}
