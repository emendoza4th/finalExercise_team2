using Excel.Model;
using LinqToExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using System.Data;
using LinqToExcel.Domain;

namespace Excel
{
    public class ExcelReader
    {

        private string filePath = "";
        
        public ExcelQueryFactory connection;
        private string worksheet = "Sheet1";
        private string workbook = "Workbook1";

        public ExcelReader(string _filepath)
        {
            this.filePath = _filepath;

            if (File.Exists(filePath))
            {
                try
                {
                    this.connection = new ExcelQueryFactory(this.filePath);
                    this.connection.DatabaseEngine = DatabaseEngine.Ace;

                }
                catch (Exception ex)
                {

                    throw ex;
                }
            }

            else
            {
                throw new FileNotFoundException("File Not found on the given path");
            }
        }

        public string PathExcelFile
        {
            get
            {
                return this.filePath;
            }
        }
        public ExcelQueryFactory Connection
        {
            get
            {
                return this.connection;
            }
        }

        public void SetWorkSheet(string _sheet)
        {
            this.worksheet = _sheet;
        }

        public void SetWorkBook(string _Workbook)
        {
            this.workbook = _Workbook;
        }

        public List<string> getColumnByName(string _column)
        {
            var columnNames = this.connection.GetColumnNames(_column).ToList();

            return columnNames;
        }
        
        public List<LinkedIn> getLinkedInRows()
        {
            List<LinkedIn> results = new List<LinkedIn>();
            var query2 = from a in this.connection.Worksheet<LinkedIn>(this.worksheet)
                         select a;
            results = query2.ToList();

            return results;
        }

        public List<string> getWorksheetNames()
        {
            var result = this.connection.GetWorksheetNames().ToList();
            return result;
        }
    }
}

