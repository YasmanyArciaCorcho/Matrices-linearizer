using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace Linealizar
{
    public class ReadExecDoc
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        string _inputPath;
        string _directoryOutput;
        string _fulldirectionOutput;
        string _fileName;
        List<string> _sheet;
        public int RowCount { get; set; }
        public int ColCount { get; set; }
        public decimal Percent
        {
            get; private set;
        }

        public ReadExecDoc(string Inpuntpath, string directoryOutput, string fileOutputName, int dimension = 50)
        {
            _inputPath = Inpuntpath;
            _directoryOutput = directoryOutput;

            _fileName = ClearFileName(fileOutputName);
            _fulldirectionOutput = ClearDirectoryPath(_directoryOutput, _fileName);

            RowCount = dimension;
            ColCount = dimension;

            xlApp = new Excel.Application();
            //  string fullPath = _directoryOutput + "\\" + _fileName;
            xlWorkbook = xlApp.Workbooks.Open(_inputPath);

            _sheet = new List<string>(xlWorkbook.Sheets.Count);
            foreach (Excel._Worksheet sheet in xlWorkbook.Sheets)
            {
                _sheet.Add(sheet.Name);
            }

            Percent = RowCount * ColCount * _sheet.Count;
        }

        public IEnumerable<decimal> Numbers()
        {
            #region Creation of table 

            foreach (var percent in CreateDBFile(_fileName, _sheet))
            {
                yield return percent;
            }
            #endregion
        }


        private IEnumerable<decimal> CreateDBFile(string fileName, IEnumerable<string> fieldsName)
        {

            //File output of this proyect finished in ML.
            if (File.Exists(_fulldirectionOutput))
            {
                File.Delete(_fulldirectionOutput);
            }

            OleDbConnection con = new OleDbConnection(GetConnection(_directoryOutput));

            OleDbCommand cmd = new OleDbCommand();

            con.Open();
            decimal count = 0;
            try
            {
                CreateTable(con, cmd, _fileName, fieldsName);
            }
            catch (Exception e)
            {
                con.Close();
                throw new Exception(e.Message, e.InnerException);
            }

            #region Adding values 

            for (int j = 2; j <= RowCount + 1; j++)
            {
                for (int k = 2; k <= ColCount + 1; k++)
                {
                    List<double> toInsert = new List<double>();

                    foreach (var sheet in xlWorkbook.Sheets)
                    {
                        Excel._Worksheet xlWorksheet = (Excel._Worksheet)sheet;
                        Excel.Range xlRange = xlWorksheet.UsedRange;

                        if (xlRange.Cells[j, k] != null && xlRange.Cells[j, k].Value2 != null)
                        {
                            toInsert.Add(double.Parse(xlRange.Cells[j, k].Value2.ToString()));
                        }
                        else
                        {
                            //here print in log that this value was fill.
                            toInsert.Add(0);
                        }
                        yield return count++;
                    }

                    if (toInsert.Count > 0)
                    {
                        cmd.CommandText = InsertElement(_fileName, toInsert); ;
                        try
                        {
                            cmd.ExecuteNonQuery();

                        }
                        catch (Exception e)
                        {
                            con.Close();
                            throw new Exception(e.Message, e.InnerException);
                        }
                    }
                }
                #endregion
            }
            con.Close();
        }

        private void CreateTable(OleDbConnection con, OleDbCommand cmd, string fileName, IEnumerable<string> fieldsName)
        {

            string createSql = "create table " + fileName + " (";

            foreach (var name in fieldsName)
            {
                createSql = createSql + "[" + name + "]" + " " + "Double " + ",";
            }
            createSql = createSql.Substring(0, createSql.Length - 1) + ");";

            cmd.Connection = con;

            cmd.CommandText = createSql;

            cmd.ExecuteNonQuery();
        }

        public string InsertElement(string fileName,
           List<double> elements)
        {
            string insertSql = "insert into " + fileName + " values(";

            foreach (var element in elements)
            {
                insertSql = insertSql + "'" + element + "',";
            }

            insertSql = insertSql.Substring(0, insertSql.Length - 1) + ");";

            return insertSql;

        }

        private static string GetConnection(string path)
        {
            return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=dBASE IV;";
        }

        public static string ReplaceEscape(string str)
        {
            str = str.Replace("'", "''");
            return str;
        }

        public void CloseDocucument()
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background

            foreach (var sheet in xlWorkbook.Sheets)
            {
                Excel._Worksheet xlWorksheet = (Excel._Worksheet)sheet;
                Excel.Range xlRange = xlWorksheet.UsedRange;

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
            }

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        private string ClearDirectoryPath(string directory, string fileName)
        {
            string auxDirectory = _directoryOutput;
            if (auxDirectory.EndsWith(@"\"))
                auxDirectory = auxDirectory.Substring(0, auxDirectory.Length - 1);
            auxDirectory = auxDirectory + "\\" + fileName + ".dbf";
            return auxDirectory;
        }

        private string ClearFileName(string fileName)
        {
            if (fileName.Contains("."))
            {
                int pointPosition = fileName.LastIndexOf('.');
                if (pointPosition >= 0)
                {
                    fileName = fileName.Substring(0, pointPosition);
                }
                fileName = fileName + "ML";
            }
            return fileName;
        }
    }
}
