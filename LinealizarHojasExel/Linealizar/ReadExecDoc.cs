using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Remoting.Messaging;


namespace Linealizar
{
    public class ReadExecDoc
    {
        string _inputPath;
        string _directoryOutput;
        string _fulldirectionOutput;
        string _fileName;

        // It seems that so far we have been working with square matrices
        // TODO: We can add support for non square matrices.
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

            RowCount = ColCount = dimension;
        }

        public IEnumerable<decimal> Numbers()
        {
            using (var stream = File.Open(_inputPath, FileMode.Open, FileAccess.Read)) 
            {
                List<string> sheetNames = new List<string>();


                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    Percent = RowCount * RowCount * reader.ResultsCount;

                    do
                    {
                        sheetNames.Add(reader.Name);

                        #region Creation of table 

                        foreach (var percent in CreateDBFile(_fileName, sheetNames, reader))
                        {
                            yield return percent;
                        }
                        #endregion
                    } while (reader.NextResult()); 
                }
            }
        }


        private IEnumerable<decimal> CreateDBFile(string fileName,IEnumerable<string> fieldsName, IExcelDataReader reader)
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

                    foreach (var sheet in _package.Workbook.Worksheets)
                    {
                        if (sheet.Cells[j, k] != null && sheet.Cells[j, k].Text != null)
                        {
                            toInsert.Add(double.Parse(sheet.Cells[j, k].Text.ToString()));
                        }
                        else
                        {
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
            GC.Collect();
            GC.WaitForPendingFinalizers();

            _package.Dispose();
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
