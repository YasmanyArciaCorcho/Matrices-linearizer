using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Remoting.Messaging;


namespace Linealizar
{
    public class ReadExecDoc
    {
        private string _inputPath;
        private string _directoryOutput;
        private string _fulldirectionOutput;
        private string _fileName;

        private List<string> _sheetNames;
        private const int _maxSheetsToProcess = 10;
        private int _totalSheetsProcessed = 0;

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

        public void Numbers()
        {
            if (File.Exists(_fulldirectionOutput))
            {
                File.Delete(_fulldirectionOutput);
            }

            using (var stream = File.Open(_inputPath, FileMode.Open, FileAccess.Read))
            {

                OleDbCommand cmd = new OleDbCommand();

                cmd.Connection = new OleDbConnection(GetConnection(_directoryOutput));

                cmd.Connection.Open();

                _sheetNames = new List<string>();
                // Create the db shema
                try
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        try
                        {
                            string createSql = "create table " + _fulldirectionOutput + " (";
                            do
                            {
                                createSql = createSql + "[" + reader.Name + "]" + " " + "Double " + ",";

                                _sheetNames.Add(reader.Name);
                            }
                            while (reader.NextResult());

                            createSql = createSql.Substring(0, createSql.Length - 1) + ");";

                            cmd.CommandText = createSql;
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception)
                        {
                            // Add error logging
                            // Error when the matrices are being linearized
                        }

                    }
                }
                catch (Exception)
                {
                    // Add error logging
                    // Error reading the Excel file
                }

                try
                {

                    // Insert data on each table
                    // A sheet turns into a column as how the linearizing process defines
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        Percent = RowCount * RowCount * reader.ResultsCount;
                        List<List<double>> rowsSheet = new List<List<double>>();

                        do
                        {
                            List<double> row = new List<double>();
                            while (reader.Read())
                            {
                                for (int col = 0; col < reader.FieldCount; col++)
                                {
                                    object cellValue = reader.GetValue(col);
                                    if (cellValue is double)
                                    {
                                        row.Add((double)cellValue);
                                    }
                                    else
                                    {
                                        row.Add(0);
                                    }
                                }
                            };

                            rowsSheet.Add(row);

                            if (rowsSheet.Count >= _maxSheetsToProcess)
                            {
                                cmd.CommandText = InsertSheetsQuery(rowsSheet);
                                cmd.ExecuteNonQuery();

                                _totalSheetsProcessed += rowsSheet.Count;
                                
                                rowsSheet.Clear();
                            }

                        } while (reader.NextResult());
                    }

                }
                catch (Exception)
                {
                    // Add error logging
                    // Error reading the Excel file
                }
                cmd.Connection.Close();
            }
        }

        public string InsertSheetsQuery(List<List<double>> SheetsValues)
        {
            string insertSql = "";
            foreach (var sheetValue in SheetsValues)
            {
                insertSql += "insert into " + _sheetNames[_totalSheetsProcessed] + " values(";

                foreach (var cellValue in sheetValue)
                {
                    insertSql = insertSql + "'" + cellValue + "',";
                }

                insertSql = insertSql.Substring(0, insertSql.Length - 1) + ");";

                _totalSheetsProcessed++;
            }

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
