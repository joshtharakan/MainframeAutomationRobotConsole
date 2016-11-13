using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using ClosedXML.Excel;

namespace MainframeAutomation
{
    class ImportExcel
    {
        System.Data.OleDb.OleDbConnection OledbConn;
        System.Data.OleDb.OleDbCommand oleExcelCommand;
        System.Data.OleDb.OleDbDataReader dataReader;
        DataTable ContentTable = null;
        string _fileName;
        string _sheetName;
        public ImportExcel(String fileName, String extrn)
        {
            this._fileName = fileName;
            string connString = "";
            if (extrn == ".xls")
                //Connectionstring for excel v8.0    

                connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _fileName + ";Extended Properties=\"Excel 8.0;HdataReader=Yes;IMEX=1\"";
            else
                //Connectionstring fo excel v12.0    
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _fileName + ";Extended Properties=\"Excel 12.0;HdataReader=Yes;IMEX=1\"";

            OledbConn = new System.Data.OleDb.OleDbConnection(connString);
            OledbConn.Close();
        }



        public DataTable ReadFile(String sheetName)
        {
            this._sheetName = sheetName;
            try
            {     
                if (!CheckForSheet(sheetName))
                {
                    throw new EhlApiExceptions("Unable to find Sheet");
                }
                else
                {
                    oleExcelCommand = new System.Data.OleDb.OleDbCommand();
                    oleExcelCommand.Connection = OledbConn;
                    OledbConn.Open();
                    oleExcelCommand.CommandText = "Select * from [" + sheetName + "$]";
                    oleExcelCommand.CommandType = CommandType.Text;
                    dataReader = oleExcelCommand.ExecuteReader();

                    ContentTable = new DataTable();

                    ContentTable.Load(dataReader);

                }
                dataReader.Close();

                OledbConn.Close();
                InsertResultColumn();
                return ContentTable;
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("Unable to connect Excel");
                return null;

            }
        }   
        //{
        //    try
        //    {
        //        String readSheet = null;
        //        DataTable schemaTable = new DataTable();
        //        Boolean sheetFound = false;
        //        oleExcelCommand = new System.Data.OleDb.OleDbCommand();
        //        oleExcelCommand.Connection = OledbConn;
        //        OledbConn.Open();

        //        schemaTable = OledbConn.GetSchema("Tables");
        //        for (int i = 0; i < schemaTable.Rows.Count; i++)
        //        {
        //            readSheet = (string)schemaTable.Rows[i]["TABLE_NAME"];
        //            if (readSheet.Contains(sheetName))
        //            {
        //                sheetFound = true;

        //            }

        //        }

        //        if (!sheetFound)
        //        {
        //            throw new EhlApiExceptions("Unable to find Sheet");
        //        }
        //        else
        //        {
        //            oleExcelCommand.CommandText = "Select * from [" + sheetName + "$]";
        //            oleExcelCommand.CommandType = CommandType.Text;
        //            dataReader = oleExcelCommand.ExecuteReader();

        //            ContentTable = new DataTable();

        //            ContentTable.Load(dataReader);
                   
        //        }
        //        dataReader.Close();

        //        OledbConn.Close();
        //        InsertResultColumn();
        //        return ContentTable;
        //    }
        //    catch (Exception ex)
        //    {
        //        System.Console.WriteLine("Unable to connect Excel");
        //        return null;

        //    }
        //}

        private bool CheckForSheet(string sheetName)
        {
            DataTable schemaTable = new DataTable();
            Boolean sheetFound = false;
            String readSheet = null;
            oleExcelCommand = new System.Data.OleDb.OleDbCommand();
            oleExcelCommand.Connection = OledbConn;
            OledbConn.Open();

            schemaTable = OledbConn.GetSchema("Tables");
            for (int i = 0; i < schemaTable.Rows.Count; i++)
            {
                readSheet = (string)schemaTable.Rows[i]["TABLE_NAME"];
                if (readSheet.Contains(sheetName))
                {
                    sheetFound = true;
                    break;
                }


            }
            OledbConn.Close();
            return sheetFound;
        }


        public void UpdateExcel(DataTable resultData, String resultSheet)
        {
            XLWorkbook WorkBook = new XLWorkbook(_fileName);
            if (CheckForSheet(resultSheet))
            {
                WorkBook.Worksheets.Delete(resultSheet);
            }
            WorkBook.Worksheets.Add(resultData, resultSheet);
            
            WorkBook.Save();
        }

        public void InsertResultColumn()
        {
            ContentTable.Columns.Add("Result", typeof(string));
        }

    }
}
