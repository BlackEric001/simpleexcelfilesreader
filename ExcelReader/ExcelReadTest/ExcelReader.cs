using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data;
using System.Data.OleDb;

using System.IO;

namespace ExcelReadTest
{
    class ExcelReader
    {
        private string fileName;

        public string excelFileName
        {
            get { return fileName; }
            set { fileName = value; this.fileType = Path.GetExtension(value); }
        }
        public string HDR { get; set; }
        public string fileType { get; set; }

        public writeLog log;

        public ExcelReader()
        {
            this.HDR = "HDR=No;";
        }

        private string getConnectionString()
        {
            ///"HDR=Yes;" indicates that the first row contains columnnames, not data. "HDR=No;" indicates the opposite.
            const string xlsConnStringTemplate = @"Provider=Microsoft.Jet.OLEDB.4.0;
                                      Data Source={0};Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:SFP=False;Extended Properties='Excel 8.0;{1}IMEX=1';";

            const string xlsxConnStringTemplate = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};
                                                Extended Properties='Excel 12.0 Xml;{1}';";

            string cStr = String.Empty;

            switch (this.fileType)
            {
                case ".xls":
                    cStr = String.Format(xlsConnStringTemplate, this.excelFileName, this.HDR);
                    break;
                case ".xlsx":
                    cStr = String.Format(xlsxConnStringTemplate, this.excelFileName, this.HDR);
                    break;
                default:
                    break;
            }

            return cStr;

        }

        public bool getTables(ref DataTable dt)
        {
            OleDbConnection oConn = null;
            try
            {
                String sConnString = getConnectionString();

                oConn = new OleDbConnection(sConnString);
                oConn.Open();

                dt = oConn.GetSchema("Tables");

                log("dt.Columns.Count = " + dt.Columns.Count);
                log("dt.Rows.Count = " + dt.Rows.Count);

                return dt.Rows.Count > 0;

            }
            catch (Exception ex)
            {
                log(ex.Message);
                return false;
            }
            finally
            {
                oConn.Close();
                oConn.Dispose();
            }

        }

        public bool getListData(string listName, ref DataTable dt)
        {
            OleDbConnection oConn = null;
            OleDbCommand oComm = null;
            OleDbDataReader oRdr = null;
            try
            {
                String sConnString = getConnectionString();

                oConn = new OleDbConnection(sConnString);
                oConn.Open();

                String sCommand = @"SELECT * FROM [" + listName + "]";
                oComm = new OleDbCommand(sCommand, oConn);
                oRdr = oComm.ExecuteReader();

                dt.Load(oRdr);

                log("dt.Columns.Count = " + dt.Columns.Count);
                log("dt.Rows.Count = " + dt.Rows.Count);

                return dt.Rows.Count > 0;
            }
            catch (Exception ex)
            {
                log(ex.Message);
                return false;
            }
            finally
            {
                if (oRdr != null) oRdr.Close();
                oRdr = null;
                if (oComm != null) oComm.Dispose();
                oConn.Close();
                oConn.Dispose();
            }
        }


    }
}
