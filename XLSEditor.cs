using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace GIBS.Module.Utils
{
    public class XLSEditor
    {

        public ExcelDocumentType DocType { get; set; }

        public string ExcelFileLocation { get; set; }

        public string NewExcelFileLocation { get; set; }

        private List<string> Changes = new List<string>();

        //write string value
        public void SetText(string sheetName, string column, string row, string value)
        {
            if (value != null && value.Contains("'"))
            {
                //this fucks up the SQL command because of the ' 
                value = value.Replace("'", "''");
            }
            string updateString = String.Format("UPDATE [{0}${1}{2}:{1}{2}] SET F1='{3}'", sheetName, column, row, value);
            Changes.Add(updateString);

        }
        //write int value 
        public void SetText(string sheetName, string column, string row, int value)
        {
            string updateString = String.Format("UPDATE [{0}${1}{2}:{1}{2}] SET F1={3}", sheetName, column, row, value.ToString());
           
            Changes.Add(updateString);

        }

        public void SetText(string sheetName, string column, string row, double value)
        {
            string updateString = String.Format("UPDATE [{0}${1}{2}:{1}{2}] SET F1={3}", sheetName, column, row, value.ToString());

            Changes.Add(updateString);

        }

        public void InsertText(string sheetName, string column, string row, string value)
        {
            if (value != null && value.Contains("'"))
            {
                //this fucks up the SQL command because of the ' 
                value = value.Replace("'", "''");
            }
            //"INSERT INTO[sheet1$B2:B2] VALUES ('" + textBox3.Text.ToString() + "')"
            string updateString = String.Format("INSERT INTO [{0}${1}{2}:{1}{2}] VALUES ('{3}')", sheetName, column, (Convert.ToInt32(row) - 1).ToString(), value);
            Changes.Add(updateString);

        }

        public void InsertText(string sheetName, string column, string row, double value)
        {
    
            string updateString = String.Format("INSERT INTO [{0}${1}{2}:{1}{2}] VALUES ({3})", sheetName, column, (Convert.ToInt32(row) - 1).ToString(), value);
            Changes.Add(updateString);

        }

        public void CommitChanges()
        {
            string sConnectionString = "";

            if (DocType == ExcelDocumentType.Excel2007)
            {
                sConnectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{0}';Extended Properties='Excel 12.0 Xml;HDR=NO'", ExcelFileLocation);
            }
            else
            {
                   sConnectionString = String.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source='{0}';Extended Properties='Excel 8.0;HDR=NO'", ExcelFileLocation);
            }


            foreach (string item in Changes)
            {
                using (OleDbConnection objConn = new OleDbConnection(sConnectionString))
                {
                    objConn.Open();
                    using (OleDbCommand objCmdSelect = new OleDbCommand(item, objConn))
                    {
                        objCmdSelect.ExecuteNonQuery();
                    }
                    objConn.Close();
                    objConn.Dispose();
                }
            }
        }
        public void FastCommitChanges()
        {
            string sConnectionString = "";

            if (DocType == ExcelDocumentType.Excel2007)
            {
                sConnectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{0}';Extended Properties='Excel 12.0 Xml;HDR=NO'", ExcelFileLocation);
            }
            else
            {
                sConnectionString = String.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source='{0}';Extended Properties='Excel 8.0;HDR=NO'", ExcelFileLocation);
            }


            
                using (OleDbConnection objConn = new OleDbConnection(sConnectionString))
                {
                    objConn.Open();
                    foreach (string item in Changes)
                    {
                        using (OleDbCommand objCmdSelect = new OleDbCommand(item, objConn))
                        {
                            objCmdSelect.ExecuteNonQuery();
                        }
                    }
                    objConn.Close();
                    objConn.Dispose();
                }
            
        }
    }


    public enum ExcelDocumentType
    {
        Excel2003,
        Excel2007
    };

}
