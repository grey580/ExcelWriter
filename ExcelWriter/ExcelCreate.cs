using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWriter
{
    public class ExcelCreate
    {
        public void Write(DataView dView, string tempPath, string conString)
        {
            // set the temp file
            string vardate = DateTime.Now.ToFileTime().ToString();  
            string fileWpath = tempPath;
            // delete the old file
            if (File.Exists(fileWpath))
            {
                File.Delete(fileWpath);
            }

            // create oledb connection
            OleDbConnection olecon = new OleDbConnection();
            // create command
            OleDbCommand olecmd = new OleDbCommand();
            // creat connection string
            string connstring = conString; //"Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fileWpath + ";" + "Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
            olecon.ConnectionString = connstring;
            olecon.Open();

            olecmd.Connection = olecon;

            // get the datatable
            DataTable dt = dView.Table;

            // Loop through each column to add column header. 
            int colcount = 1;
            string header = "CREATE TABLE [Sheet1] (";
            string header2 = String.Empty;
            foreach (DataColumn col in dt.Columns)
            {
                if (colcount > 1)
                {
                    //sw.Write("\n");
                    header += ", ";
                    header2 += ", ";
                }

                // Output the value of column's header.
                //sw.Write(col.ColumnName.ToString() + "\t");
                header += col.ColumnName.ToString().Replace("#", "").Replace(" ", "") + "1 VARCHAR";
                header2 += col.ColumnName.ToString().Replace("#", "").Replace(" ", "") + "1";

                colcount++;
            }
            header += ");";


            //~~> Command to create the table
            olecmd.CommandText = header;   // "CREATE TABLE Sheet1 (Sno Int, " + "Employee_Name VARCHAR, " + "Company VARCHAR, " + "Date_Of_joining DATE, " + "Stipend DECIMAL, " + "Stocks_Held DECIMAL)";
            olecmd.ExecuteNonQuery();

            int i;
            int j = 0;
            string row = String.Empty;
            foreach (DataRow dr in dt.Rows)
            {
                if (j > 0)
                {
                    //row += ";";
                }

                row += "INSERT INTO [Sheet1] (" + header2 + ") values (";
                for (i = 0; i < dt.Columns.Count; i++)
                {
                    if (i > 0)
                    {
                        row += ",";
                    }
                    //sw.Write(tab + dr[i].ToString());
                    row += "'" + dr[i].ToString().Replace("'", " ").Replace("\"", " ").Replace("#", " ") + "'";
                }
                row += ");";

                //~~> Adding Data
                olecmd.CommandText = row;
                olecmd.ExecuteNonQuery();
                row = String.Empty;
                if (j == 1)
                {
                    //break;
                }
                j++;
            }

            //~~> Adding Data
            //olecmd.CommandText = row;  // "INSERT INTO Sheet1 (Sno, Employee_Name, Company,Date_Of_joining,Stipend,Stocks_Held) values " + 
            //"('1', 'Siddharth Rout', 'Defining Horizons', '20/7/2014','2000.75','0.01')";
            //olecmd.ExecuteNonQuery();

            //~~> Close the connection
            olecon.Close();

            // get a file to stream
            MemoryStream ms = new MemoryStream();
            using (FileStream fs = File.OpenRead(fileWpath))
            {
                fs.CopyTo(ms);
            }

            // delete the old file
            if (File.Exists(fileWpath))
            {
                //File.Delete(fileWpath);
            }

            /*
            // output file to browser
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ContentType = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment; filename=export.xlsx");
            HttpContext.Current.Response.BinaryWrite(ms.ToArray());
            // myMemoryStream.WriteTo(Response.OutputStream); //works too
            HttpContext.Current.Response.Flush();
            HttpContext.Current.Response.Close();
            HttpContext.Current.Response.End();
            */
        }
    }
}
