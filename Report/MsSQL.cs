using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using ExcelApp = Microsoft.Office.Interop.Excel;
using System.Data;
using Utils;

namespace Report
{
    public class MsSQL
    {
        public IEnumerable<cParameter> RunMsSQL(сRequest pSQL)
        {
            List<cParameter> res = new List<cParameter>();
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "sqlsrv2.vopak.local";
                builder.UserID = "dwreader";
                builder.Password = "DW_Reader";
                builder.InitialCatalog = "for_cubes";
                builder.ConnectTimeout = 600;
                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();
                    // String sql = "SELECT name, collation_name FROM sys.databases";
                    using (SqlCommand command = new SqlCommand(pSQL.Request, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                cParameter r = new cParameter() { Par1 = reader.GetString(0), Name = reader.GetString(1), EMail = reader.GetString(2) };
                                if (reader.FieldCount > 3)
                                    r.Par2 = reader.GetString(3);
                                res.Add(r);
                                // Console.WriteLine("{0} {1} {2}", reader.GetString(0), reader.GetString(1), reader.GetString(2));
                            }
                        }
                    }
                }
            }
            catch (SqlException e)
            {
                FileLogger.WriteLogMessage($"MsSQL.Run Error=>  {e}");                
            }
            return res;
        }

        public IEnumerable<cParameter> Run(сRequest pSQL)
        {
            List<cParameter> res = new List<cParameter>();
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "sqlsrv2.vopak.local";
                builder.UserID = "dwreader";
                builder.Password = "DW_Reader";
                builder.InitialCatalog = "for_cubes";
                builder.ConnectTimeout = 300;
                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();
                    // String sql = "SELECT name, collation_name FROM sys.databases";
                    using (SqlCommand command = new SqlCommand(pSQL.GetRequest, connection))
                    {
                        command.CommandTimeout = 300;
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            FileLogger.WriteLogMessage($"MsSQL.Run({pSQL.GetRequest},{pSQL.Row},{pSQL.Column}) ");
                            DataTable dt = new DataTable();
                            dt.Load(reader);

                            int i = 0;
                            if (pSQL.IsHead)
                            {
                                foreach (DataColumn c in dt.Columns)
                                    pSQL.Sheet.Cells[pSQL.Row, pSQL.Column + i++].value = c.ColumnName;
                                //pSQL.Row++;
                            }
                            object[,] arr = new object[dt.Rows.Count, dt.Columns.Count];
                            for (int r = 0; r < dt.Rows.Count; r++)
                            {
                                DataRow dr = dt.Rows[r];
                                for (int c = 0; c < dt.Columns.Count; c++)
                                {
                                    arr[r, c] = dr[c];
                                }
                            }
                            ExcelApp.Range c1 = (ExcelApp.Range)pSQL.Sheet.Cells[pSQL.Row+ (pSQL.IsHead?1:0), pSQL.Column];
                            ExcelApp.Range c2 = (ExcelApp.Range)pSQL.Sheet.Cells[pSQL.Row+ (pSQL.IsHead ? 1 : 0) + dt.Rows.Count - 1, dt.Columns.Count+ pSQL.Column-1];
                            ExcelApp.Range range = pSQL.Sheet.get_Range(c1, c2);
                            range.Value = arr;
                        }
                    }
                }
            }
            catch (SqlException e)
            {
                FileLogger.WriteLogMessage($"MsSQL.Run Error=>  {e}");
                //Console.WriteLine(e.ToString());
            }
            return res;
        }
    }
}
