using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

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
                                    r.Par2 = reader.GetString(2);
                                res.Add(r);
                                // Console.WriteLine("{0} {1} {2}", reader.GetString(0), reader.GetString(1), reader.GetString(2));
                            }
                        }
                    }
                }
            }
            catch (SqlException e)
            {
                Console.WriteLine(e.ToString());
            }
            return res;
        }
    }
}
