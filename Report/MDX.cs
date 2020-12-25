using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AnalysisServices.AdomdClient;
namespace Report
{
   public class MDX
    {
        void RunMDX(сRequest pRequest, cParameter pParameter)
        {

            AdomdConnection conn = new AdomdConnection(
    "Data Source=localhost;Catalog=YourDatabase");
            conn.Open();

            string commandText = @"SELECT FLATTENED 
    PredictAssociation()
    From
    [Mining Structure Name]
    NATURAL PREDICTION JOIN
    (SELECT (SELECT 1 AS [UserId]) AS [Vm]) AS t ";
            AdomdCommand cmd = new AdomdCommand(commandText, conn);
            AdomdDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                Console.WriteLine(Convert.ToString(dr[0]));
            }
            dr.Close();
            conn.Close();
        }
    }
}
