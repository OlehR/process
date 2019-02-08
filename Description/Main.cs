using System;
using Microsoft.AnalysisServices;
using Microsoft.AnalysisServices.AdomdClient;

namespace TestConsole
{
	class Program
	{
		public static void Main(string[] args)
		{
			string varSRV="SRV-OLAP";
			string varDB="dw_olap";
			string varCube="Рух товарів";
			for (int i=0;i<args.Length;i++)
        	{
        		if (args[i].ToUpper().StartsWith("/SERVER:"))
        		    varSRV=args[i].Substring(8);
        		else if (args[i].ToUpper().StartsWith("/DB:"))
        		    varDB=args[i].Substring(4);         
        		else if (args[i].ToUpper().StartsWith("/CUBE:"))
        		    varCube=args[i].Substring(6);         
			}
			
			string varSeparator = "\t";
			Proc p=new Proc() ;
			p.CreateLog("d:\\Cube_"+varCube+".txt");
    		Microsoft.AnalysisServices.Server s = new Microsoft.AnalysisServices.Server();
            s.Connect(@"Data Source="+varSRV+";Provider=msolap;Initial Catalog="+varDB);
            Microsoft.AnalysisServices.Database d = s.Databases.FindByName(varDB);
            Microsoft.AnalysisServices.Cube c = d.Cubes.FindByName(varCube);
            p.Log("База:"+varDB+" Куб:"+varCube);
            p.Log("Розмірності");
            foreach (Microsoft.AnalysisServices.CubeDimension dim in c.Dimensions)
            { 
            	p.Log(dim.Name + varSeparator + dim.DimensionID+varSeparator+dim.Description );
            	foreach (Microsoft.AnalysisServices.CubeAttribute attr in dim.Attributes )
            		p.Log(varSeparator+attr.Attribute+varSeparator+ attr.AttributeID+"\t" +attr.Attribute.Description);            		
            }
            p.Log("Групи мір");
            foreach (Microsoft.AnalysisServices.MeasureGroup mg in c.MeasureGroups)
            {
            	p.Log(mg.Name+varSeparator+mg.ID +varSeparator+mg.Description);
            	foreach (Microsoft.AnalysisServices.Measure m in mg.Measures )
            		p.Log(varSeparator+m.ID+varSeparator+m.Name +varSeparator+ m.Description );
                    
            }
          
            p.Log("Калькульовані міри");
         
		Microsoft.AnalysisServices.AdomdClient.AdomdConnection cn = new Microsoft.AnalysisServices.AdomdClient.AdomdConnection("Data Source="+varSRV+";Provider=msolap;Initial Catalog="+varDB);
            cn.Open();

            /*foreach ( Microsoft.AnalysisServices.AdomdClient.CubeDef tt in cn.Cubes)
            	p.Log(tt.Name+varSeparator+tt.Caption );*/
            try 
            {   
            foreach (Microsoft.AnalysisServices.AdomdClient.Measure m in cn.Cubes[varCube].Measures)
              if  (  string.IsNullOrEmpty(m.Expression)==false  )  
            		p.Log(m.UniqueName+varSeparator ); //+m.Expression +varSeparator+ m.Description );
        
            //Console.WriteLine("{0}: {1}",m.UniqueName,m.Expression );

		    	
            }	
            catch {            }
            finally {              cn.Close();         };
            //Console.ReadLine();
            
			
			p.CloseLog();
		}
	}
}