using System;
using System.Data;
using System.Threading;
using Microsoft.AnalysisServices;
using Microsoft.AnalysisServices.Xmla;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Globalization;
using System.Data.OleDb;
using System.Xml;
namespace Test
{

    class Program
	{
		
		public static void Main(string[] args)
		{
			
			StringBuilder builder = new StringBuilder();
			string varView;
			DataSet dataSet = new DataSet();
			DataTable TTable = dataSet.Tables.Add("table");
			using (StreamReader sr = new StreamReader(@"D:\WORK\DW_OLAP20130212.xmla")) {
    		String line;
		    while ( (line = sr.ReadLine() ) != null) {
        		builder.AppendLine(line);
    		}
    sr.Close();
}

			String  str=builder.ToString().ToUpper();
			Console.WriteLine("XMLA");
			
		 OleDbConnection myOleDbConnection = new OleDbConnection("Provider=OraOLEDB.Oracle.1;Data Source=Sprut3;Persist Security Info=True;user id=vimas;password=1");
         OleDbDataAdapter adapterTable =
                        new OleDbDataAdapter("select o.OBJECT_NAME from sys.dba_objects o where o.OWNER='DW' and o.status='INVALID' and  o.OBJECT_TYPE='VIEW'", myOleDbConnection);
                    adapterTable.Fill(TTable);
                    foreach (DataRow row in TTable.Rows)
                    {
                    	varView = row[0].ToString() ;
                    	if ( str.IndexOf(varView)<=0)
//                    	 Console.WriteLine( varView);	
                  	//else
                    	 Console.WriteLine("Drop view DW."+ varView+ ";");
                    }
                    TTable.Clear();
                    myOleDbConnection.Close();
			


                    
			
                                 
 //Console.WriteLine(tt);
 Console.WriteLine("END");
 Console.ReadKey();
/*	DataTable table = new DataTable();
	dt.Columns.Add("Dosage", typeof(int));
	dt.Columns.Add("Drug", typeof(string));
	dt.Columns.Add("Patient", typeof(string));
	dt.Columns.Add("Date", typeof(DateTime));
	
*/	
/* 
DataRow row1;
row1 = new DataRow();
 
XmlDocument doc = new XmlDocument();
 doc.Load("D:\\INSTALL\\GI8120\\Каналы_Spark\\tv_prog.xml");
 XmlNodeList NL;
 NL=doc.DocumentElement.SelectNodes("prog");
 foreach (XmlNode n in NL)
 {
 	row1["sat_key"]=n.Attributes["sat_key"].Value.Trim();
	row1["tp_key"]=n.Attributes["tp_key"].Value.Trim();
	row1["service_key"]=n.Attributes["service_key"].Value.Trim();
	row1["tuner_type_idex"]=n.Attributes["tuner_type_idex"].Value.Trim();
 }
 Console.WriteLine(NL.Count);

 

sat_key
tp_key
service_key
tuner_type_idex
sat_type
service_id
pmt_pid
video_pid
audio_pid
pcr_pid
block
bskip
audio_lang
audio_type
video_type
provider
encrypt
hd
name
	
 
*/ 
/*		}
		public string GetVar(string parKey1,string parKey2="")
		{
			try
			{
				if(parKey2.Length==0)
				  return doc.DocumentElement.SelectSingleNode(parKey1).InnerText.Trim() ;
				else
			      return doc.DocumentElement.SelectSingleNode(parKey1).SelectSingleNode(parKey2).InnerText.Trim() ;
			}
			catch ( Exception ex)
			{
			 return null;	
			}
			
		}
		public string GetAttribute(string parAttribute, string parKey1, string parKey2="" )
		{
			try
			{
				if(parKey2.Length==0)
					return doc.DocumentElement.SelectSingleNode(parKey1).Attributes[parAttribute].Value.Trim() ;
				else
			      return doc.DocumentElement.SelectSingleNode(parKey1).SelectSingleNode(parKey2).Attributes[parAttribute].Value.Trim() ;
			}
			catch ( Exception ex)
			{
			 return null;	
			}
			
		}
		
	} 

return ;			
			
		DateTime varDate=	DateTime.Parse("31.01.2012");
		return;
	string varString;
	Server s = new Server();
		 s.Connect(@"Data Source=srv-bat;Provider=msolap;");
         Database varDB = s.Databases.FindByName("DW_OLAP");
         Microsoft.AnalysisServices.Cube varCube = varDB.Cubes.FindByName("ОПТІМ");
         varString = varCube.MeasureGroups["FACT OPTIM TEST"].Name;
         foreach (MeasureGroup g in varCube.MeasureGroups)
         {
         	varString = g.Name;
         }
         


			string stt=strings.GetWordNum("ggghh;sdrgsd;dsfgsdf;jkkj",7,";") ;
			 stt=strings.GetWordNum("ggghh;sdrgsd;dsfgsdf;jkkj",3,";") ;
			 stt=strings.GetWordNum("ggghh;sdrgsd;dsfgsdf;jkkj",4,";") ;
			 stt=strings.GetWordNum("ggghh;sdrgsd;dsfgsdf;jkkj",5,";") ;
			stt=strings.GetWordNum("ggghh;sdrgsd;dsfgsdf;dsfgdrt;jkkj",1,";") ;
			stt=strings.GetWordNum("ggghh;sdrgsd;dsfgsdf;dsfgdrt;jkkj",1,",") ;
			return;
			try
			{
			XmlDocument doc = new XmlDocument();
			doc.Load("d:\\config.xml");
			string st= doc.DocumentElement.SelectSingleNode("Step1").SelectSingleNode("prepare2").InnerText.Trim() ;
			}
			
			catch ( Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			
			
    //XmlTextReader reader = new XmlTextReader("config.xml");
    //reader.WhitespaceHandling = WhitespaceHandling.None;			
			
			
			
			Console.WriteLine("Hello World!");
			string varMDX=@"SELECT
	NON EMPTY {[Measures].[SUMMA]} 
 ON COLUMNS
FROM [1С]
WHERE ([Компанії_1С].[Компанія].&[683073175],[Час].[Рік_Місяць].[Місяць].&[201108],
[ACC CRE].[Рахунок].&[1238629343]},
{[ACC DEB].[Рахунок].&[-143771587]},{[Статті Бютжету_1С].[Стаття].&[233030047]} )";

			Cube cube =new Cube ("Data Source=SRV-BAT;Provider=msolap;Initial Catalog=DW_OLAP");

			double aa=cube.RunMDX(varMDX);
			return;



			
//			FastZip fZip = new FastZip();
//			fZip.CreateZip(@"d:\\temp\\new.zip", @"d:\\temp\\",  false, "11.txt");
			
			MyMail.SendMailZIP("gelo@pakko.ua","d:\\11.reg" ,"Щоденна розсилання");
			
return;				
				
			
			
			Console.WriteLine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location));
			   Console.ReadKey(true);
return;			   
			// TODO: Implement Functionality Here
			
			string varRez;
			string varXML=@"
<Batch xmlns='http://schemas.microsoft.com/analysisservices/2003/engine'>
<Parallel>
   <Process>
      <Object>
        <DatabaseID>DW_OLAP</DatabaseID>
        <CubeID>Mer 1</CubeID>
        <MeasureGroupID>FACT MINI  INVENTORY PLAN WEEK</MeasureGroupID>
        <PartitionID>FACT MINI INVENTORY PLAN WEEK 20100528</PartitionID>
      </Object>
      <Type>ProcessFull</Type>
    </Process>
</Parallel>
</Batch>
";
			
            XmlaClient clnt = new XmlaClient();
            clnt.Connect(@"Data Source=SRV-OLAP;Provider=msolap;Initial Catalog=DW_OLAP");

            

            clnt.Execute(varXML,"",out varRez,false,false );
            int i= varRez.IndexOf("<Error");
          	
			
			Console.Write(varRez);
			Console.ReadKey(true);*/
		}
	}
}