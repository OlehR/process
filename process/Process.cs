using System;
using System.Linq;
using System.Data;

using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Globalization;
using System.Data.OleDb;
using System.Xml;
//using System.Xml.XPath;
//using System.Threading.Tasks;

namespace Process
{
	 
    class Program
    {
        public static void Main(string[] args)
        {
        	string varFileXML=Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)+"\\process.xml";
        	string varKey = @"
Доступнi ключi: /XML:process.xml /Server:localhost /DB:dw_olap /Cube: /Step:0 /PARALLEL:8 /Metod:0
/XMLA: 1 - процесити через XMLA, 0 - не використовувати XMLA для не INDEX процесінга
/Metod:0 - ігнорувати помилки ключі з конвертацією, 1 - пробувувати без ігнорування
/DAY: - за скільки днів перепроцешувати куби  за замовчуванням 20
/STATE (:2) - Показує стани кубів. :2-Розширена інформація
/DATESTART:DD.MM.YYYY - З Якої дати процесити партіциї.
/PROCESSDIMENSION: - (NONE,UPDATE,FULL)  По замовчуванню UPDATE
/PROCESSCUBE: - (NONE,DATA,FULL) По замовчуванню FULL
/ARX:01.01.2012 - Процесити в режимі архів. (З якої дати розширювати партиції) дата не обов'язковий параметр.
/DAYPROCESS:20 - за скільки днів перепроцешувати партіциї (Переважає значення з налаштувань Виміру)

";
       	
			//Перевіряємо чи є в параметрах XML файл
 			for (int i=0;i<args.Length;i++)
   				if (args[i].ToUpper().StartsWith("/XML:"))
        		    varFileXML=args[i].Substring(5);
 				else if (args[i].ToUpper().StartsWith("/STEP:"))
        			GlobalVar.varStep=Convert.ToInt32( args[i].Substring(6));
 			
 			//Перевіряємо наявність XML файла
 			if(File.Exists(varFileXML))
 			{
 				MyXML myXML =new MyXML(varFileXML);
 				GlobalVar.varServer= ( myXML.GetVar("Server") == null? GlobalVar.varServer : myXML.GetVar("Server") );
 				GlobalVar.varDB = ( myXML.GetVar("Database") == null? GlobalVar.varDB : myXML.GetVar("Database") );
 				//GlobalVar.varStep = ( myXML.GetVar("DefaultStep") == null? GlobalVar.varStep : Convert.ToInt32( myXML.GetVar("DefaultStep")) );
 				GlobalVar.varMaxParallel = (myXML.GetAttribute("maxParallel","XMLA")==null ?  GlobalVar.varMaxParallel : Convert.ToInt32(myXML.GetAttribute("maxParallel","XMLA") )) ;
 				GlobalVar.varConectSQL = myXML.GetVar("ConectSQL");
 				GlobalVar.varKeyErrorLogFile = (myXML.GetVar("KeyErrorLogPath")==null ? GlobalVar.varKeyErrorLogFile :myXML.GetVar("KeyErrorLogPath")+"\\Error_"+DateTime.Now.ToString("yyyyMMdd")+".log");
 				if(myXML.GetVar("Step" +GlobalVar.varStep.ToString().Trim(),"ProcessDimension")!=null )
 					MyXMLA.SetProcessTypeDimension( myXML.GetVar("Step" +GlobalVar.varStep.ToString().Trim(),"ProcessDimension"));
 				GlobalVar.varPrepareSQL  = myXML.GetVar("Step" +GlobalVar.varStep.ToString().Trim(),"PrepareSQL");
 				GlobalVar.varWaitSQL     = myXML.GetVar("Step" +GlobalVar.varStep.ToString().Trim(),"WaitSQL");
 				GlobalVar.varTimeStart = ( myXML.GetAttribute("Start","Step" +GlobalVar.varStep.ToString().Trim(),"Time")==null ? GlobalVar.varTimeStart : Convert.ToInt32(myXML.GetAttribute("Start","Step" +GlobalVar.varStep.ToString().Trim(),"Time") ));
 				GlobalVar.varTimeEnd = ( myXML.GetAttribute("End","Step" +GlobalVar.varStep.ToString().Trim(),"Time")==null ? GlobalVar.varTimeEnd : Convert.ToInt32(myXML.GetAttribute("End","Step" +GlobalVar.varStep.ToString().Trim(),"Time") ));
 				                                           
 			}
// 			string var= GlobalVar.varProcessDimension;
 			
            //Параметри з командного рядка мають перевагу.
        	for (int i=0;i<args.Length;i++)
        	{
                if (args[i].ToUpper().StartsWith("/SERVER:"))
                    GlobalVar.varServer = args[i].Substring(8);
                else if (args[i].ToUpper().StartsWith("/DB:"))
                    GlobalVar.varDB = args[i].Substring(4);
                else if (args[i].ToUpper().StartsWith("/CUBE:"))
                    GlobalVar.varCube = args[i].Substring(6);
                else if (args[i].ToUpper().StartsWith("/STEP:"))
                    GlobalVar.varStep = Convert.ToInt32(args[i].Substring(6));
                else if (args[i].ToUpper().StartsWith("/PARALLEL:"))
                    GlobalVar.varMaxParallel = Convert.ToInt32(args[i].Substring(10));
                else if (args[i].ToUpper().StartsWith("/XML:"))
                    GlobalVar.varFileXML = args[i].Substring(5);
                else if (args[i].ToUpper().StartsWith("/DAY:"))
                    GlobalVar.varDayProcess = Convert.ToInt32(args[i].Substring(5));
                else if (args[i].ToUpper().StartsWith("/STATE:2"))
                    GlobalVar.varStep = -9998;
                else if (args[i].ToUpper().StartsWith("/STATE"))
                    GlobalVar.varStep = -9999;
                else if (args[i].ToUpper().StartsWith("/DATESTART:"))
                    GlobalVar.varDateStartProcess = DateTime.Parse(args[i].Substring(11));
                else if (args[i].ToUpper().StartsWith("/DAYPROCESS:"))
                    GlobalVar.varDayProcess = Convert.ToInt32(args[i].Substring(12));
                else if (args[i].ToUpper().StartsWith("/PROCESSDIMENSION:"))
                    MyXMLA.SetProcessTypeDimension(args[i].Substring(17));
                else if (args[i].ToUpper().StartsWith("/PROCESSCUBE:"))
                    MyXMLA.SetProcessTypeCube(args[i].Substring(13));
                else if (args[i].ToUpper().StartsWith("/ARX"))
                {
                    GlobalVar.varIsArx = true;
                    if (args[i].ToUpper().Length == 14)
                        GlobalVar.varArxDate = DateTime.ParseExact(args[i].ToUpper().Substring(5), "dd.MM.yyyy", CultureInfo.InvariantCulture);
                }
                else if (args[i].ToUpper().StartsWith("/?"))
                {
                    Console.Write(varKey);
                    Console.ReadKey(true);
                    return;
                }
                else
                {
                    Console.Write("Колюч=>" + args[i].ToUpper() + " невірний. " + varKey);
                    Console.ReadKey(true);
                    return;
                }
        	}
       
            
        	GlobalVar.varFileLog = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)+
        		"\\log\\process_"+GlobalVar.varDB.Trim()+"_"+
        		DateTime.Now.ToString("yyyyMMdd")+"_"+ GlobalVar.varStep.ToString().Trim() +".txt";

        	Log.log("START=> /Server:"+GlobalVar.varServer+" /DB:"+GlobalVar.varDB+" /CUBE:"+GlobalVar.varCube+" /Step: "+GlobalVar.varStep.ToString()   );
            XMLABuilder.Process(@"Data Source="+GlobalVar.varServer+";Provider=msolap;", GlobalVar.varDB,GlobalVar.varCube,GlobalVar.varStep,GlobalVar.varMetod);
            Log.log("END=> /Server:"+GlobalVar.varServer+" /DB:"+GlobalVar.varDB+" /CUBE:"+GlobalVar.varCube+" /Step:"+GlobalVar.varStep.ToString()   );
            if(GlobalVar.varStep<-9990)
        		Console.ReadKey(true);
        }
    }
    
}