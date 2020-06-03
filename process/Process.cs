/* Призначення програми - автоматичний процесінг кубів. + автоматичне створення партицій по потребі.
* Процесінг кубів може бути розбитий на кроки.Програма вміє очікувати завершення розрахунку даних в SQL
* А також вміє контролювати допустимий період процесінгу наприклад щоб робочий час не процесилось.
* Типічний рядок запуску в job: D:\Process\process.exe /DB:DW_OLAP3 /Server:localhost /step:10
* Запуск для процесінга куба, наприклад після змін з командного рядка: process.exe /db:dw_olap3 /cube:чеки_товар /parallel:2
* Програма вміє перезапускати службу MSSQLServerOLAPService а при старті крока перевіряє чи служба запущена і якщо ні пробує її запустити.
* Права на це повинні бути у користувача від чийого імені вона стартує.
* Групи мір, яку необхідно процесити в даному кроці визначається наступним чином.
* в полі description групи мір має бути рядок виду pr=>31,10; pr=> - признак що це інфа про крок і період перепроцесінга.крок в даному випадку 31, 10 -скільки днів перепроцешувати партіцию за попередній період.
* Якщо  description незаповнений або не починається з pr=>  то крок по замовчуванню для цьої групи мір - 11.
* Партіциї створються на основі партіциї з назвами  template,template_Month,template_Quarter,template_Year,template_4Week,template_Week
* Партіциї template - це помісячні партіциї в них запит має відповідати шаблону
* SELECT * FROM "DW"."FACT_REST_DAY"   WHERE 1=0 and date_report>=to_date('20130601','YYYYMMDD')
* Де 20130601 - початковий період.Доступо створення партіций як для Oracle так MSSQL.
* Для MSSQL запит виду SELECT * FROM "DW"."FACT_REST_DAY"   WHERE 1=0 and date_report>='20130601' 
* Для партіций з назвами template_Month, template_Quarter, template_Year, template_4Week, template_Week
* дещо інший спосіб формування запитів. Вони допускають довільної складності запит.
* Логіка формування наступна: to_date('20130601','YYYYMMDD') -вказує на початок періоду і заміняється на реальний початок періоду а
* to_date('00010101','YYYYMMDD') заміняється на кінець періоду
* Доступо створення партіций тільки для Oracle
* Для шаблонів template_4Week, template_Week дата початку періоду може бути будь-який день тижня - визначається початковою датою.
*
* Програма має конфігураційний файл process.xml Параметри з рядка запуску мають перевагу над конфігураційним файлом.
* Типовий конфігураційний файл у нас.Достатньо очевидні параметри.
<Config>
 <Server>srv-bat</Server>
 <Database>DW_OLAP</Database>
 <XMLA maxParallel = "6" > true </ XMLA >
 < DefaultStep > 0 </ DefaultStep >
 < ConectSQL > Provider = OraOLEDB.Oracle.1; Data Source = mer; Persist Security Info=True;user id = *; password=*</ConectSQL>
 <Metod>Fast</Metod> <!-- default:Fast(Fact, Full, Normal)  -->
 <KeyErrorLogPath>d:\process\log</KeyErrorLogPath> <!-- delault - program_path\LOG ; path -server\path -->
 <ServicesOlap>MSSQLServerOLAPService</ServicesOlap> 
 <Services>MSSQLServer,AgentMSSQLServer</Services> 
 <Step0>
  <Time Start = "7" End="24">true</Time>
  <ProcessDimension>UPDATE</ProcessDimension> <!-- default:UPDATE(UPDATE, FULL) --> 
  <RestartServicesOlap>1</RestartServicesOlap> <!-- default:0 - no restart,1 - before,2-after,3 before and after  -->
 </Step0>
 <Step1>
  <PrepareSQL> begin null; end;  </PrepareSQL> <!-- planed  -->
 </Step1>
 <Step2>
   <WaitSQL>SELECT* FROM dw.v_statecons</WaitSQL> 
 </Step2>
 
 <Step3>
   <WaitSQL>SELECT* FROM dw.v_state_MIN_MAX</WaitSQL> 
 </Step3>
 <Step99>
  <Services>1<Services> <!-- default:0 - no restart,1 - before,2-after,3 before and after  -->
 <Step99>
</Config>
*/

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
                GlobalVar.varRestartServicesOlap = (myXML.GetVar("Step" + GlobalVar.varStep.ToString().Trim(), "RestartServicesOlap") == null ? GlobalVar.varRestartServicesOlap : Convert.ToInt32(myXML.GetVar("Step" + GlobalVar.varStep.ToString().Trim(), "RestartServicesOlap")));
                GlobalVar.varServicesOlap = (myXML.GetVar("ServicesOlap") == null ? GlobalVar.varDB : myXML.GetVar("ServicesOlap"));

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

            if (1==0 && GlobalVar.varServicesOlap.Trim().Length > 0)
            {
                var varService = new Services(GlobalVar.varServicesOlap, GlobalVar.varServer);
                if (!varService.IsStart() || GlobalVar.varRestartServicesOlap == 1 || GlobalVar.varRestartServicesOlap == 3)
                {
                    if (varService.IsStart())
                    {
                        Log.log("Try ReStart =>" + GlobalVar.varServicesOlap + " in " + GlobalVar.varServer);
                        if (!varService.ReStart())
                        {
                            Log.log("No ReStart =>" + GlobalVar.varServicesOlap + " in " + GlobalVar.varServer + "\n" + varService.log);
                            return;
                        }
                        else
                            Log.log("ReStart OK");
                    }
                    else
                    {
                        if (!varService.Start())
                        {
                            Log.log("Try start =>" + GlobalVar.varServicesOlap + " in " + GlobalVar.varServer);
                            if (!varService.Start())
                            {
                                Log.log("No start =>" + GlobalVar.varServicesOlap + " in " + GlobalVar.varServer + "\n" + varService.log);
                                return;
                            }
                            else
                                Log.log("Start OK");
                        }
                    }
                }
            }

            XMLABuilder.Process(@"Data Source="+GlobalVar.varServer+";Provider=msolap;", GlobalVar.varDB,GlobalVar.varCube,GlobalVar.varStep,GlobalVar.varMetod);
            Log.log("END=> /Server:"+GlobalVar.varServer+" /DB:"+GlobalVar.varDB+" /CUBE:"+GlobalVar.varCube+" /Step:"+GlobalVar.varStep.ToString()   );
            if(GlobalVar.varStep<-9990)
        		Console.ReadKey(true);
        }
    }
    
}