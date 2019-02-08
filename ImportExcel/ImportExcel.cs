using System;
using System.Text;
using System.IO;

namespace ImportExcel
{
	class Program
	{
		public static void Main(string[] args)
		{
			if(args.Length !=1 || args[0].Length==0)
			{
				Console.WriteLine("Необхідно запускати ImportExcel.exe file.xls");
				Console.ReadKey(true);
				return;
			}
		string varFile=args[0],varCellBudgetItem,varCommand;
		int VarColumnCode, varColumnFirstMonth,varColumnLastMonth,varStartMonth,varEndMonth,varStartRow,varEndRow,varYear,varStartPage,varEndPage;
		int varCodeBudgetItem,varCodeMonth, varCodeStructure;
		int i;
		int []varArrayCode;
		string varConnect;
		string varDate;
		string varFileXML=Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)+"\\config.xls";
		
		MyExcel Ex,ExConfig;
		
		if(Path.GetExtension(varFile).ToUpper()!="XLS" || Path.GetExtension(varFile).ToUpper()!="XLSX")
		   {
		   	Console.WriteLine("Файл=>" + varFile + "Не є Ексель файлом" );
		   }
		
        Console.WriteLine("Імпортуємо файл=>" + varFile );
        try
        {
        	Ex = new MyExcel(varFile);
		}
 		catch (Exception e)
 		{
 			Console.WriteLine( "Помилка при відкритті Ексель=>"+ e.Message );
 			Console.ReadKey(true);
			return;
 		}
 		
		if (File.Exists(varFileXML))
		{
		   try
        	{
        	ExConfig = new MyExcel(varFileXML);
			}
 			catch (Exception e)
 			{
 			Console.WriteLine( "Помилка при відкритті Config.xls =>"+ e.Message );
 			Console.ReadKey(true);
			return;
			}
		}
		else
			ExConfig = new MyExcel(varFileXML);
			//ExConfig=Ex;

			
 		
		/* Відкриваємо конект і зчитуємо запит на оновлення */
 		
    	ExConfig.SetWorksheet("Налаштування");
    	
		varConnect = ExConfig.GetStringCell(13,2);
 		StringBuilder script = new StringBuilder();
 		for(i=14;i<50;i++)
 		{
			if(Ex.GetStringCell(i,2).Trim().Length>0)
			 script.AppendLine(Ex.GetStringCell(i,2));
 		}
 		varCommand= script.ToString();
		
 		Console.WriteLine( "Відкриваємо базу=>"+varConnect );
 		MyOleDb  varOleDb;
 		try 
 		{
 		  varOleDb= new MyOleDb(varConnect);
 		}
 		catch (Exception e)
 		{
 			Console.WriteLine( "Помилка при відкритті Бази=>"+ e.Message );
 			Console.ReadKey(true);
			return; 
 		}
 		

    	int j = 2;
		while (ExConfig.GetStringCell(4,j).Trim().Length>0 )
		{
			
			varColumnFirstMonth = ExConfig.GetStringCell(4,j).ToUpper()[0]-'A'+1 ;
			varColumnLastMonth = ExConfig.GetStringCell(5,j).ToUpper()[0]-'A'+1 ;
			varStartMonth = ExConfig.GetIntCell(6,j);
    		varEndMonth = ExConfig.GetIntCell(7,j);
    		varStartRow = ExConfig.GetIntCell(8,j);
    		varEndRow = ExConfig.GetIntCell(9,j);
    		varCellBudgetItem = ExConfig.GetStringCell(10,j);
    		varStartPage = ExConfig.GetIntCell(11,j);
    		varEndPage = ExConfig.GetIntCell(12,j);
    		
			if (ExConfig.GetStringCell(2,j).ToUpper()[0]>='A')
			{
				VarColumnCode = ExConfig.GetStringCell(2,j).ToUpper()[0]-'A'+1 ;
			    varArrayCode = new int[0];
			 }
			else
			{  VarColumnCode= ExConfig.GetIntCell(2,j);
				varArrayCode = new int[varEndRow];
				for(i=0;i<varEndRow; i++)
					varArrayCode[i]= ExConfig.GetIntCell(VarColumnCode,i);
				VarColumnCode=-1;
			}
				
			if (ExConfig.GetStringCell(3,j).ToUpper()[0]>='A') 
			{
				varYear=0;
				varDate	= ExConfig.GetStringCell(3,j).ToUpper();
			}
			 else
			{
				varYear = ExConfig.GetIntCell(3,j);
				varDate = null;
			}
    	 		
 		
 		for(int varPage=varStartPage; varPage<=varEndPage;varPage++)
 		    {
				Ex.SetWorksheet(varPage);
		    
				varCodeBudgetItem =Ex.GetIntCell(varCellBudgetItem);
				Console.WriteLine( "Імпортую сторінку=>"+varPage.ToString()+" Статя => "+varCodeBudgetItem.ToString() );
				for(int varRow=varStartRow;varRow<=varEndRow;varRow++)
				{   if(VarColumnCode>0)
					 varCodeStructure=Ex.GetIntCell(varRow, VarColumnCode);
					else
						varCodeStructure= varArrayCode[varRow];
					
					if(varCodeStructure>0)
					for(int varMonth=varStartMonth;varMonth<=varEndMonth;varMonth++)
					{
						varCodeMonth=varYear*100+varMonth;
						double varValue = Ex.GetDoubleCell(varRow,varColumnFirstMonth+varMonth-1);
						string varExCommand = 
						varCommand.Replace("?parMonth",		Convert.ToString(varCodeMonth)).
								   Replace("?parStructure",	Convert.ToString(varCodeStructure)).
							       Replace("?parAccount",	Convert.ToString(varCodeBudgetItem)).
							 	   Replace("?parAmount",	Convert.ToString(varValue));
							                                                                   
						 varOleDb.RunSQL(varExCommand);
					}
					
				}
			
 		    }
			
		}
			Ex.Close();
			
			
			Console.WriteLine("Завершили. Натисніть будь-яку клавішу" );
			Console.ReadKey(true);
		}

	}


}