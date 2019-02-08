using System;
//using NetExcel = Microsoft.Office.Interop.Excel; 
using adomd = Microsoft.AnalysisServices.AdomdClient;
using System.Data;
using System.Data.OleDb;


namespace Mail
{
	class Program
	{
		public static void Main(string[] args)
		{
			
			Mail MyMail =new Mail();
			Console.WriteLine("Hello World!");
			double[,] aa = new double[2,8];
			 MyMail.SendShops();
			// TODO: Implement Functionality Here
				Console.WriteLine(aa.GetLength(0));
			Console.WriteLine (aa.GetLength(1));
			Console.Write("Press any key to continue . . . ");
			//Console.ReadKey(true);
		}
	}

	class Mail
	{
      object misValue = System.Reflection.Missing.Value;
      string varTmp="d:\\temp\\";
      string varSrtServer="srv-bat";
      string varStrCatalog="DW_OLAP";
      Proc proc;
      Cube cube;
      public Mail()
      {
      	proc=new Proc();
      	cube=new Cube("Data Source=" + varSrtServer + ";Provider=msolap;Initial Catalog=" + varStrCatalog);
      }
       //adomd.AdomdConnection parConnection;    
       public void SendShops()
       {
       		string varStr=@"select wh.code_warehouse,wh.name_warehouse, mz.pack_opts.opts_get_vchar_value(20000074,shw.code_shop ) mail
       from mz.warehouse wh , mz.shop_warehouse shw 
        where wh.code_warehouse=shw.code_warehouse and 
        c.proc.GetTypeWarehouse(wh.code_warehouse)=1
              and length(trim(mz.pack_opts.opts_get_vchar_value(20000074,shw.code_shop ) ))>0";

                DataSet dataSet = new DataSet();
                DataTable TTable = dataSet.Tables.Add("table");
                String connectionString = "Provider=OraOLEDB.Oracle.1;Data Source=mer;Persist Security Info=True;user id=c;password=c";
                OleDbConnection myOleDbConnection = new OleDbConnection(connectionString);
                OleDbDataAdapter adapterTable =  new OleDbDataAdapter(varStr, myOleDbConnection);
                adapterTable.Fill(TTable);
                myOleDbConnection.Close();
                foreach (DataRow row in TTable.Rows)
                	SendShop( Convert.ToInt32(row[0]), Convert.ToString(row[2]));
       			
       }
        
       public void SendShop(int parCodeWarehouse,string parMail,int parYearMonth = 0)
       {
       	//int varCodeWarehouse;
       	string varStrCodeWarehouse=Convert.ToString(parCodeWarehouse);
       	if(parYearMonth==0) 
       		parYearMonth=DateTime.Now.Year*100+ DateTime.Now.Month;
       	string varStrYearMonth=Convert.ToString(parYearMonth);
       	
        System.Globalization.CultureInfo oldCI   =   System.Threading.Thread.CurrentThread.CurrentCulture;
        System.Threading.Thread.CurrentThread.CurrentCulture =   new System.Globalization.CultureInfo("en-US");
       	try
        {
       	  string varMDX;
       	  adomd.CellSet varCellSet;
       	  
       	  //double[,]  varArr = new double [2,5] { { 8, 10, 2, 6, 3 },{ 1, 2, 3, 4, 5 }};
    	  MyExcel Ex = new MyExcel("D:\\WORK\\CS4\\cube\\Mail\\Новий_Аналiз.xls",varTmp+"Shop_"+parCodeWarehouse.ToString()+".xls");
    	  Ex.SetWorksheet(1);
    	  //
    	  varMDX = @"
    	  	with member [DN] as [Час].[Календар].PROPERTIES('KEY')
				SELECT
				{[Measures].[прод_грн],[Measures].[вал_без_пдв],[Measures].[PLAN TO],[Measures].[PLAN VAL],[DN] } ON COLUMNS,
				  NON EMPTY {[Час].[Календар].[День].AllMembers} 
					DIMENSION PROPERTIES PARENT_UNIQUE_NAME, CHILDREN_CARDINALITY ON ROWS
				FROM [Рух товарів]
				WHERE ([Час].[Рік_Місяць].[Місяць].&[" + varStrYearMonth + @"],
				[Склади].[Склади].[All].&["+varStrCodeWarehouse+"]) ";
    	  
			varCellSet = cube.RunMDXCellSet(varMDX);
			if(varCellSet!=null)
			{
 			 Ex.CellSet(3 ,3  ,varCellSet,new int[1] {0});
			 Ex.CellSet(3 ,13 ,varCellSet,new int[2] {1,3});
			 Ex.CellSet(3 ,7  ,varCellSet,new int[1] {2});
			// Ex.CellSet(3 ,14 ,varCellSet,new int[1] {3});
			}
			
				
    	  //    	  Ex.Array(2,3,varArr);
    	  
    	  Ex.SetWorksheet("відємна зведена");
    	  Ex.PivotTablesFields("СводнаяТаблица1",new string [] {"[Магазини].[По регіонах]"},new string []{"[Магазини].[По регіонах].[All]"});
//    	  Ex.PivotTables("СводнаяТаблица1").PivotFields("[Магазини].[По регіонах]").CurrentPageName = "[Магазини].[По регіонах].[All]";
//    	  Ex.PivotTables("СводнаяТаблица1").PivotFields("[Час].[Рік_Місяць]").CurrentPageName = "[Час].[Рік_Місяць].[День].&["+ varStrYearMonth + "]";
    	  	
    	  Ex.Close();
    	  MyMail.SendMailZIP("gelo@pakko.ua", varTmp+"Shop_"+parCodeWarehouse.ToString()+".xls","Щоденна розсилання" );
    	  
       	}
        finally
        {
         System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;        
        }
        
       }
	}
}