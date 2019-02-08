
using System;
using System.Data;
using System.Data.OleDb;

namespace DescriptionTable
{
	class Program
	{
		public static void Main(string[] args)
		{
			
			DataSet dataSet = new DataSet();
			DataTable TTable = dataSet.Tables.Add("table");
			DataTable TCol = dataSet.Tables.Add("columns");
			string varSeparator = "\t";
			Proc p=new Proc() ;
			p.CreateLog("d:\\DW_TABLE.txt");
			
			string connectionString = "provider=MSDAORA;data source=Mer;user id=c;password=c";
    OleDbConnection myOleDbConnection = new OleDbConnection(connectionString);

    
    OleDbDataAdapter adapterTable = 
      new OleDbDataAdapter(	"select t.TABLE_NAME,t.COMMENTS,t.owner from sys.all_tab_comments t where  t.owner = 'DW' and Table_type='TABLE'  union "+
						    "select t.TABLE_NAME,substr(t.COMMENTS,4),t.owner from sys.all_tab_comments t where  t.owner = 'C' and Table_type='TABLE' and  substr(t.COMMENTS,1,3)='DW;' "+
						    "union select t.TABLE_NAME,t.COMMENTS,t.owner from sys.all_tab_comments t where  t.owner = 'SHAKH'  and TABLE_NAME='MIN_MAX_REST'", myOleDbConnection);
    adapterTable.Fill(TTable);
    p.Log("Таблички");
    foreach (DataRow row in TTable.Rows)
    {
    	p.Log(row[2].ToString()+'.'+row[0].ToString()+varSeparator+row[1].ToString());
     	string varSQL="select col.COLUMN_NAME, col.DATA_TYPE, com.Comments "+
       		"from sys.all_tab_columns col, sys.all_col_comments com "+
       		"where col.owner = '"+row[2].ToString()+"' and col.table_name = '"+row[0].ToString()+"' "+
       		"and com.Owner (+) = col.owner "+
       		"and com.Table_Name (+) = col.table_name "+
       		"and com.Column_Name (+) = col.Column_Name "+
       		"order by col.column_id ";
    	
    	
      	OleDbDataAdapter adapterCol = 
      	new OleDbDataAdapter(varSQL, myOleDbConnection);
    	
    	adapterCol.Fill(TCol);
    	foreach (DataRow rowCol in TCol.Rows)
    	{
    			p.Log(varSeparator+rowCol[0].ToString()+varSeparator+rowCol[1].ToString()+varSeparator+rowCol[2].ToString());
    	}
		TCol.Clear();
//    	Console.WriteLine(row[0].ToString() );
//    	Console.WriteLine(row[1].ToString()) ;
    }
    p.Log("Представлення");
      adapterTable = 
      new OleDbDataAdapter("select t.TABLE_NAME,t.COMMENTS,t.owner from sys.all_tab_comments t where  t.owner = 'DW' and Table_type='VIEW'", myOleDbConnection);
      TTable.Clear();
    adapterTable.Fill(TTable);
    foreach (DataRow row in TTable.Rows)
    	p.Log(row[2].ToString()+'.'+row[0].ToString()+varSeparator+row[1].ToString());
    	
    	
    myOleDbConnection.Close(); 

    p.CloseLog();
    	//Console.ReadLine();
		}
	}
}