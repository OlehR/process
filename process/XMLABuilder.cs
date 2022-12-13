using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.AnalysisServices;
using Microsoft.AnalysisServices.Xmla;

namespace Process
{
    static class XMLABuilder
    {
        static List<ProcessingOnLine> GlobalListOnLine = new List<ProcessingOnLine>();
        static List<ProcessingOnLine> LocalListOnLine = new List<ProcessingOnLine>();
        static XmlaClient XMLACl;
        static private Partition CurrentPartitionFind(MeasureGroup aMeasureGroup, int aThisYear, int aThisMonth)
        {
            return PartitionFind(aMeasureGroup, aThisYear, aThisMonth);
        }
        static private Partition PreviousPartitionFind(MeasureGroup aMeasureGroup, int aThisYear, int aThisMonth)
        {
            int prevYear = aThisMonth == 1 ? aThisYear - 1 : aThisYear;
            int prevMonth = aThisMonth == 1 ? 12 : aThisMonth - 1;
            return PartitionFind(aMeasureGroup, prevYear, prevMonth);
        }
        static private Partition PartitionFind(MeasureGroup aMeasureGroup, int aYear, int aMonth)
        {
            string suffix = ToYYYYMM(aYear, aMonth);
            foreach (Partition p in aMeasureGroup.Partitions)
                if (p.Name.EndsWith(suffix))
                    return p;
            return null;
        }
        static private Partition PartitionFind(MeasureGroup aMeasureGroup, DateTime varDT)
        {
            string suffix = ToYYYYMM(varDT);
            foreach (Partition p in aMeasureGroup.Partitions)
                if (p.Name.EndsWith(suffix))
                    return p;
            suffix = ToYYYYMMDD(varDT);
            foreach (Partition p in aMeasureGroup.Partitions)
                if (p.Name.EndsWith(suffix))
                    return p;
            return null;
        }


        /*        static private Partition PartitionCreate(int aYear, int aMonth, bool aForever, string aTableName,
            MeasureGroup aMeasugeGroup, Partition aTemplatePartition, string DateFieldName, bool isOracle)
        {
            try
            {
                Partition p = aMeasugeGroup.Partitions.Add(aMeasugeGroup.Name + " " + ToYYYYMM(aYear, aMonth));
                p.AggregationDesignID = aTemplatePartition.AggregationDesignID;
                DateTime dStart = new DateTime(aYear, aMonth, 1);
                DateTime dEnd = new DateTime(aMonth == 12 ? aYear + 1 : aYear, aMonth == 12 ? 1 : aMonth + 1, 1);

                string sql = "select * from " + aTableName + " where " + DateFieldName + " >= " + (isOracle ? "to_date('" + ToYYYYMMDD(dStart) + "','YYYYMMDD')" : ToYYYYMMDD(dStart));
                if (!aForever)
                    sql += " AND " + DateFieldName + " < " + (isOracle ? "to_date('" + ToYYYYMMDD(dEnd) + "','YYYYMMDD')" : ToYYYYMMDD(dEnd));
                p.Source = new QueryBinding(aTemplatePartition.DataSource.ID, sql);
                p.Slice = "[Дата].[Місяць]" + ".&[" + ToYYYYMM(aYear, aMonth) + "]";
                return p;
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
            }
            return null;
        }*/


        /// <summary>
        /// Створюємо партіцию на основі шаблона.
        /// </summary>
        /// <param name="parTemplatePartition">Шаблон</param>
        /// <param name="parDStart">Початкова дата партіциї </param>
        /// <param name="parDEnd">Цінцева дата Партіциї</param>
        /// <param name="parType">Періодичність партіциї</param>
        /// <returns></returns>
        static private Partition PartitionCreate(Partition parTemplatePartition, DateTime parDStart, DateTime parDEnd, TypePeriod parType)
        {
            var MeasugeGroup = parTemplatePartition.Parent;
            try
            {

                string varSQL = (parTemplatePartition.Source as QueryBinding).QueryDefinition;
                bool varIsOracle = IsOracle(varSQL);
                string varStartDate = XMLABuilder.GetStartDate(varSQL); //Початкова дата для створення партіций
                

                string varNewSQL;
                if (varSQL.IndexOf("1=0") > 0)
                {
                    string varDateField = DateFieldNameFromQueryGet(varSQL);//tmp
                    string table = TableNameFromQueryGet(varSQL); //tmp
                    string Where = WhereNameFromQueryGet(varSQL);
                    varNewSQL = $"select * from {table}  where {Where} {varDateField} >= " + (varIsOracle ? "to_date('" + ToYYYYMMDD(parDStart) + "','YYYYMMDD')" : ToYYYYMMDD(parDStart))+" AND " + varDateField + " < " + (varIsOracle ? "to_date('" + ToYYYYMMDD(parDEnd) + "','YYYYMMDD')" : ToYYYYMMDD(parDEnd));
                }
                else
                {
                    varNewSQL = varSQL.Replace(">= ", ">=").Replace(">= ", ">=").Replace(">= ", ">=").Replace("< ", "<").Replace("< ", "<").Replace("< ", "<").
                        Replace(">=to_date('" + varStartDate, ">= to_date('" + ToYYYYMMDD(parDStart)).Replace("<to_date('00010101", "< to_date('" + ToYYYYMMDD(parDEnd));
                }

                Partition p = MeasugeGroup.Partitions.Add(MeasugeGroup.Name + " " + (parType ==TypePeriod.Week || parType==TypePeriod.Week4 ? ToYYYYMMDD(parDStart) : ToYYYYMM(parDStart)));
                p.AggregationDesignID = parTemplatePartition.AggregationDesignID;
                p.Source = new QueryBinding(parTemplatePartition.DataSource.ID, varNewSQL);
                if (parType == TypePeriod.Month)
                    p.Slice = /*"[Час].[Календар].[Місяць]*/"[Дата].[Місяць]" + ".&[" + ToYYYYMM(parDStart) + "]";
                return p;
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
            }
            return null;
        }

        private static string WhereNameFromQueryGet(string aSql)
        {
            string res = "";
            string sqlmod = aSql.Replace("\n", " ").Replace("\t", " ");
            string rest = sqlmod.Substring(sqlmod.ToLower().IndexOf("where") + 5).Trim();
            int ind = rest.IndexOf("1=0");
            if (ind > 0)
                res = rest.Substring(0, ind);
            // Console.WriteLine("@" + rest.Substring(0, rest.IndexOf(" ")) + "@");
            return res;
        }

        private static string TableNameFromQueryGet(string aSql)
        {
            string sqlmod = aSql.Replace("\n", " ").Replace("\t", " ");
            string rest = sqlmod.Substring(sqlmod.ToLower().IndexOf("from") + 5).Trim();
            // Console.WriteLine("@" + rest.Substring(0, rest.IndexOf(" ")) + "@");
            return rest.Substring(0, rest.IndexOf(" ")).Trim();
        }
        private static string DateFieldNameFromQueryGet(string aSql)
        {
            int ind = aSql.IndexOf("1=0");
            if (ind > 0)
                aSql = aSql.Substring(ind + 4);

            string sqlmod = aSql.Replace("\n", " ").Replace("\t", " ");
            string rest = sqlmod.Substring(sqlmod.ToLower().IndexOf("and") + 4).Trim();
            // Console.WriteLine("@" + rest.Substring(0, rest.IndexOf(" ")) + "@");
            return rest.Substring(0, rest.IndexOf(">")).Trim();
        }
        public static string GetStartDate(string aSql)
        {
            string sqlmod = aSql.Replace("\n", " ").Replace("\t", " ").Replace(">= ", ">=").Replace(">= ", ">=").Replace(">= ", ">=").Replace(">= ", ">=");
            string rest;
            if (sqlmod.ToLower().IndexOf(">=to_date(")!=-1)
              rest = sqlmod.Substring(sqlmod.ToLower().IndexOf(">=to_date(") + 11).Trim();
            else
                rest = sqlmod.Substring(sqlmod.ToLower().IndexOf(">=") + 2).Trim();
            return rest.Substring(0, 8);
        }

        private static bool IsOracle(string aSql)
        {
            string sqlmod = aSql.Replace("\n", " ").Replace("\t", " ");
            return (sqlmod.ToLower().IndexOf("to_date(") > 0);
        }

        private static bool IsCurrentProcess(string varStr, int parStep)
        {
            if (parStep == 0) return true;
            if (varStr == null) { if (parStep == 11) return true; else return false; };
            if ((varStr.Trim().ToLower().Substring(0, 4) == "pr=>"))
            {
                if (Convert.ToInt16(strings.GetWordNum(strings.GetWordNum(varStr.Trim().Substring(4), 1, ";"), 1, ",")) == parStep)
                    return true;
            }
            else
                if (parStep == 11) return true;
            return false;
        }

        private static int DayProcess(string varStr)
        {
            if (GlobalVar.varDayProcess > 0) return GlobalVar.varDayProcess;
            int varRes = GlobalVar.varDefaultDayProcess;
            if (varStr == null) return varRes;
            if ((varStr.Trim().ToLower().Substring(0, 4) == "pr=>"))
            {
                try
                {
                    varRes = Convert.ToInt16(strings.GetWordNum(strings.GetWordNum(varStr.Trim().Substring(4), 1, ";"), 2, ","));
                }
                catch
                {                    
                }
            }
            return varRes;
        }

        static private string ToYYYYMM(int aYear, int aMonth)
        {
            return aYear.ToString() + (aMonth < 10 ? "0" : "") + aMonth.ToString();
        }

        static private string ToYYYYMM(DateTime aDate)
        {
            int year = aDate.Year;
            int month = aDate.Month;
            return year.ToString() + (month < 10 ? "0" : string.Empty) + month.ToString();
        }

        static private string ToYYYYMMDD(DateTime aDate)
        {
            int year = aDate.Year;
            int month = aDate.Month;
            int day = aDate.Day;
            return year.ToString() + (month < 10 ? "0" : string.Empty) + month.ToString() +
                   (day < 10 ? "0" : string.Empty) + day.ToString();
        }
        static private ProcessType SafeProcTypeGet(IProcessable aObject, ProcessType aNeededProcessType)
        {
            if (aObject.State != AnalysisState.Processed)
                return ProcessType.ProcessFull;
            return aNeededProcessType;
        }

        static public void WaitOracle(int parStep)
        {
            if (GlobalVar.varWaitSQL != null && GlobalVar.varConectSQL != null)
            {
                DataSet dataSet = new DataSet();
                DataTable TTable = dataSet.Tables.Add("table");
                string varSqlConect = GlobalVar.varConectSQL;
                int state = 0;

                Log.log("Star WaitOracle");
                do
                {
                    OleDbConnection myOleDbConnection = new OleDbConnection(GlobalVar.varConectSQL);
                    OleDbDataAdapter adapterTable =
                        new OleDbDataAdapter(GlobalVar.varWaitSQL, myOleDbConnection);
                    adapterTable.Fill(TTable);
                    foreach (DataRow row in TTable.Rows)
                        state = Convert.ToInt16(row[0]);
                    TTable.Clear();
                    myOleDbConnection.Close();
                    if (state == 0) Thread.Sleep(1000 * 5 * 60); //5 minets
                } while (state == 0);
                Log.log("End WaitOracle");
            }
        }



        static public TypePeriod GetTypePartition(MeasureGroup g)
        {

            if (g.Partitions.FindByName("template") != null)
                return TypePeriod.Month;
            else if (g.Partitions.FindByName("template_Month") != null)
                return TypePeriod.Month;
            else if (g.Partitions.FindByName("template_Quarter") != null)
                return TypePeriod.Quarter;
            else if (g.Partitions.FindByName("template_Year") != null)
                return TypePeriod.Year;
            else if (g.Partitions.FindByName("template_4Week") != null)
                return TypePeriod.Week4;
            else if ( g.Partitions.FindByName("template_Week") != null)
                return TypePeriod.Week;
            return TypePeriod.NotDefined;
        }

        static public Partition GetTemplatePartition(MeasureGroup g)
        {
            Partition pTemplate = null;
            if ((pTemplate = g.Partitions.FindByName("template")) != null)
                return pTemplate;
            else if ((pTemplate = g.Partitions.FindByName("template_Month")) != null)
                return pTemplate;
            else if ((pTemplate = g.Partitions.FindByName("template_Quarter")) != null)
                return pTemplate;
            else if ((pTemplate = g.Partitions.FindByName("template_Year")) != null)
                return pTemplate;
            else if ((pTemplate = g.Partitions.FindByName("template_4Week")) != null)
                return pTemplate;
            else if ((pTemplate = g.Partitions.FindByName("template_Week")) != null)
                return pTemplate;
            return pTemplate;
        }


        static public void CreatePartition(Cube parCube)
        {
            DateTime varNow = DateTime.Now;
            DateTime varDateArxCube = new DateTime(1, 1, 1);
            CultureInfo provider = CultureInfo.InvariantCulture;
            if (parCube.Description != null && strings.GetWordNum(parCube.Description, 2, ";").Length > 5 && strings.GetWordNum(parCube.Description, 2, ";").Substring(0, 5).ToUpper() == "ARX=>")
                varDateArxCube = DateTime.ParseExact(strings.GetWordNum(parCube.Description, 2, ";").Substring(5, 10).ToUpper(), "dd.MM.yyyy", provider);
            foreach (MeasureGroup g in parCube.MeasureGroups)
                try
                {
                    Partition pTemplate = GetTemplatePartition(g);
                    TypePeriod varType = GetTypePartition(g);
                    if (varType != TypePeriod.NotDefined)
                    {
                        Partition currentPartition = null;

                        if (pTemplate.Source is TableBinding)
                            throw new ApplicationException(
                                "template partition should have query binding \"select * from tablename where 1=0\"");
                        //string table = TableNameFromQueryGet((pTemplate.Source as QueryBinding).QueryDefinition);
                        string varStrStartDate = GetStartDate((pTemplate.Source as QueryBinding).QueryDefinition);
                         DateTime varRealStartDate = new DateTime(Convert.ToInt32(varStrStartDate.Substring(0, 4)), Convert.ToInt32(varStrStartDate.Substring(4, 2)), Convert.ToInt32(varStrStartDate.Substring(6, 2)));
                        DateTime varStartDate;
                        if (GlobalVar.varIsArx)
                        {
                            if (GlobalVar.varArxDate == new DateTime(1, 1, 1))
                                varStartDate = varDateArxCube;
                            else
                                varStartDate = GlobalVar.varArxDate;
                           
                        }
                        else
                            varStartDate = varRealStartDate;
                        DateTime varEndDate = varStartDate;

                        //string varDateField = "", table = "";
/*                        if (g.Partitions.FindByName("template") != null) //TMP
                        {
                            varDateField = DateFieldNameFromQueryGet((pTemplate.Source as QueryBinding).QueryDefinition);//tmp
                            table = TableNameFromQueryGet((pTemplate.Source as QueryBinding).QueryDefinition); //tmp
                        }*/
                        while (varStartDate <= varNow)
                        {
                            currentPartition = PartitionFind(g, varStartDate);
                            varEndDate = CountNextPreviousDate(varStartDate, varType);
                            if (currentPartition == null)
                            {
                             //   if (g.Partitions.FindByName("template") == null)
                                    currentPartition = PartitionCreate(pTemplate, varStartDate, varEndDate, varType);
                             // else
                             //     currentPartition = PartitionCreate(varStartDate.Year, varStartDate.Month, false, table, g, pTemplate, varDateField, varIsOracle);

                                currentPartition.Update();
                                Console.WriteLine("Створено партіцию=>" + currentPartition.Name);
                            }
                            varStartDate = varEndDate;
                        }


                    }
                }
                catch (Exception e)
                {
                    Log.log("Група мір=>" + g.Name + " Error =>" + e.Message);
                }

        }


        static public string BildXMLA(List<ProcessingOnLine> parList)
        {
            StringBuilder script = new StringBuilder();
            script.AppendLine("<Batch xmlns=\"http://schemas.microsoft.com/analysisservices/2003/engine\">");
            script.AppendLine("<Parallel>");
            //"+(parProcessType == ProcessType.ProcessDefault?"": " maxParallel="" )+     "

            foreach (ProcessingOnLine task in parList)
                script.AppendLine(task.GetXMLA());

            script.AppendLine("</Parallel>");
            string varKeyErrorLogFile = (GlobalVar.varKeyErrorLogFile == null ? Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) : GlobalVar.varKeyErrorLogFile.Trim()) +
                    "\\log\\Error_" + DateTime.Now.ToString("yyyyMMdd") + ".log";

            script.AppendLine(@"
 <ErrorConfiguration>"+
      /*<KeyErrorLogFile>" + GlobalVar.varKeyErrorLogFile + @"</KeyErrorLogFile>*/@"
      <KeyErrorAction>ConvertToUnknown</KeyErrorAction>
      <KeyNotFound>IgnoreError</KeyNotFound>
      <NullKeyConvertedToUnknown>IgnoreError</NullKeyConvertedToUnknown>
</ErrorConfiguration>
");
            script.AppendLine("</Batch>");
            return script.ToString();
        }


        /// <summary>
        /// Процес списка 
        /// </summary>
        /// <param name="parList"></param>
        static public void ProcessList(List<ProcessingOnLine> parList)
        {
            if (!RunXMLA(BildXMLA(parList)))
                foreach (ProcessingOnLine task in parList)
                    task.Process();
        }
        /// <summary>
        /// Процес кубів частинами методом XMLA
        /// </summary>
        /// <param name="parParallel"></param>
        /// <param name="parProcessType"></param>
        /// <param name="parInclude"></param>
        static public void ProcessPartXMLA(int parParallel = 0, ProcessType parProcessType = ProcessType.ProcessDefault, bool parInclude = true)
        {
            if (parParallel == 0)
                parParallel = GlobalVar.varMaxParallel;
            if (parParallel == 0)
                return;

            int i = 0;
            foreach (ProcessingOnLine task in GlobalListOnLine.OrderByDescending(x => x.GetYearMonth()))
                if ((parProcessType == ProcessType.ProcessDefault) || (parInclude && parProcessType == task.ProcessingType) || (!parInclude && parProcessType != task.ProcessingType))
                {
                    if (i == 0)
                        LocalListOnLine.Clear();
                    i++;
                    LocalListOnLine.Add(task);
                    if (i == parParallel)
                    {
                        ProcessList(LocalListOnLine);
                        i = 0;                        
                    }
                }
            if (i != 0)
                ProcessList(LocalListOnLine);
        }

        /// <summary>
        /// Запуск XMLA
        /// </summary>
        /// <param name="parXMLA"></param>        
        /// <returns></returns>
        static public bool RunXMLA(string parXMLA)
        {
           if (!(DateTime.Now.Hour >= GlobalVar.varTimeStart) && (DateTime.Now.Hour <= GlobalVar.varTimeEnd))
            {
                Log.log("Час за межами діапазону(" + GlobalVar.varTimeStart.ToString().Trim() + "-" + GlobalVar.varTimeEnd.ToString().Trim() + ") Зараз=>" + DateTime.Now.ToString());
                return true;
            }
            try
            {
                string varRez;
                Log.log(parXMLA);
                XMLACl.Execute(parXMLA, "", out varRez, false, false);
                Log.log("Rez XMLA=>" + varRez);

                if (varRez.IndexOf("<Error") == -1)
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }

        }

        static public DateTime GetDateStartPartition(Partition parPartition)
        {
            TypePeriod varType = GetTypePartition(parPartition.Parent);
            string varSTR = parPartition.Name.Substring(parPartition.Name.Length - ((int)varType < 10 ? 6 : 8)) + ((int) varType < 10 ? "01" : "");
            try
            {
                return new DateTime(Convert.ToInt32(varSTR.Substring(0, 4)), Convert.ToInt32(varSTR.Substring(4, 2)), Convert.ToInt32(varSTR.Substring(6, 2)));
            }
            catch
            {
                return new DateTime(1, 1, 1);
            }
        }

        static public DateTime CountNextPreviousDate(DateTime parStartDate, TypePeriod parType, int parCoef = 1)
        {
            switch (parType)
            {
                case TypePeriod.Month:
                    return parStartDate.AddMonths(parCoef * 1);
                case TypePeriod.Quarter:
                    return parStartDate.AddMonths(parCoef * 3);
                case TypePeriod.Year:
                    return parStartDate.AddYears(parCoef * 1);
                case TypePeriod.Week4:
                    return parStartDate.AddDays(parCoef * 28);
                case TypePeriod.Week:
                    return parStartDate.AddDays(parCoef * 7);
            }
            return new DateTime(9999, 12, 31);
        }


        static public Partition PreviousLastPartition(MeasureGroup parMG)
        {
            TypePeriod varType = GetTypePartition(parMG);
            Partition varCurPar = FindLastPartition(parMG);
            DateTime varDate = CountNextPreviousDate(GetDateStartPartition(varCurPar), varType, -1);
            return PartitionFind(parMG, varDate);
        }
        static public Partition FindLastPartition(MeasureGroup parMG)
        {
            TypePeriod varType = GetTypePartition(parMG);
            DateTime varMax = new DateTime(1, 1, 1), varCur;
            Partition rezPar = null;
            foreach (Partition p in parMG.Partitions)
            {
                varCur = GetDateStartPartition(p);
                if (varCur > varMax)
                {
                    rezPar = p;
                    varMax = varCur;
                }
            }
            return rezPar;
        }

        static public int CountPartition(MeasureGroup parG, bool parProcess = false)
        {
            int i = 0;
            foreach (Partition p in parG.Partitions)
                if (p.State == AnalysisState.Unprocessed || !parProcess)
                    i++;
            return i;
        }

        static public double CountMeasureSize(MeasureGroup parG)
        {
            double s = 0;
            foreach (Partition p in parG.Partitions)
                s += (p.EstimatedSize / (1024 * 1024));
            return s;
        }

        static public double CountCubeSize(Cube parCube)
        {
            double s = 0;
            foreach (MeasureGroup varMG in parCube.MeasureGroups)
                s += CountMeasureSize(varMG);
            return s;
        }

        /// <summary>
        /// Добавляємо в список процесінга необхідні партиції куба.
        /// </summary>
        /// <param name="parCube">Куб</param>
        /// <param name="parStep">Крок процесінга</param>
        static public void AddListProcessPartition(Cube parCube, int parStep)
        {
            if (GlobalVar.varIsArx && strings.GetWordNum(parCube.Description, 2, ";").Substring(0, 5).ToUpper() != "ARX=>")
                return;

            foreach (MeasureGroup g in parCube.MeasureGroups)
            {
                if (IsCurrentProcess(g.Description, parStep))
                {
                    if (GetTypePartition(g) == TypePeriod.NotDefined)
                    {
                        if (g.State == AnalysisState.Unprocessed || g.State == AnalysisState.PartiallyProcessed || parStep != 0)
                            //                 GlobalList.Add(new ProcessingTask(ProcessType.ProcessFull, parCube.Parent.ID, parCube.ID, g.ID));
                            /*if (g.State == AnalysisState.Unprocessed || CountPartition(g) <= 2)
                                GlobalListOnLine.Add(new ProcessingOnLine(GlobalVar.varProcessCube, g));
                            else*/
                                foreach (Partition p in g.Partitions)
                                    if (p.State == AnalysisState.Unprocessed || (p.Description == null ? null : p.Description.Substring(0, 8)) == "current;" || g.Partitions.Count==1)
                                        GlobalListOnLine.Add(new ProcessingOnLine(GlobalVar.varProcessCube, p));
                    }
                    else
                    {
                        Partition varCurrentPartition = null, varPreviousPartition = null;
                        if (parStep != 0)
                        {
                            varCurrentPartition = FindLastPartition(g);
                            //Якщо дані в партіциї
                            if (GetDateStartPartition(varCurrentPartition) > DateTime.Now.AddDays(-(DayProcess(g.Description) )))
                                varPreviousPartition = PreviousLastPartition(g);
                        }
                        foreach (Partition p in g.Partitions)
                        {
                            if ((p == varCurrentPartition) || (p == varPreviousPartition) || (p.State == AnalysisState.Unprocessed) || GetDateStartPartition(p) >= GlobalVar.varDateStartProcess)
                                GlobalListOnLine.Add(new ProcessingOnLine(GlobalVar.varProcessCube, p));
                            else if (parStep != 0) //(p.State == AnalysisState. )
                            {
                                //                                        GlobalListOnLine.Add(new ProcessingOnLine(ProcessType.ProcessIndexes, p));
                            }
                        }
                    }
                }
            }
        }


        static public void QuickUp1(Cube parCube)
        {
            Partition varTemplatePartition;
            foreach (MeasureGroup g in parCube.MeasureGroups)
            {
                if (!g.IsLinked || (g.IsLinked && (StateMeasure(parCube.Parent, g.Source.CubeID, g.Source.MeasureGroupID) == AnalysisState.Processed)))
                    if (GetTypePartition(g) == TypePeriod.NotDefined)
                    {
                        if (g.State == AnalysisState.Unprocessed)
                            if ((varTemplatePartition = g.Partitions.FindByName("template_QuickUp")) != null)

                                GlobalListOnLine.Add(new ProcessingOnLine(GlobalVar.varProcessCube, varTemplatePartition));
                            else
                                foreach (Partition p in g.Partitions)
                                {
                                    GlobalListOnLine.Add(new ProcessingOnLine(GlobalVar.varProcessCube, p));
                                    break;
                                }
                    }
                    else
                    {
                        varTemplatePartition = GetTemplatePartition(g);
                        if (varTemplatePartition != null && varTemplatePartition.State == AnalysisState.Unprocessed)
                            GlobalListOnLine.Add(new ProcessingOnLine(GlobalVar.varProcessCube, varTemplatePartition));
                    }
            }
        }

        static public void QuickUp2(Cube parCube)
        {
            Partition varLastPartition;
            foreach (MeasureGroup g in parCube.MeasureGroups)
            {
                if (!g.IsLinked || (g.IsLinked && (StateMeasure(parCube.Parent, g.Source.CubeID, g.Source.MeasureGroupID) == AnalysisState.Processed)))
                    if (GetTypePartition(g) == 0)
                        foreach (Partition p in g.Partitions)
                            if (p.State == AnalysisState.Unprocessed)
                                GlobalListOnLine.Add(new ProcessingOnLine(GlobalVar.varProcessCube, p));
                            else
                            {
                                varLastPartition = FindLastPartition(g);
                                if (varLastPartition != null && varLastPartition.State == AnalysisState.Unprocessed)
                                    GlobalListOnLine.Add(new ProcessingOnLine(GlobalVar.varProcessCube, varLastPartition));
                            }
            }
        }

        static public void QuickUp(Cube parCube)
        {
            GlobalListOnLine.Clear();
            QuickUp1(parCube);
            ProcessPartXMLA();

            /*GlobalListOnLine.Clear();
            QuickUp2(parCube);
            ProcessPartXMLA();
            */
            GlobalListOnLine.Clear();

        }
        static public void ProcessIndex(Cube parCube)
        {
            foreach (MeasureGroup g in parCube.MeasureGroups)
                GlobalListOnLine.Add(new ProcessingOnLine(ProcessType.ProcessIndexes, g));
        }

        static public AnalysisState StateMeasure(Database varDB, string parCube, string parMeasure)
        {
            return varDB.Cubes[parCube].MeasureGroups[parMeasure].State;

        }


        static public void Process(string parConnectionString, string parDB, string parCube, int parStep = 0, int parMetod = 0)
        {

            Server s = new Server();
            try
            {
                Log.log("Try Connect=>" + parDB);
                s.Connect(parConnectionString);
                Database varDB = s.Databases.FindByName(parDB);

                XMLACl = new XmlaClient();
                XMLACl.Connect(parConnectionString);

                //                AddSlicePartition(varDB); //TMP
                Log.log("Connect=>"+ parDB);

                if (parStep < -9990)
                {
                    foreach (Cube varCube in varDB.Cubes)
                    {
                        Log.log("Куб=>" + varCube.Name + "\t Size=>" + CountCubeSize(varCube).ToString() + " :" + varCube.State.ToString() + " :" + varCube.Description);
                        if (varCube.State != AnalysisState.Processed || parStep == -9998)
                        {
                            foreach (MeasureGroup varMG in varCube.MeasureGroups)
                            {
                                if (varMG.State != AnalysisState.Processed || parStep == -9998)
                                    Log.log(" Група мір =>" + varMG.Name + "\t Size=>" + CountMeasureSize(varMG).ToString() + " :" + varMG.State.ToString() + " :" + (GetTemplatePartition(varMG) == null ? "" : GetTemplatePartition(varMG).Name) + " :" + varMG.Description);
                            }
                        };
                    }
                    return;
                }
                // dimensions processing   || Процесимо діменшини.              
                if (parCube == null && GlobalVar.varIsProcessDimension)
                {
                    Log.log("Dimensions processing" + parDB);
                    foreach (Dimension dim in varDB.Dimensions)
                        GlobalListOnLine.Add(new ProcessingOnLine(SafeProcTypeGet(dim, GlobalVar.varProcessDimension), dim));
                    ProcessPartXMLA();
                    GlobalListOnLine.Clear();
                }

                //Створення нових партіций.
                Log.log("Try CreatePartition");
                if (parCube == null)
                    foreach (Cube varC in varDB.Cubes)
                        CreatePartition(varC);
                else
                    CreatePartition(varDB.Cubes.FindByName(parCube));

                // Швидке підняття куба
                
                if (parCube == null)
                    foreach (Cube varC in varDB.Cubes)
                        QuickUp(varC);
                else
                    QuickUp(varDB.Cubes.FindByName(parCube));

                
                //Процес куба.

                Log.log("Try ListProcess");
                if (parCube == null)
                    foreach (Cube varCube in varDB.Cubes)
                        AddListProcessPartition(varCube, parStep);
                else
                {
                    Cube varCube = varDB.Cubes.FindByName(parCube);
                    AddListProcessPartition(varCube, parStep);
                }
                WaitOracle(parStep);
                Log.log("Try Process");
                ProcessPartXMLA();

                // Всі куби процесимо INDEX (Якщо процесили діменшини)
                //ProcessList(parMetod,parStep,parXMLA);
                if (parCube == null && GlobalVar.varIsProcessDimension)
                {
                    foreach (Cube varCube in varDB.Cubes)
                        ProcessIndex(varCube);
                    ProcessPartXMLA();
                    //                 	ProcessList(parMetod,parStep,1);
                    GlobalListOnLine.Clear();
                }

            }
            finally
            {
                /*                s.Disconnect();
                                clnt.Disconnect(); */

            }
                 
        }


        static public void AddSlicePartition(Database parDB)
        {
            foreach (Cube varC in parDB.Cubes)
                foreach (MeasureGroup g in varC.MeasureGroups)
                {
                    TypePeriod varType = GetTypePartition(g);
                    if (varType != TypePeriod.Month )
                        foreach (Partition p in g.Partitions)
                            if (p.Slice != null && p.Slice.Trim().Length > 0)
                            {
                                p.Slice = "";
                                p.Update();
                            }

                    /*
                                    if((p.Slice == null || p.Slice.Trim().Length==0 ) && p.Name.ToLower().Substring(0,8 )!="template")
                            { 
                                string varS=p.Name.Substring(p.Name.Length-6,6);
                                p.Slice=  "[Час].[Календар].[Місяць].&["+varS+"]";
                                p.Update();
                            }
                    */
                }

        }


    }

}
