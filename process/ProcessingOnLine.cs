using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.AnalysisServices;
using Microsoft.AnalysisServices.Xmla;
namespace Process
{
    class ProcessingOnLine
    {
        Partition Par = null;
        Dimension Dim = null;
        MeasureGroup Meas = null;

        ErrorConfiguration ErrConf;

        public int GetYearMonth()
        {
            int res = 0;
            if (Par == null || Par.Name.Length < 6)
                return 999999;
            string per = Par.Name.Substring(Par.Name.Length - 6, 6);
            if (int.TryParse(per, out res))
                return res;

            per = per.Substring(2);
            int.TryParse(per, out res);
            return res;

        }
        public ProcessType ProcessingType;
        public ProcessingOnLine(ProcessType parProcessingType, Dimension parDimension)
        {
            ProcessingType = parProcessingType;
            Dim = parDimension;
        }
        public ProcessingOnLine(ProcessType parProcessingType, Partition parPartition)
        {
            ProcessingType = parProcessingType;
            Par = parPartition;
        }
        public ProcessingOnLine(ProcessType parProcessingType, MeasureGroup parMeas)
        {
            ProcessingType = parProcessingType;
            Meas = parMeas;
        }

        public void Process(int parMetod = -1)
        {

            if (parMetod == -1)
                parMetod = GlobalVar.varMetod;
            //Log log =new Log();
            ErrConf = new ErrorConfiguration();
            ErrConf.KeyErrorAction = KeyErrorAction.ConvertToUnknown;
            ErrConf.KeyNotFound = ErrorOption.IgnoreError;
            ErrConf.NullKeyConvertedToUnknown = ErrorOption.IgnoreError;
            if (Dim != null)
                try
                {
                    Log.log("ProcesDimension=>" + Dim.Name);
                    Dim.Process(ProcessingType);
                }
                catch (Exception e)
                {
                    if (ProcessType.ProcessFull != ProcessingType)
                        try
                        {
                            Log.log("try ProcessADD =>" + Dim.Name + "\n" + e.Message);
                            Dim.Process(ProcessType.ProcessAdd);
                        }
                        catch (Exception e2)
                        {
                            Log.log("Error ProcessFull =>" + Dim.Name + "\n" + e2.Message);
                        }
                }
                finally
                {
                    Log.log("End ProcesDimension=>" + Dim.Name);
                };

            if (parMetod == 0)
            {

                if (Par != null)
                    try
                    {
                        Log.log("ProcesPartition (ConvertToUnknown) =>" + Par.ParentCube.Name + "." + Par.Parent.Name + "." + Par.Name);
                        Par.Process(ProcessingType, ErrConf);
                    }
                    catch (Exception e)
                    {
                        if (ProcessingType != ProcessType.ProcessFull)
                        {
                            try
                            {
                                Log.log("try ConvertToUnknown + ProcessFull =>" + Par.Name + "\n" + e.Message);
                                Par.Process(ProcessType.ProcessFull, ErrConf);
                            }
                            catch (Exception e2)
                            {
                                Log.log("Error ConvertToUnknown =>" + Par.Name + "\n" + e2.Message);
                            }
                        }
                    }

                    finally
                    {
                        Log.log("End ProcesPartition=>" + Par.ParentCube.Name + "." + Par.Parent.Name + "." + Par.Name);
                    };
                if (Meas != null)
                    try
                    {
                        Log.log("ProcesPartition (ConvertToUnknown) =>" + Meas.Parent.Name + "." + Meas.Name);
                        Meas.Process(ProcessingType, ErrConf);
                    }
                    catch (Exception e)
                    {
                        if (ProcessingType != ProcessType.ProcessFull)
                        {

                            try
                            {
                                Log.log("try ConvertToUnknown + ProcessFull =>" + Meas.Name + "\n" + e.Message);
                                Meas.Process(ProcessType.ProcessFull, ErrConf);
                            }
                            catch (Exception e2)
                            {
                                Log.log("Error ConvertToUnknown =>" + Meas.Name + "\n" + e2.Message);
                            }
                        }
                    }

                    finally
                    {
                        Log.log("End ProcesPartition=>" + Meas.Parent.Name + "." + Meas.Name);
                    };

            }
            else
            {

                if (Par != null)
                    try
                    {
                        Log.log("ProcesPartition=>" + Par.ParentCube.Name + "." + Par.Parent.Name + "." + Par.Name);
                        Par.Process(ProcessingType);
                    }
                    catch (Exception e)
                    {
                        try
                        {
                            Log.log("try ConvertToUnknown + ProcessFull =>" + Par.Name + "\n" + e.Message);
                            Par.Process(ProcessType.ProcessFull, ErrConf);
                        }
                        catch (Exception e2)
                        {
                            Log.log("Error ConvertToUnknown =>" + Par.Name + "\n" + e2.Message);
                        }
                    }

                    finally
                    {
                        Log.log("End ProcesPartition=>" + Par.ParentCube.Name + "." + Par.Parent.Name + "." + Par.Name);
                    };
                if (Meas != null)
                    try
                    {
                        Log.log("ProcesPartition=>" + Meas.Parent.Name + "." + Meas.Name);
                        Meas.Process(ProcessingType);
                    }
                    catch (Exception e)
                    {
                        try
                        {
                            Log.log("try ConvertToUnknown + ProcessFull =>" + Meas.Name + "\n" + e.Message);
                            Meas.Process(ProcessType.ProcessFull, ErrConf);
                        }
                        catch (Exception e2)
                        {
                            Log.log("Error ConvertToUnknown =>" + Meas.Name + "\n" + e2.Message);
                        }
                    }

                    finally
                    {
                        Log.log("End ProcesPartition=>" + Meas.Parent.Name + "." + Meas.Name);
                    };

            }
        }

        public string GetXMLA()
        {
            string process = null;

            if (ProcessingType == ProcessType.ProcessAdd)
                process = "ProcessAdd";
            else if (ProcessingType == ProcessType.ProcessFull)
                process = "ProcessFull";
            else if (ProcessingType == ProcessType.ProcessIndexes)
                process = "ProcessIndexes";
            else if (ProcessingType == ProcessType.ProcessUpdate)
                process = "ProcessUpdate";
            else if (ProcessingType == ProcessType.ProcessData)
                process = "ProcessData";

            if (Par != null)
                return
       @"   <Process>
      <Object>
        <DatabaseID>" + Par.ParentDatabase.ID + @"</DatabaseID>
        <CubeID>" + Par.ParentCube.ID + @"</CubeID>
        <MeasureGroupID>" + Par.Parent.ID + @"</MeasureGroupID>
        <PartitionID>" + Par.ID + @"</PartitionID>
      </Object>
      <Type>" + process + @"</Type>
     </Process>";

            if (Meas != null)
                return
                    @"   <Process>
      <Object>
        <DatabaseID>" + Meas.ParentDatabase.ID + @"</DatabaseID>
        <CubeID>" + Meas.Parent.ID + @"</CubeID>
        <MeasureGroupID>" + Meas.ID + @"</MeasureGroupID>
      </Object>
      <Type>" + process + @"</Type>
    </Process>";
            if (Dim != null)
                return
    @"    <Process>
      <Object>
        <DatabaseID>" + Dim.Parent.ID + @"</DatabaseID>
        <DimensionID>" + Dim.ID + @"</DimensionID>
      </Object>
      <Type>" + process + @"</Type>
      <WriteBackTableCreation>UseExisting</WriteBackTableCreation>
     </Process>";
            return "";
        }

    }

}
