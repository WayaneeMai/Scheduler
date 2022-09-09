using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;
using System.Data.OleDb;
using System.Data;

namespace TestingScheduling
{
    class JobAndOperation
    {
        //string[,] setUpTime;
        int job;
        int[] totalOperationsOfJob;
        List<Processing> uph=new List<Processing>();
        List<Processing> processingTime_Job_Operation_MachineType = new List<Processing>();
        List<Processing> setupLists = new List<Processing>();
        List<Processing> setupTime_Job_Operation_MachineType = new List<Processing>();

        /// <summary>
        /// Read the UPH or SetupTime file to set initial schedule environment.
        /// </summary>
        /// <param name="FileAddress"></param>
        /// <param name="ProcessingData">UPH or SetupTime</param>
        /// <param name="SheetName"></param>
        /// <returns></returns>
        public void ReadXlsx(string FileAddress, ProcessingDataType ProcessingData,string SheetName)
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileAddress + ";Extended Properties='Excel 8.0;HDR=Yes;'";
            switch (ProcessingData)
            {
                case ProcessingDataType.UPH:                   
                    using (OleDbConnection connection = new OleDbConnection(con))
                    {
                        connection.Open();
                        OleDbCommand command = new OleDbCommand("select * from" + "[" + SheetName + "$]", connection);
                        using (OleDbDataReader dr = command.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                string partNumber = dr[0].ToString();
                                string step = dr[7].ToString();
                                string testerType = dr[9].ToString();
                                string handlerType = dr[10].ToString();
                                double.TryParse(dr[13].ToString(), out double unitProcessPerHour);
                                SetUPH(partNumber, step, testerType, handlerType, unitProcessPerHour, dr[17].ToString());//to save uph
                            }
                        }
                    }
                    break;
                case ProcessingDataType.SetupTime:
                    using (OleDbConnection connection = new OleDbConnection(con))
                    {
                        connection.Open();
                        OleDbCommand command = new OleDbCommand("select * from" + "[" + SheetName + "$]", connection);
                        using (OleDbDataReader dr = command.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                string customerCode = dr[0].ToString();
                                string package = dr[1].ToString();
                                string device = dr[2].ToString();
                                string testerType = dr[3].ToString().Substring(0, 3);
                                string handlerType = dr[4].ToString().Substring(0, 3);
                                double.TryParse(dr[5].ToString(), out double setupTime);
                                SetSetupTime(customerCode, package, device, testerType, handlerType, setupTime);//to save setup time
                            }
                        }
                    }
                    break;
            }                    
        }
        public void SetUPH(string PartNumber,string Step, string TesterType,string HandlerType,double UnitProcessPerHour,string Default)
        {
            uph.Add(new Processing()
            {
                PartNumber = PartNumber,
                FTStep = Step,
                TesterType = TesterType,
                HandlerType = HandlerType,
                UnitProcessingPerHour = UnitProcessPerHour,
                IsDefault=Default
            });
        }
        public void SetSetupTime(string CustomerCode, string Package, string Device, string TesterType,string HandlerType,double SetupTime)// Define setup time
        {
            setupLists.Add(new Processing()
            {
                CustomerCode = CustomerCode,//新增0501
                Package = Package,//新增0501
                DeviceName = Device,
                TesterType = TesterType,
                HandlerType = HandlerType,
                SetupTime = SetupTime
            });
        }

        public List<Job_Operation_Index> SetFamily_Of_Job(List<Job_Operation_Index> UnrunningLots)
        {
            List<Job_Operation_Index> faimly_of_job = new List<Job_Operation_Index>();
            List<string> devices=new List<string>();// to save all the device for generate family of job
            foreach(Job_Operation_Index lot in UnrunningLots)
            {
                var target = devices.FirstOrDefault(devices => devices == lot.DeviceName);
                if (target == null)
                {
                    devices.Add(lot.DeviceName);
                }
            }
            int familyIndex;
            int job = GetJob();
            for(int i = 1; i <= job; i++)
            {
                var target = UnrunningLots.FirstOrDefault(_ => _.JobIndex == i);
                var taget_family = faimly_of_job.FirstOrDefault(_ => _.DeviceName == target.DeviceName);
                if (taget_family != null)
                {
                    familyIndex = taget_family.FamilyIndex;
                }
                else
                {
                    if (faimly_of_job.Count == 0)
                    {
                        familyIndex = 1;
                    }
                    else
                    {
                        familyIndex = faimly_of_job.Max(_ => _.FamilyIndex) + 1;
                    }                    
                }
                faimly_of_job.Add(new Job_Operation_Index() { JobIndex = i, DeviceName = target.DeviceName, FamilyIndex = familyIndex});
            }        
            return faimly_of_job;
        }

        public void SetJob(List<Job_Operation_Index> UnrunningLots)
        {
            job = UnrunningLots.Max(_ => _.JobIndex);
            totalOperationsOfJob=new int[job];
        }

        public int GetJob()
        {
            return job;
        }

        public void SetJobOperation(List<Job_Operation_Index> UnrunningLots)
        {
            for (int jobIndex = 1; jobIndex < GetJob()+1; jobIndex++)
            {
                var target = UnrunningLots.Where(_ => _.JobIndex == jobIndex).ToList();
                int numberOfOperation = target.Max(_ => _.OperationIndex);
                totalOperationsOfJob[jobIndex-1]=numberOfOperation;// set number of operation for job i
            }
        }

        public int[] GetNumberOfOperation()
        {
            return totalOperationsOfJob;
        }

        public double GetSetupTime(string CustomerCode,string DeviceName, string Package, string TesterType,string HandlerType)
        {
            var target = setupLists.FirstOrDefault(setup => setup.DeviceName == DeviceName && setup.TesterType == TesterType && setup.HandlerType == HandlerType);
            if (target != null)
            {
                return target.SetupTime;
            }
            else//when setup not found
            {
                target = setupLists.FirstOrDefault(setup => setup.CustomerCode == CustomerCode && setup.Package == Package 
                && setup.TesterType == TesterType && setup.HandlerType == HandlerType);
                if (target != null)
                {
                    return target.SetupTime;
                }
                else
                {
                    target = setupLists.FirstOrDefault(setup => setup.CustomerCode == CustomerCode && setup.TesterType == TesterType&& setup.HandlerType==HandlerType&&
                    setup.DeviceName==""&& setup.Package=="");
                    if (target != null)
                    {
                        return target.SetupTime;
                    }
                    else
                    {
                        return -1;
                    }                    
                }                
            }
        }


        /// <summary>
        /// To set setup time of job i, operation k and machine type m.
        /// 1.
        /// </summary>
        /// <param name="UnrunningLots"></param>
        /// <param name="Scheduled"></param>
        /// <param name="MachineType"></param>
        /// <param name="EligibleTesters"></param>
        /// <param name="EligibleHandlers"></param>
        /// <returns></returns>
        public List<Processing> Set_SetupTime_Job_Operation_Machine(List<Job_Operation_Index> UnrunningLots, Scheduler Scheduled, 
            Machine MachineType, List<Resource> EligibleTesters, List<Resource> EligibleHandlers)
        {
            //(1)使用job and operation index取得對應的device                       
            foreach (Job_Operation_Index lot in UnrunningLots)
            {
                int jobIndex = lot.JobIndex;
                int operationIndex = lot.OperationIndex;
                string deviceName = Scheduled.GetDeviceName(UnrunningLots, lot.WorkOrderNumber, lot.FTStep);//To get device
                string customerCode = lot.CustomerCode;//add 0501
                string package = lot.Package;//add 0501
                List<int> machineTypes = new List<int>();               
                machineTypes = MachineType.Get_MachineType_By_Job_Operation(lot.WorkOrderNumber, lot.FTStep);//to get available machineType of job and operation
                foreach (int machineTypeIndex in machineTypes)
                {
                    int testerIndex = MachineType.Get_Resource_Index_By_Machine_Index(Machine.Machine_Type.Tester, machineTypeIndex);
                    int handlerIndex = MachineType.Get_Resource_Index_By_Machine_Index(Machine.Machine_Type.Handler, machineTypeIndex);
                    string testerType = MachineType.Get_Resource_Name_By_Resource_Index(EligibleTesters, testerIndex).Substring(0, 3);//getTesterTypeName
                    string handlerType = MachineType.Get_Resource_Name_By_Resource_Index(EligibleHandlers, handlerIndex).Substring(0, 3);//getHandlerTypeName
                    if (GetSetupTime(customerCode, deviceName, package, testerType, handlerType) != -1)
                    {
                        double setupTime = GetSetupTime(customerCode, deviceName, package, testerType, handlerType)*60;//unit is minutes
                        setupTime_Job_Operation_MachineType.Add(new Processing()
                        {
                            JobIndex = jobIndex,
                            OperationIndex = operationIndex,
                            MachineTypeIndex = machineTypeIndex,
                            SetupTime = setupTime
                        });
                    }                                        
                }
            }
            return setupTime_Job_Operation_MachineType;
        }

        public double GetUPH(string PartNumber, string Step, string TesterType, string HandlerType)
        {
            var target = uph.Where(_ => _.PartNumber == PartNumber && _.FTStep == Step &&
            _.TesterType == TesterType && _.HandlerType == HandlerType).ToList();
            if (target.Count()>1)
            {
                if(target.FirstOrDefault(_ => _.IsDefault == "Y") != null)
                {
                    return target.FirstOrDefault(_ => _.IsDefault == "Y").UnitProcessingPerHour;
                }
                else
                {
                    return target[0].UnitProcessingPerHour;
                } 
            }
            else if (target.Count() ==1)
            {
                return target[0].UnitProcessingPerHour;
            }
            else//when uph not found
            {                
                return -1;
            }
        }

        /// <summary>
        /// To set processing time of job i, operation k and machine type m.
        /// </summary>
        /// <param name="UnrunningLots"></param>
        /// <param name="Scheduled"></param>
        /// <param name="MachineType"></param>
        /// <param name="EligibleTesters"></param>
        /// <param name="EligibleHandlers"></param>
        /// <returns></returns>
        public List<Processing> Set_ProcessingTime_Job_Operation_Machine(List<Job_Operation_Index> UnrunningLots, Scheduler Scheduled, 
            Machine MachineType, List<Resource> EligibleTesters, List<Resource> EligibleHandlers)
        {
            //2.Define ProcessingTime ikm parameter
            //get uph by device, step, tester and handler
            foreach (Job_Operation_Index lot in UnrunningLots)
            {
                int jobIndex = lot.JobIndex;
                int operationIndex = lot.OperationIndex;
                string partNumber = Scheduled.GetPartNumber(UnrunningLots, jobIndex); //GetPartNumberFromUnschedule(jobIndex);//To get device
                string step = Scheduled.GetStep(UnrunningLots, jobIndex, operationIndex);//To get step
                int lotQuantity = Scheduled.GetLotQuantity(UnrunningLots, jobIndex);
                List<int> machineTypes = new List<int>();                
                machineTypes = MachineType.Get_MachineType_By_Job_Operation(lot.WorkOrderNumber, lot.FTStep);//to get available machineType of job and operation
                foreach (int machineTypeIndex in machineTypes)//to set processing to each available machine type
                {
                    int testerIndex = MachineType.Get_Resource_Index_By_Machine_Index(Machine.Machine_Type.Tester, machineTypeIndex);
                    int handlerIndex = MachineType.Get_Resource_Index_By_Machine_Index(Machine.Machine_Type.Handler,machineTypeIndex);
                    string testerType = MachineType.Get_Resource_Name_By_Resource_Index(EligibleTesters, testerIndex).Substring(0, 3);//getTesterTypeName
                    string handlerType = MachineType.Get_Resource_Name_By_Resource_Index(EligibleHandlers, handlerIndex).Substring(0, 3);//getHandlerTypeName
                    double uph = GetUPH(partNumber, step, testerType, handlerType);
                    int temp = lot.Temperature;
                    double heat_cool_time = Get_Heating_Cooling_Time(temp);//heating or cooling time
                    double processingTime = lotQuantity / uph*60 +heat_cool_time;//unit: minutes
                    processingTime_Job_Operation_MachineType.Add(new Processing()
                    {
                        JobIndex = jobIndex,
                        OperationIndex = operationIndex,
                        MachineTypeIndex = machineTypeIndex,
                        ProcessingTime = processingTime
                    });
                }
            }
            return processingTime_Job_Operation_MachineType;            
        }

        public double Get_Heating_Cooling_Time(int Temperature)
        {
            if (Temperature > 28)//meaning hign temperature test
            {
                return 90;
            }else if (Temperature < 22)
            {
                return 120;
            }
            else
            {
                return 0;
            }
        }

        public bool IsUPH_Exist(string PartNumber,string Step,string TesterType,string HandlerType)
        {
            if(uph.Where(_ => _.PartNumber == PartNumber && _.FTStep == Step &&
            _.TesterType == TesterType && _.HandlerType == HandlerType).ToList().Count > 0)//如果存在，則回傳true
            {
                return true;
            }
            else
            {
                return false;   
            }
        }

        public List<Job_Operation_Index> SetJobInformation(List<Job_Operation_Index> UnrunningLots, DateTime ScheduleTime)
        {
            List<Job_Operation_Index> job_operating_data = new List<Job_Operation_Index>();
            foreach(Job_Operation_Index lot in UnrunningLots)
            {
                //setReleaseTime
                //set weight and due date
                var target=job_operating_data.FirstOrDefault(job => job.JobIndex == lot.JobIndex);
                if (target == null)
                {
                    double releaseTime = Set_ReleaseTime_Job(lot.TestReleaseDate, ScheduleTime);
                    double dueTime = Set_DueTime_Job(lot.DueDate, ScheduleTime);
                    job_operating_data.Add(new Job_Operation_Index()
                    {
                        JobIndex = lot.JobIndex,
                        ReleaseTime_of_Job = releaseTime,
                        DueTime_of_Job = dueTime,
                        Weight = lot.Weight
                    });
                }
            }
            return job_operating_data;
        }

        public double Set_ReleaseTime_Job(DateTime ReleaseDate, DateTime ScheduleTime)
        {            
            double releaseTime;
            if (ReleaseDate > ScheduleTime)
            {
                releaseTime = Time_Caculator.CaculateTimeSpan(ScheduleTime, ReleaseDate);
            }
            else
            {
                releaseTime = 0;
            }
            return releaseTime;
        }

        public double Set_DueTime_Job(DateTime DueDate, DateTime ScheduleTime)
        {
            double dueTime;
            if (DueDate > ScheduleTime)
            {
                dueTime = Time_Caculator.CaculateTimeSpan(ScheduleTime, DueDate);
            }
            else
            {
                dueTime = 0;
            }
            return dueTime;
        }

        /*
        public List<Job_Operation_Index> Set_ReleaseTime_Job(List<Job_Operation_Index> UnrunningLots,DateTime ScheduleTime)
        {
            List<Job_Operation_Index> releaseTime_of_job = new List<Job_Operation_Index>();
            foreach(Job_Operation_Index lot in UnrunningLots)
            {
                DateTime releaseDate= lot.TestReleaseDate;
                double releaseTime;
                var target = releaseTime_of_job.FirstOrDefault(_ => _.JobIndex == lot.JobIndex);
                if (target == null)
                {
                    if (releaseDate > ScheduleTime)
                    {
                        releaseTime = Time_Caculator.CaculateTimeSpan(ScheduleTime, releaseDate);
                    }
                    else
                    {
                        releaseTime = 0;
                    }
                    releaseTime_of_job.Add(new Job_Operation_Index() { JobIndex = lot.JobIndex, ReleaseTime_of_Job = releaseTime });
                }                
            }
            return releaseTime_of_job;
        }
        */

        public static double CaculateProcessingTime(int LotQuantity,double UPH)
        {
            double processingTime = LotQuantity / UPH*60;//unit is minutes
            return processingTime;
        }


        public enum ProcessingDataType
        {
            UPH,
            SetupTime
        }
    }
}
