using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;
using System.Data;
using System.Data.OleDb;

namespace TestingScheduling
{
    class Scheduler
    {
        List<Job_Operation_Index> runningLots = new List<Job_Operation_Index>();
        List<Job_Operation_Index> unRunningLots = new List<Job_Operation_Index>();

        public void ReadData_Makespan(string FileAddress, string JobListSheetName,string LotListSheetName, List<string> CustomerCode,string File_SmartTable_DueDate,string SheetName_smartTable)
        {
            List<Job_Operation_Index> jobLists = new List<Job_Operation_Index>();
            jobLists = ReadJobList(FileAddress, JobListSheetName, CustomerCode);
            List<Job_Operation_Index> lotLists = ReadLotList(FileAddress, LotListSheetName, CustomerCode);
            List<SmartTable_DueDate> smartTable_DueDates = ReadSmartTable(File_SmartTable_DueDate, SheetName_smartTable);
          
            foreach (Job_Operation_Index jobList in jobLists)//define weight and due date to each job in job lists
            {
                var target_lotList = lotLists.FirstOrDefault(lotList => lotList.WorkOrderNumber == jobList.WorkOrderNumber);
                double processDay = GetProcessTime_SmartTable(smartTable_DueDates, jobList.PartNumber.Substring(0, 1), jobList.CustomerCode, jobList.Package,
                    jobList.Package, jobList.TestTotalStep, jobList.FTStep, jobList.DeviceName);//to caculate estimate due time
                jobList.DueDate = SetDueDate(target_lotList.DueDate, jobList.AssyInputDate, jobList.TestReleaseDate, jobList.PartNumber, target_lotList.AMO, processDay);//set due date
                jobList.Weight = target_lotList.Weight;//set weught
            }
            SetScheduleAndUnscheduleJobAndOperation(jobLists);
        }


        public void ReadData_Tardiness(string FileAddress, string JobListSheetName,string LotListSheetName, List<string> CustomerCode, string File_SmartTable_DueDate, string SheetName_smartTable)
        {
            List<Job_Operation_Index> jobLists = new List<Job_Operation_Index>();
            jobLists = ReadJobList(FileAddress, JobListSheetName, CustomerCode);
            List<Job_Operation_Index> lotLists = new List<Job_Operation_Index>();
            lotLists = ReadLotList(FileAddress, LotListSheetName, CustomerCode);
            List<SmartTable_DueDate> smartTable_DueDates = new List<SmartTable_DueDate>();
            smartTable_DueDates = ReadSmartTable(File_SmartTable_DueDate, SheetName_smartTable);

            //define weight and due to each job in job lists
            foreach(Job_Operation_Index jobList in jobLists)
            {
                var target_lotList = lotLists.FirstOrDefault(lotList => lotList.WorkOrderNumber == jobList.WorkOrderNumber);
                double processDay = GetProcessTime_SmartTable(smartTable_DueDates, jobList.PartNumber.Substring(0, 1), jobList.CustomerCode, jobList.Package, 
                    jobList.Package, jobList.TestTotalStep, jobList.FTStep, jobList.DeviceName);
                jobList.DueDate = SetDueDate(target_lotList.DueDate,jobList.AssyInputDate,jobList.TestReleaseDate,jobList.PartNumber, target_lotList.AMO,processDay);//set due date
                jobList.Weight = target_lotList.Weight;//set weught
            }
            SetScheduleAndUnscheduleJobAndOperation(jobLists);
        }

        public List<Job_Operation_Index> ReadJobList(string FileAddress,string SheetName,List<string> CustomerCode)
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileAddress + ";Extended Properties='Excel 8.0;HDR=Yes;'";
            List<Job_Operation_Index> allLots = new List<Job_Operation_Index>();
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from" + "[" + SheetName + "$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string customerCode = dr[8].ToString();
                        var target = CustomerCode.FirstOrDefault(CustomerCode => CustomerCode == customerCode);
                        if (target != null)// To add the lot when customer code meeting the requirement.
                        {
                            int.TryParse(dr[21].ToString(), out int test_total_step);
                            DateTime.TryParse(dr[13].ToString(), out DateTime assyInputDate);
                            DateTime.TryParse(dr[14].ToString(), out DateTime assyPlanoutDate);
                            int.TryParse(dr[20].ToString(), out int lotQuantity);
                            DateTime.TryParse(dr[25].ToString(), out DateTime trackIndate);//navingo track in date 
                            int temperature;
                            if (dr[26].ToString().Length < 1)
                                temperature = 25;
                            else
                                int.TryParse(dr[26].ToString(), out temperature);                            
                            DateTime.TryParse(dr[27].ToString(), out DateTime testInputDate);//i.e. TEST Input Date
                            DateTime.TryParse(dr[28].ToString(), out DateTime testPlanoutDate);//test plan out date                         
                            allLots.Add(new Job_Operation_Index()
                            {
                                CustomerCode = customerCode,
                                Tester = dr[3].ToString(),
                                Handler = dr[6].ToString(),
                                WorkOrderNumber = dr[0].ToString().Substring(0, 7),
                                FTStep = dr[0].ToString().Substring(7),
                                TestTotalStep =test_total_step,
                                PartNumber = dr[10].ToString(),//i.e, Tpart e.g., TA9PA048XXX
                                LotNumber = dr[11].ToString(),
                                Package = dr[18].ToString(),
                                DeviceName = dr[19].ToString(),
                                LotQuantity = lotQuantity,
                                Temperature = temperature,
                                AssyInputDate = assyInputDate,
                                AssyPlanoutDate = assyPlanoutDate,
                                TrackInDate = trackIndate,
                                TestReleaseDate = testInputDate,
                                CompletionDate = testPlanoutDate,
                                CurrentStatus = dr[16].ToString()
                            });
                        }
                    }
                }
            }
            return allLots;
        }

        public List<Job_Operation_Index> ReadLotList(string FileAddress, string SheetName, List<string> CustomerCode)
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileAddress + ";Extended Properties='Excel 8.0;HDR=Yes;'";
            List<Job_Operation_Index> lotList = new List<Job_Operation_Index>();
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from" + "[" + SheetName + "$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string customerCode = dr[1].ToString();
                        DateTime.TryParse(dr[43].ToString(), out DateTime r_tn_date1);
                        int weight;
                        if (dr[44] != null)
                        {
                            if (dr[44].ToString().Count() == 0)
                            {
                                weight = 20;//預設值
                            }
                            else
                            {
                                int.TryParse(dr[44].ToString(), out weight);
                            }
                        }
                        else
                        {
                            weight = 20;//預設值
                        }
                        
                        var target = CustomerCode.FirstOrDefault(CustomerCode => CustomerCode == customerCode);
                        if (target != null)// To add the lot when customer code meeting the requirement.
                        {
                            lotList.Add(new Job_Operation_Index()
                            {
                                CustomerCode = customerCode,
                                AMO= dr[2].ToString(),
                                WorkOrderNumber = dr[37].ToString(),
                                DueDate=r_tn_date1,
                                Weight=weight
                            });
                        }
                    }
                }
            }
            return lotList;
        }

        public List<SmartTable_DueDate> ReadSmartTable(string FileAddress, string SheetName)
        {
            List<SmartTable_DueDate> smartTable = new List<SmartTable_DueDate>(); ;
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileAddress + ";Extended Properties='Excel 8.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from" + "[" + SheetName + "$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        double.TryParse(dr[0].ToString(), out double day);
                        int.TryParse(dr[5].ToString(), out int test_total_step);
                        smartTable.Add(new SmartTable_DueDate()
                        {
                            ProcessingDay = day,
                            PartType = dr[1].ToString(),
                            CustCode = dr[2].ToString(),
                            PKGType = dr[3].ToString(),
                            PKG = dr[4].ToString(),
                            TestTotalStep = test_total_step,
                            Step = dr[6].ToString(),
                            DeviceName = dr[7].ToString()
                        });
                    }               
                }
            }
            return smartTable;
        }


        public double GetProcessTime_SmartTable(List<SmartTable_DueDate> SmartTables, string PartType,string CustCode,string PKGType,string PKG,int TotalStep,string Step,string DeviceName)
        {
            var target = SmartTables.FirstOrDefault(table => table.DeviceName == DeviceName && table.PartType == PartType);
            if (target != null)
            {
                return target.ProcessingDay;
            }
            else
            {
                target = SmartTables.FirstOrDefault(table => table.PartType == PartType && table.Step == Step);
                if (target != null)
                {
                    return target.ProcessingDay;
                }
                else
                {
                    target = SmartTables.FirstOrDefault(table => table.PartType == PartType && table.CustCode == CustCode && table.TestTotalStep == TotalStep);
                    if (target != null)
                    {
                        return target.ProcessingDay;
                    }
                    else
                    {
                        return -1;
                    }
                }
            }
        }



        public class SmartTable_DueDate{
            public double ProcessingDay { get; set; }
            public string PartType { get; set; }
            public string CustCode { get; set; }
            public string PKGType { get; set; }
            public string PKG { get; set; }
            public int TestTotalStep { get; set; }
            public string Step { get; set; }
            public string DeviceName { get; set; }
        }

        public DateTime SetDueDate(DateTime R_TN_date1, DateTime AssyInputDate,DateTime TestInputDate, string Part,string AMO,double ProcessingTime)
        {
            DateTime N = new DateTime(1, 1, 1, 0, 0, 0);
            if (R_TN_date1 != N)
            {
                return R_TN_date1;
            }
            else
            {
                DateTime dueDate=new DateTime();
                switch (Part.Substring(0, 1))
                {
                    case "A":
                        dueDate = AssyInputDate.AddMinutes(ProcessingTime);
                        break;
                    case "P":
                        if (TestInputDate != N)
                        {
                            dueDate = TestInputDate.AddMinutes(ProcessingTime);
                        }
                        else
                        {
                            dueDate = AssyInputDate.AddMinutes(ProcessingTime);
                        }
                        break;
                    case "T":
                        if (AMO != "")
                        {
                            dueDate = AssyInputDate.AddMinutes(ProcessingTime);
                        }
                        else
                        {
                            dueDate = TestInputDate.AddMinutes(ProcessingTime);
                        }
                        break;
                    default:
                        dueDate = Chromosome.Data.currentTime;
                        break;
                }
                return dueDate;
            }
        }


        


        /// <summary>
        /// Read the wip data to set initial schedule environment. The Wip will be divided into two groups, including running lot(固定既有排程) and unrunning lot(待排程).
        /// </summary>
        /// <param name="AllLot"></param>
        public void SetScheduleAndUnscheduleJobAndOperation(List<Job_Operation_Index> AllLot)
        {
            int row = 0;
            do
            {
                string currentRowMoAndStep =AllLot[row].WorkOrderNumber+AllLot[row].FTStep;
                if ((AllLot[row].CurrentStatus != "Hold" && IsWipMoAndStepDuplicated(AllLot,currentRowMoAndStep) ==false ))
                    //represent not duplicated, not "holding the product" and customer code meet the requirement
                {
                    string mo = AllLot[row].WorkOrderNumber;
                    DateTime releaseDate = AllLot[row].TestReleaseDate;
                    if (AllLot[row].CurrentStatus == "Run") //if status is Run, it will not need to assign index.
                    {
                        SetSchedule(AllLot[row]);//set the schedule for processing job and operation                        
                        int nextRowIndex=row;//to update next operation of releasing date
                        if (row+1 < AllLot.Count())
                            nextRowIndex = row + 1;//to update next operation of releasing date
                        while (IsThereNextOperationOfTheJob(mo, AllLot[nextRowIndex].WorkOrderNumber) == true)// 
                        {
                            AllLot[nextRowIndex].TestReleaseDate = AllLot[row].CompletionDate;//update release date of next operation
                            nextRowIndex++;
                            if (nextRowIndex == AllLot.Count())
                                break;
                        }                       
                    }
                    else
                    {
                        if (AllLot[row].CurrentStatus == "BS")//BS's release will be test input date+4hours
                            releaseDate = releaseDate.AddHours(4);
                        else if (AllLot[row].CurrentStatus == "Assy")
                            releaseDate = AllLot[row].AssyPlanoutDate;
                        
                        AllLot[row].TestReleaseDate= releaseDate;
                        SetUnSchedule(AllLot[row]);
                    }
                }
                row++;                
            } while (row < AllLot.Count());
        }


        /// <summary>
        /// To validate if any missing in uph or setup time for unrunning lots. To ensure correctness of unrunning lot, validate each lot's UPH and setup time information.
        /// To validate if accessory is enough to process
        /// Especially to investigate if any missing. Procedure as follow:
        /// 1.To validate every lot in unrunning lots
        /// 2.remove the available machine of job and operation when missing uph or setup time
        /// 3.Remove the job when there isn't available machine to process
        /// </summary>
        /// <param name="MachineType"></param>
        /// <param name="EligibleTesters">eligible and available testers</param>
        /// <param name="EligibleHandlers">eligible and available handlers</param>
        /// <param name="LotInformation"></param>
        public void ValidateUnrunningLot(ref Machine MachineType,List<Resource> EligibleTesters,List<Resource> EligibleHandlers,JobAndOperation LotInformation,Accessory Acc)
        {
            foreach(Job_Operation_Index validateLot in unRunningLots.ToList())
            {
                if(ValidateUPHAndSetupTimeData(validateLot, EligibleTesters, EligibleHandlers, ref MachineType, LotInformation) < 1)
                {                   
                    continue;
                }
                else
                {
                    //remove any operation which its lot(job) is missing uph or setup data from unrunning lot.
                    var target = runningLots.FirstOrDefault(x => x.WorkOrderNumber == validateLot.WorkOrderNumber && x.CurrentStatus.Contains("Maintain"));
                    if (target != null)
                    {
                        RemoveLotFromUnschedule(validateLot.WorkOrderNumber, validateLot.FTStep);
                    }                                                        
                    VaildateAccessoryQuantity(validateLot, Acc);//to check if there is enougn accessory to process.
                }                              
            }
        }

        public int ValidateUPHAndSetupTimeData(Job_Operation_Index ValidateLot, List<Resource> EligibleTesters, List<Resource> EligibleHandlers,
            ref Machine Machines, JobAndOperation LotInformation)
        {
            string partNumber = GetPartNumber(unRunningLots, ValidateLot.WorkOrderNumber); // GetPartNumberFromUnschedule(workOrder);
            string deviceName = GetDeviceName(unRunningLots, ValidateLot.WorkOrderNumber, ValidateLot.FTStep);
            List<int> machineTypes = new List<int>(Machines.Get_MachineType_By_Job_Operation(ValidateLot.WorkOrderNumber, ValidateLot.FTStep)); //to get available machineType of job and operation
            int setupDataFound = 0;
            int uphDataFound = 0;
            if (partNumber != "NA" || deviceName != "NA")
            {
                string missing_testerType = "NA";
                string missing_handlerType = "NA";
                foreach (int machineTypeIndex in machineTypes)
                {
                    int testerIndex = Machines.Get_Resource_Index_By_Machine_Index(Machine.Machine_Type.Tester, machineTypeIndex);
                    int handlerIndex = Machines.Get_Resource_Index_By_Machine_Index(Machine.Machine_Type.Handler, machineTypeIndex);
                    string testerType = Machines.Get_Resource_Name_By_Resource_Index(EligibleTesters, testerIndex).Substring(0, 3);//getTesterTypeName
                    string handlerType = Machines.Get_Resource_Name_By_Resource_Index(EligibleHandlers, handlerIndex).Substring(0, 3);//getHandlerTypeName
                    if (LotInformation.GetSetupTime(ValidateLot.CustomerCode, deviceName, ValidateLot.Package, testerType, handlerType) < 0 
                        && LotInformation.GetUPH(partNumber, ValidateLot.FTStep, testerType, handlerType) < 0)//if missing setup and uph data
                    {
                        Machines.Remove_Machine_Type_of_Job_Operation(ValidateLot.WorkOrderNumber, ValidateLot.FTStep, machineTypeIndex);
                        missing_handlerType = handlerType;
                        missing_testerType = testerType;
                    }
                    else if(LotInformation.GetSetupTime(ValidateLot.CustomerCode, deviceName, ValidateLot.Package, testerType, handlerType) < 0)//if missing setup data
                    {
                        Machines.Remove_Machine_Type_of_Job_Operation(ValidateLot.WorkOrderNumber, ValidateLot.FTStep, machineTypeIndex);
                        uphDataFound = uphDataFound + 1;
                        missing_handlerType = handlerType;
                        missing_testerType = testerType;
                    }
                    else if(LotInformation.GetUPH(partNumber, ValidateLot.FTStep, testerType, handlerType) < 0)//if missing uph data
                    {
                        Machines.Remove_Machine_Type_of_Job_Operation(ValidateLot.WorkOrderNumber, ValidateLot.FTStep, machineTypeIndex);
                        setupDataFound = setupDataFound + 1;
                        missing_handlerType = handlerType;
                        missing_testerType = testerType;
                    }
                    else//not missing uph or setup data
                    {
                        setupDataFound = setupDataFound + 1;
                        uphDataFound = uphDataFound + 1;
                    }
                }
                Output_Validate_Result(setupDataFound, uphDataFound, missing_handlerType, missing_testerType, ValidateLot);
            }
            return Math.Min(setupDataFound,uphDataFound);
        }

        public void Output_Validate_Result(int SetupDataFoundCount,int UphDataFoundCount, string Missing_handlerType, string Missing_testerType, Job_Operation_Index ValidateLot)
        {
            if (SetupDataFoundCount < 1 && UphDataFoundCount < 1)//if there is not a machine type to process for job i, it will be remove from the wip list.
            {
                RemoveLotFromUnschedule(ValidateLot.WorkOrderNumber, ValidateLot.FTStep);//remove the mo from unrunning lot
                if (Missing_handlerType == "NA" || Missing_testerType == "NA")
                    ValidateLot.CurrentStatus = "Maintain UPH and setup time";
                else
                    ValidateLot.CurrentStatus = "Maintain UPH and setup time for tester type " + Missing_testerType + " and handler type " + Missing_handlerType;
                SetSchedule(ValidateLot);//set the lot in runningLots sets and mark with notes
            }
            else if (SetupDataFoundCount < 1 && UphDataFoundCount > 0)
            {
                RemoveLotFromUnschedule(ValidateLot.WorkOrderNumber, ValidateLot.FTStep);//remove the mo from unrunning lot
                if (Missing_handlerType == "NA" || Missing_testerType == "NA")
                    ValidateLot.CurrentStatus = "Maintain setup time";
                else
                    ValidateLot.CurrentStatus = "Maintain setup time for tester type " + Missing_testerType + " and handler type " + Missing_handlerType;
                SetSchedule(ValidateLot);//set the lot in runningLots sets and mark with notes
            }
            else if (UphDataFoundCount < 1 && SetupDataFoundCount > 0)
            {
                RemoveLotFromUnschedule(ValidateLot.WorkOrderNumber, ValidateLot.FTStep);//remove the mo from unrunning lot
                if (Missing_handlerType == "NA" || Missing_testerType == "NA")
                    ValidateLot.CurrentStatus = "Maintain UPH";
                else
                    ValidateLot.CurrentStatus = "Maintain UPH for tester type " + Missing_testerType + " and handler type " + Missing_handlerType;
                SetSchedule(ValidateLot);//set the lot in runningLots sets and mark with notes
            }
        }

        public void VaildateAccessoryQuantity(Job_Operation_Index ValidationLot,Accessory Acc)
        {
            var require_accessory = Acc.Get_Require_Accessory_Job_Operation().Where(x => x.WorkOrderNumber == ValidationLot.WorkOrderNumber && x.FTStep == ValidationLot.FTStep).ToList();
            List<AvailableAccessory> target_sec_accessory = new List<AvailableAccessory>();
            while (require_accessory.Count() != 0)//to check all required accessory to the validation lot.
            {
                var target_pri_accessory = Acc.GetAvailableAccessories().Where(availableAccessory => availableAccessory.Accessory_Type_Index == require_accessory[0].ResourceIndex).ToList();
                int available_sec_accessory = 0;
                if (Acc.IsSecondaryAccessoryExist(require_accessory[0].ResourceType, require_accessory[0].ResourceName))
                {
                    string sec_acc_type = Acc.GetSecondaryAccessoryType(require_accessory[0].ResourceType, require_accessory[0].ResourceName);//get secondary type and name
                    string sec_acc_name = Acc.GetSecondaryAccessoryName(require_accessory[0].ResourceType, require_accessory[0].ResourceName);
                    int sec_accIndex = Acc.Get_Accessory_Index(sec_acc_name, sec_acc_type);
                    target_sec_accessory=Acc.GetAvailableAccessories().Where(_=>_.Accessory_Type_Index==sec_accIndex).ToList();
                    available_sec_accessory = target_sec_accessory.Count();
                }            
                if (target_pri_accessory.Count()+ available_sec_accessory < require_accessory[0].RequireQuantity)//the accessory is not enougn to process
                {
                    string missing_acc_type = Copy.DeepClone(require_accessory[0].ResourceType);
                    string missing_acc_name = Copy.DeepClone(require_accessory[0].ResourceName);
                    RemoveLotFromUnschedule(ValidationLot.WorkOrderNumber);//remove the mo from unrunning lot
                    ValidationLot.CurrentStatus = "Accessory type "+ missing_acc_type + " accessory "+ missing_acc_name + " available quantity is not enough to process.";
                    SetSchedule(ValidationLot);//set the lot in runningLots sets and mark with notes
                    break;
                }
                require_accessory.Remove(require_accessory[0]);
            }
        }

        public List<Job_Operation_Index> Set_Job_Operation_Index()//to set index to unrunning lots job and operation as the input of algorithm
        {
            int jobIndex = 1;
            int operationIndex = 1;
            int rowLength = unRunningLots.ToList().Count;
            for(int row = 0; row < rowLength-1; row++)
            {
                unRunningLots[row].JobIndex= jobIndex;
                unRunningLots[row].OperationIndex = operationIndex;
                if (unRunningLots[row].WorkOrderNumber == unRunningLots[row + 1].WorkOrderNumber)
                {
                    operationIndex++;
                }
                else
                {
                    jobIndex++;
                    operationIndex = 1;
                }
            }
            unRunningLots[rowLength-1].JobIndex = jobIndex;
            unRunningLots[rowLength-1].OperationIndex = operationIndex;
            return unRunningLots;
        }

        public List<Job_Operation_Index> GetAllUnRunningLot()
        {
            return unRunningLots;
        }

        public List<Job_Operation_Index> GetAllRunngingLot()
        {
            return runningLots;
        }

        public bool IsWipMoAndStepDuplicated(List<Job_Operation_Index> SearchingLists, string CurrentRowMoAndStep)
        {
            string mo = CurrentRowMoAndStep.Substring(0, 7);
            string step = CurrentRowMoAndStep.Substring(7);
            int foundTimes = SearchingLists.Where(_ => _.WorkOrderNumber == mo && _.FTStep == step).ToList().Count();
            if (foundTimes > 1)
            {
                return true;//represent duplicated
            }
            else
            {
                return false;
            }
        }

        public bool IsThereNextOperationOfTheJob(string CurrentRowMO,string NextRowMO)
        {
            if ( CurrentRowMO== NextRowMO)//the release date of next operation will be estimate completion date
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public List<Job_Operation_Index> SetSchedule(Job_Operation_Index WipMO)
        {
            runningLots.Add(new Job_Operation_Index()
            {
                Tester = WipMO.Tester,
                Handler = WipMO.Handler,
                CustomerCode=WipMO.CustomerCode,
                WorkOrderNumber = WipMO.WorkOrderNumber,
                FTStep = WipMO.FTStep,
                LotNumber = WipMO.LotNumber,
                DeviceName = WipMO.DeviceName,
                Package = WipMO.Package,
                LotQuantity = WipMO.LotQuantity,
                Temperature = WipMO.Temperature,
                TrackInDate = WipMO.TrackInDate,
                TrackInDate_Processing=WipMO.TrackInDate_Processing,
                CompletionDate = WipMO.CompletionDate,
                PartNumber = WipMO.PartNumber,
                CurrentStatus=WipMO.CurrentStatus,
                DueDate=WipMO.DueDate,
                TestReleaseDate= WipMO.TestReleaseDate,
            });
            return runningLots;
        }

        public void SetUnSchedule(Job_Operation_Index WipMO)
        {
            unRunningLots.Add(new Job_Operation_Index()
            {
                //JobIndex = Job,
                //OperationIndex = Operation,
                CustomerCode=WipMO.CustomerCode,//add 0502
                WorkOrderNumber = WipMO.WorkOrderNumber,    
                FTStep = WipMO.FTStep,
                PartNumber = WipMO.PartNumber,
                LotNumber = WipMO.LotNumber,
                Package = WipMO.Package,
                DeviceName = WipMO.DeviceName,
                LotQuantity = WipMO.LotQuantity,
                Temperature = WipMO.Temperature,
                TestReleaseDate = WipMO.TestReleaseDate,
                CurrentStatus = WipMO.CurrentStatus,
                Weight = WipMO.Weight,
                DueDate = WipMO.DueDate
            });
            //return unRunningLot;
        } 

        public void RemoveLotFromUnschedule(string MO,string FTStep)
        {
            var totalCount=unRunningLots.Where(lot=>lot.WorkOrderNumber==MO&&lot.FTStep==FTStep).Count();
            for(int removeTime = 0; removeTime < totalCount; removeTime++)
            {
                int index = unRunningLots.FindIndex(_ => _.WorkOrderNumber == MO&&_.FTStep==FTStep);
                unRunningLots.RemoveAt(index);
            }
        }

        public void RemoveLotFromUnschedule(string MO)
        {
            var totalCount = unRunningLots.Where(lot => lot.WorkOrderNumber == MO).Count();
            for (int removeTime = 0; removeTime < totalCount; removeTime++)
            {
                int index = unRunningLots.FindIndex(_ => _.WorkOrderNumber == MO);
                unRunningLots.RemoveAt(index);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ScheduleResult"></param>
        /// <param name="FileName">output file name</param>
        /// <param name="Lots"></param>
        public void OutPutScheduleResult(List<Job_Operation_Index> ScheduleResult,string FileName,JobAndOperation Lots)
        {
            using (StreamWriter scheduled = new StreamWriter(FileName+".csv",false, System.Text.Encoding.UTF8)) //output scheduleResult
            {
                scheduled.WriteLine("Tester,Handler,CustID,T_MO,Step,Part,LOT#,DEVICE NAME,PKG,QTY,SetupTime,UPH,Processing Time( unit: hour),Temp.,TEST Input Date,"+
                    "Estimate TrackIn Time,TEST Estimate Finish Date, Due Date, Status");
                foreach (Job_Operation_Index operation in ScheduleResult)
                {
                    double uph;
                    double processingTime;
                    double setupTime;
                    if (operation.CurrentStatus == "Run")//run時，setup 欄位是空白
                    {
                        processingTime = Math.Round(Time_Caculator.CaculateTimeSpan(operation.CompletionDate, operation.TrackInDate)/60,1);
                        setupTime = Lots.GetSetupTime(operation.CustomerCode, operation.DeviceName, operation.Package, operation.Tester.Substring(0, 3), operation.Handler.Substring(0, 3));
                        uph = Lots.GetUPH(operation.PartNumber, operation.FTStep, operation.Tester.Substring(0, 3), operation.Handler.Substring(0, 3));
                        if (uph == -1)
                            uph = Math.Round(operation.LotQuantity / processingTime,1);
                        if (setupTime == -1)
                            setupTime = 0;
                        scheduled.WriteLine(operation.Tester + "," + operation.Handler + "," + operation.CustomerCode + "," + operation.WorkOrderNumber + "," + operation.FTStep + ","
                        + operation.PartNumber + "," + operation.LotNumber + "," + operation.DeviceName + "," + operation.Package + "," + operation.LotQuantity + ", ," + uph + ","
                        + processingTime + "," + operation.Temperature + "," + operation.TestReleaseDate.ToString("yyyy/MM/dd HH:mm") + "," + operation.TrackInDate.ToString("yyyy/MM/dd HH:mm") + ","
                        + operation.CompletionDate.ToString("yyyy/MM/dd HH:mm") + "," + operation.DueDate.ToString("yyyy/MM/dd HH:mm") + "," + operation.CurrentStatus);
                    }
                    else
                    {
                        if (operation.Tester != null || operation.Handler != null)
                        {
                            uph = Lots.GetUPH(operation.PartNumber, operation.FTStep, operation.Tester.Substring(0, 3), operation.Handler.Substring(0, 3));
                            processingTime = Math.Round(operation.LotQuantity / uph + Lots.Get_Heating_Cooling_Time(operation.Temperature) / 60, 1);
                            setupTime = Lots.GetSetupTime(operation.CustomerCode, operation.DeviceName, operation.Package, operation.Tester.Substring(0, 3), operation.Handler.Substring(0, 3));
                        }
                        else
                        {
                            setupTime = -1;
                            uph = -1;
                            processingTime = -1;
                        }
                        scheduled.WriteLine(operation.Tester + "," + operation.Handler + "," + operation.CustomerCode + "," + operation.WorkOrderNumber + "," + operation.FTStep + ","
                        + operation.PartNumber + "," + operation.LotNumber + "," + operation.DeviceName + "," + operation.Package + "," + operation.LotQuantity + "," + setupTime + "," + uph + ","
                        + processingTime + "," + operation.Temperature + "," + operation.TestReleaseDate.ToString("yyyy/MM/dd HH:mm") + "," + operation.TrackInDate.ToString("yyyy/MM/dd HH:mm") + ","
                        + operation.CompletionDate.ToString("yyyy/MM/dd HH:mm") + "," + operation.DueDate.ToString("yyyy/MM/dd HH:mm") + "," + operation.CurrentStatus);
                    }                   
                }
            }
        }



        public void OutPutScheduleResult_ForResearchOnly(List<Job_Operation_Index> ScheduleResult, string FileName, JobAndOperation Lots)
        {
            using (StreamWriter scheduled = new StreamWriter(FileName + ".csv", false, System.Text.Encoding.UTF8)) //output scheduleResult
            {
                scheduled.WriteLine("Tester,Handler,CustID,T_MO,Step,Part,LOT#,DEVICE NAME,PKG,QTY,SetupTime,UPH,Processing Time( unit: hour),Temp.,TEST Input Date," +
                    "Estimate TrackIn Time,TEST_Start_Processing,TEST Estimate Finish Date, Due Date, Status");
                foreach (Job_Operation_Index a in ScheduleResult)
                {
                    double uph;
                    double processingTime;
                    double setupTime;
                    if (a.CurrentStatus == "Run")
                    {
                        processingTime = Math.Round(Time_Caculator.CaculateTimeSpan(a.CompletionDate, a.TrackInDate) / 60, 1);
                        setupTime = Lots.GetSetupTime(a.CustomerCode, a.DeviceName, a.Package, a.Tester.Substring(0, 3), a.Handler.Substring(0, 3));
                        uph = Lots.GetUPH(a.PartNumber, a.FTStep, a.Tester.Substring(0, 3), a.Handler.Substring(0, 3));
                        if (uph == -1)
                        {
                            uph = Math.Round(a.LotQuantity / processingTime, 1);
                        }
                    }
                    else if ((a.Tester != "" && a.Tester != null) || (a.Handler != "" && a.Handler != null))
                    {
                        uph = Lots.GetUPH(a.PartNumber, a.FTStep, a.Tester.Substring(0, 3), a.Handler.Substring(0, 3));
                        processingTime = Math.Round(a.LotQuantity / uph + Lots.Get_Heating_Cooling_Time(a.Temperature) / 60, 1);
                        setupTime = Lots.GetSetupTime(a.CustomerCode, a.DeviceName, a.Package, a.Tester.Substring(0, 3), a.Handler.Substring(0, 3));
                    }
                    else
                    {
                        setupTime = -1;
                        uph = -1;
                        processingTime = -1;
                    }
                    scheduled.WriteLine(a.Tester + "," + a.Handler + "," + a.CustomerCode + "," + a.WorkOrderNumber + "," + a.FTStep + "," + a.PartNumber + "," + a.LotNumber
                        + "," + a.DeviceName + "," + a.Package + "," + a.LotQuantity + "," + setupTime + "," + uph + "," + processingTime + "," + a.Temperature + "," + a.TestReleaseDate.ToString("yyyy/MM/dd HH:mm")
                        + "," + a.TrackInDate.ToString("yyyy/MM/dd HH:mm") + ","+a.TrackInDate_Processing.ToString("yyyy/MM/dd HH:mm") +"," + 
                        a.CompletionDate.ToString("yyyy/MM/dd HH:mm") + "," + a.DueDate.ToString("yyyy/MM/dd HH:mm") + "," + a.CurrentStatus);
                }
            }
        }
   

        public string GetPartNumber(List<Job_Operation_Index> Lots, string ProcessingLot)
        {
            var target = Lots.FirstOrDefault(_ => _.WorkOrderNumber == ProcessingLot);
            if (target != null)
            {
                return target.PartNumber;
            }
            else
            {
                return "NA";                
            }                      
        }

        public string GetPartNumber(List<Job_Operation_Index> Lots, int JobIndex)
        {
            var target = Lots.FirstOrDefault(_ => _.JobIndex == JobIndex);
            if (target != null)
            {
                return target.PartNumber;
            }
            else
            {
                Console.WriteLine(JobIndex + "不存在於job list中");
                return "NA";
            }
        }

        public string GetDeviceName(List<Job_Operation_Index> Lots, string MO, string Step)
        {
            var target = Lots.FirstOrDefault(_ => _.WorkOrderNumber == MO && _.FTStep == Step);
            if (target != null)
            {
                return target.DeviceName;
            }
            else
            {
                return "NA";
            }
        }

        public string GetStep(List<Job_Operation_Index> Lots, int JobIndex, int OperationIndex)
        {
            var target = Lots.FirstOrDefault(_ => _.JobIndex == JobIndex && _.OperationIndex == OperationIndex);
            return target.FTStep;
        }

        public int GetTemperature(List<Job_Operation_Index> Lots, string MO,string Step)
        {
            var target = Lots.FirstOrDefault(_ => _.WorkOrderNumber == MO && _.FTStep == Step);
            return target.Temperature;
        }
        

        public int GetLotQuantity(List<Job_Operation_Index> Lots, string ProcessingLot)
        {
            var target = Lots.FirstOrDefault(_ => _.WorkOrderNumber == ProcessingLot);
            
            if (target != null)
            {
                return target.LotQuantity;
            }
            else
            {
                return 0;
            }
        }

        public int GetLotQuantity(List<Job_Operation_Index> Lots, int JobIndex)
        {
            var target= Lots.FirstOrDefault(_ => _.JobIndex == JobIndex);
            return target.LotQuantity;
        }

        public string GetTesterNumber(List<Job_Operation_Index> Lots, string ProcessingLot)
        {
            var target = Lots.FirstOrDefault(_ => _.WorkOrderNumber == ProcessingLot);
            if (target != null)
            {
                return target.Tester;
            }
            else
            {
                return "job list 無該筆作業";
            }
            
        }

        public string GetHandlerNumber(List<Job_Operation_Index> Lots, string ProcessingLot)
        {
            var target = Lots.FirstOrDefault(_ => _.WorkOrderNumber == ProcessingLot);
            if (target != null)
            {
                return target.Handler;
            }
            else
            {
                return "job list 無該筆作業";
            }
        }

    }
}
