using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Data;
using System.Data.OleDb;

namespace TestingScheduling
{
    class Machine
    {          
        List<MachineType> machineTypes = new List<MachineType>();//機器型別
        List<MachineType> available_machineType_job_operation = new List<MachineType>();//job i operation k and machine type m
        private List<Resource> eligibleTesters = new List<Resource>();// eligible testers for mo and step
        private List<Resource> eligibleHandlers = new List<Resource>();// eligible handlers for mo and step
        private List<Resource> handlers = new List<Resource>();// to save all available handler
        private List<Resource> testers=new List<Resource>();// to save all available tester

        /// <summary>
        /// Read the machine list file to set initial schedule environment.
        /// </summary>
        /// <param name="FileAddress"></param>
        /// <param name="SheetName"></param>
        public List<Resource> ReadMachineListData(string FileAddress, string SheetName)
        {
            List<Resource> machineLists = new List<Resource>();
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileAddress + ";Extended Properties='Excel 8.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from" + "[" + SheetName + "$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string resourceName = FixResourceName(dr[0].ToString());//the machine name count is 6 characters.
                        string processingLot;
                        if (dr[9].ToString().Length > 7)
                        {
                            processingLot = dr[9].ToString().Substring(0, 7);//to search corresspond part, step and quantity by the lot
                        }
                        else
                        {
                            processingLot = dr[9].ToString();
                        }
                        DateTime.TryParse(dr[11].ToString(), out DateTime trackInTime);
                        DateTime.TryParse(dr[12].ToString(), out DateTime estimateAvailableTime);
                        machineLists.Add(new Resource()
                        {
                            ResourceName = resourceName,
                            ResourceType = dr[1].ToString(),
                            ResourceStatus = dr[4].ToString(),
                            ResourceSubStatus= dr[5].ToString(),//MACHINE SUB STATUS
                            ResourceLocation = dr[8].ToString(),
                            WorkOrderNumber = processingLot,
                            FTStep = dr[10].ToString(),
                            TrackInDate = trackInTime,
                            EstimationAvailableTime = estimateAvailableTime
                        });
                        if (IsConstainHandler(dr[3].ToString().Length))
                        {
                            var target = machineLists.FirstOrDefault(handler => handler.ResourceName == dr[3].ToString());
                            if ( target == null)
                            {
                                machineLists.Add(new Resource()
                                {
                                    ResourceName = dr[3].ToString(),
                                    ResourceType = "Handler_G",//直接設定此為Handler
                                    ResourceStatus = dr[4].ToString(),
                                    ResourceSubStatus = dr[5].ToString(),//MACHINE SUB STATUS
                                    ResourceLocation = dr[8].ToString(),
                                    WorkOrderNumber = processingLot,
                                    FTStep = dr[10].ToString(),
                                    TrackInDate = trackInTime,
                                    EstimationAvailableTime = estimateAvailableTime
                                });
                            }
                            else
                            {
                                if(target.EstimationAvailableTime< estimateAvailableTime)//當handler有重複出現時
                                {
                                    target.ResourceStatus = dr[4].ToString();
                                    target.ResourceSubStatus=dr[5].ToString();
                                    target.WorkOrderNumber=processingLot;
                                    target.FTStep= dr[10].ToString();
                                    target.TrackInDate = trackInTime;
                                    target.EstimationAvailableTime=estimateAvailableTime;
                                }
                            }                          
                        }
                    }
                }
            }
            return machineLists;
        }

        public void ReadEligibleTesterData(string FileAddress,string SheetName,List<Job_Operation_Index> UnRunningLots)
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileAddress + ";Extended Properties='Excel 8.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from" + "[" + SheetName + "$]", connection);
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                adapter.SelectCommand = command;
                DataSet ds = new DataSet();
                adapter.Fill(ds);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string mo = dr[0].ToString().Substring(0, 7);
                        string step = dr[0].ToString().Substring(7);
                        var target = UnRunningLots.FirstOrDefault(_ => _.WorkOrderNumber == mo && _.FTStep == step);
                        if (target != null)
                        {    
                            for (int column_of_tester = 38; column_of_tester < dr.FieldCount; column_of_tester++)//set eligible tester for each job and operation
                            {
                                if (dr[column_of_tester].ToString() == "v")
                                {
                                    string eligibleTester=ds.Tables[0].Columns[column_of_tester].ColumnName;
                                    SetEligibleTester(mo, step, eligibleTester);
                                }
                            }
                        }
                    }
                }
            }
        }

        public void SetEligibleTester(string MO, string Step,string Tester)//set eligible tester
        {
            string resourceType = Tester.Substring(0, 3);
            eligibleTesters.Add(new Resource()
            {
                WorkOrderNumber = MO,
                FTStep = Step,
                ResourceType = resourceType,//Tester Type
                ResourceName = Tester//Tester Name
            });
        }


        /// <summary>
        /// To set the index to all available tester and handler depend on its machine status and location. Assign earliest available time to tester and handler.
        /// Procedure:
        /// 1. To caculate machine available time according to its status.
        /// 2. To set index and available to tester and handler
        /// </summary>
        /// <param name="RunningLots"></param>
        /// <param name="MachineList"></param>
        /// <param name="Scheduled"></param>
        /// <param name="LotInformation"></param>
        /// <param name="CurrentTime"></param>
        public void SetTesterAndHandlerIndex(List<Job_Operation_Index> RunningLots, List<Resource> MachineList, Scheduler Scheduled, JobAndOperation LotInformation, DateTime CurrentTime)//assign index to each tester and handler
        {
            int row = 0;            
            do
            {
                string resourceName;
                if (MachineList[row].ResourceName.Length>6)
                {
                    resourceName = MachineList[row].ResourceName.Substring(0, 6);//機台僅取前六碼
                }
                else
                {
                    resourceName = MachineList[row].ResourceName;
                }
                double machineAvailableTime = 0;
                
                if (IsMachineAvailable(MachineList[row].ResourceLocation, MachineList[row].ResourceStatus, MachineList[row].EstimationAvailableTime)==true)
                {
                    if (MachineList[row].ResourceStatus == "DOWN" || MachineList[row].ResourceStatus == "ENGR" || MachineList[row].ResourceStatus == "PM")//earliest available time will be '預計可作業時間'
                    {
                        if (CurrentTime > MachineList[row].EstimationAvailableTime)//it mean schedule time is greater than machine available time
                            machineAvailableTime = 0;
                        else
                            machineAvailableTime = Time_Caculator.CaculateTimeSpan(CurrentTime, MachineList[row].EstimationAvailableTime);                      
                    }
                    else if (MachineList[row].ResourceStatus == "RUN" && MachineList[row].ResourceSubStatus == "TrackIn" && MachineList[row].WorkOrderNumber != null)//i.e. machine status is "RUN" and lot is not empty
                    {                        
                        string partNumber = Scheduled.GetPartNumber(RunningLots, MachineList[row].WorkOrderNumber);
                        string testerType = Scheduled.GetTesterNumber(RunningLots, MachineList[row].WorkOrderNumber).Substring(0, 3);
                        string handlerType = Scheduled.GetHandlerNumber(RunningLots, MachineList[row].WorkOrderNumber).Substring(0, 3);
                        if (LotInformation.IsUPH_Exist(partNumber, MachineList[row].FTStep, testerType, handlerType) == true)//是否有維護UPH資料，以推算機台可用的時間
                        {
                            DateTime test_plan_out_date_operation = RunningLots.FirstOrDefault(x => x.WorkOrderNumber == MachineList[row].WorkOrderNumber && x.FTStep == MachineList[row].FTStep).CompletionDate;
                            //1125更改進度
                            
                            double test_plan_out_time = Time_Caculator.CaculateTimeSpan(CurrentTime, test_plan_out_date_operation);                            
                            machineAvailableTime = Math.Max(machineAvailableTime, test_plan_out_time);
                            //if (machineAvailableTime < 0)
                                //machineAvailableTime = 0;
                        }
                        else
                        {
                            row++;
                            continue;//remove the machine
                        }
                    }//i.e. machine status is "IDLE", "RUN" && TrackOut/ Abort or "SETUP", the availableTime is 0                   
                    SetTesterOrHandlerIndexNameAvailableTime(MachineList[row].ResourceType, resourceName, machineAvailableTime);
                }
                row++;
            } while (row < MachineList.Count());            
        }

        public string FixResourceName(string ResourceName)
        {
            if (ResourceName.Length > 6)
            {
                return ResourceName.Substring(0, 6);
            }
            else
            {
               return ResourceName;
            }
        }

        public bool IsConstainHandler(int HandlerLength)
        {
            Resource handler = new Resource();
            if (HandlerLength > 4)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// To check if machine is available by location and status
        /// </summary>
        /// <param name="MachineLocation"></param>
        /// <param name="MachineStatus"></param>
        /// <param name="MachineEstimationTime"></param>
        /// <returns></returns>
        public bool IsMachineAvailable(string MachineLocation,string MachineStatus,DateTime MachineEstimationTime)// To check if machine is available by location and status
        {           
            DateTime N = new DateTime(1, 1, 1, 0, 0, 0);
            if ((MachineLocation == "AT6F-FTD" && MachineStatus != "NULL" && MachineStatus != "OTHERS") ||
                    (MachineLocation == "AT7F-FTD" && MachineStatus != "NULL" && MachineStatus != "OTHERS") ||
                    (MachineLocation == "AT9F-FTD" && MachineStatus != "NULL" && MachineStatus != "OTHERS"))
            {
                if((MachineStatus=="DOWN"&&MachineEstimationTime==N)||
                    (MachineStatus=="ENGR" && MachineEstimationTime == N) || 
                    (MachineStatus=="PM" && MachineEstimationTime == N))
                {
                    return false;
                }
                else
                {
                    return true;
                }               
            }
            else
            {
                return false;
            }
        }

        public bool IsMachineTester(string ResourceType)//true mean tester,false mean handler
        {
            if(ResourceType == "FT" || ResourceType == "MT" || ResourceType == "UFT")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void SetTesterOrHandlerIndexNameAvailableTime(string ResourceType, string ResourceName,double AvailableTime)//set tester and handler Index
        {
            int resourceIndex;
            if (IsMachineTester(ResourceType)==true && IsTesterDuplicated(ResourceName) == false)//Tester
            {                
                if (testers.Count == 0)
                {
                    resourceIndex = 1;
                }
                else
                {
                    resourceIndex = testers.Max(_ => _.ResourceIndex) + 1;
                }                
                testers.Add(new Resource() { ResourceIndex = resourceIndex, ResourceName = ResourceName, AvailableTime = AvailableTime});
            }
            else if (IsMachineTester(ResourceType) == false && IsHandlerDuplicated(ResourceName) == false)//Handler
            {               
                if (handlers.Count == 0)
                {
                    resourceIndex = 1;
                }
                else
                {
                    resourceIndex = handlers.Max(_ => _.ResourceIndex) + 1;
                }                
                handlers.Add(new Resource() { ResourceIndex = resourceIndex, ResourceName = ResourceName, AvailableTime = AvailableTime });
            }         
        }


        public List<Resource> Get_Elibible_Testers_List()
        {
            return eligibleTesters;
        }

        public List<Resource> Get_Eligible_Handlers_List()
        {
            return eligibleHandlers;
        }

        public string GetTesterName(int TesterIndex)
        {
            var target = testers.FirstOrDefault(_ => _.ResourceIndex == TesterIndex);
            return target.ResourceName;
        }

        public int GetHandlerIndex(string ResourceName)
        {
            var target = handlers.FirstOrDefault(_ => _.ResourceName == ResourceName);
            return target.ResourceIndex;
        }

        public string GetHandlerName(int HandlerIndex)
        {
            var target=handlers.FirstOrDefault(_ => _.ResourceIndex == HandlerIndex);
            return target.ResourceName;
        }
        /// <summary>
        /// To get correspond tester or handler index by machine type index.
        /// </summary>
        /// <param name="Tester_Handler">Input 'Tester' or 'Handler'</param>
        /// <param name="MachineIndex"></param>
        /// <returns></returns>
        public int Get_Resource_Index_By_Machine_Index(Machine_Type Tester_Handler, int MachineIndex)
        {
            var target = machineTypes.FirstOrDefault(_ => _.MachineTypeIndex == MachineIndex);
            switch (Tester_Handler)
            {
                case Machine_Type.Tester:
                    return target.TesterIndex;
                case Machine_Type.Handler:
                    return target.HandlerIndex;
                default:
                    return -1;
            }      
        }



        public bool IsTesterDuplicated(string ResourceName)//檢查是否有重複
        {
            var target = testers.FirstOrDefault(_ => _.ResourceName == ResourceName);
            if (target != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool IsHandlerDuplicated(string ResourceName)//檢查是否有重複
        {
            var target = handlers.FirstOrDefault(_ => _.ResourceName == ResourceName);
            if (target != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public List<string> GetCorrespondHandlerMachine(string HandlerType)
        {
            List<string> handlerMachine = new List<string>();
            List<Resource> target = handlers.Where(_ => _.ResourceName.Substring(0,3) == HandlerType).ToList();
            foreach (Resource resource in target)
            {
                handlerMachine.Add(resource.ResourceName);
            }
            return handlerMachine;
        }

        /// <summary>
        /// Remove the tester which is not going to use and assign resource index to eligible tester sets. e.g. not available tester in eligible tester sets will be removed.
        /// </summary>
        public void RemoveNotAvailableTesterMachine()//移除不可用的機台與更新機台標號
        {
            SetResourceIndexToEligibleTesterSets(); // assign resource index to eligible tester sets           
            RemoveNotAvailableTesterFromEligibleTesterSets();// remove not available tester from eligible tester sets
        }

        public void RemoveNotAvailableTesterFromEligibleTesterSets()
        {
            foreach (Resource tester in eligibleTesters.ToList())
            {
                var target = eligibleTesters.FirstOrDefault(_ => _.ResourceIndex == 0);//resourceIndex will be 0, if eligible tester is not available
                if (target != null)
                {
                    eligibleTesters.Remove(target);
                }
            }
        }

        

        public void SetEligibleHandler(List<Job_Operation_Index> UnrunningLot, Accessory Accessory)
        {
            foreach (Job_Operation_Index wip in UnrunningLot)
            {
                List<string> handlerTypes = new List<string>();
                handlerTypes = Accessory.GetHandlerType(wip.CustomerCode, wip.DeviceName, wip.FTStep);
                foreach (string handlerType in handlerTypes)
                {
                    List<string> handlerMachineList = GetCorrespondHandlerMachine(handlerType);
                    foreach (string handlerMachine in handlerMachineList)
                    {
                        int handlerIndex = GetHandlerIndex(handlerMachine);
                        eligibleHandlers.Add(new Resource()
                        {
                            WorkOrderNumber = wip.WorkOrderNumber,
                            FTStep = wip.FTStep,
                            ResourceName = handlerMachine,
                            ResourceType = handlerType, 
                            ResourceIndex = handlerIndex });
                    }
                }
            }
        }

        public void SetResourceIndexToEligibleTesterSets()// to set resource index to eligible tester when tester is availablity
        {
            foreach (Resource tester in testers)
            {
                string availableTester = tester.ResourceName;//可用的機台
                int startIndex = 0;
                int totalCount = eligibleTesters.Where(_ => _.ResourceName == availableTester).ToList().Count;
                int updateTimes = 1;//實際更新次數
                if (totalCount > 0)//如果有找到，則更新resourceIndex
                {
                    while (updateTimes <= totalCount)// assign resourceIndexToElgigbleTester
                    {
                        int index = eligibleTesters.FindIndex(startIndex, _ => _.ResourceName == availableTester);
                        int testerIndex = tester.ResourceIndex;
                        eligibleTesters[index].ResourceIndex = testerIndex;
                        startIndex = index + 1;
                        updateTimes++;
                    }
                }
            }
        }

        /// <summary>
        /// To generate machine type as input of algorithm. machine type is a combination of tester and handler. The generating procedure as follows:
        /// 1.set the tester and handler index
        /// 2-1.Remove the tester which is not going to use and assign resource index to eligible tester sets.
        /// 2-2.define eligible handler and assign resource index to operation
        /// 3.generate machine type, tester, handler and assign machine type index
        /// 4.define available tester and handler
        /// 5.generate tester and handler's availability
        /// 
        /// </summary>
        /// <param name="MachineLists"></param>
        /// <param name="UnrunningLots"></param>
        /// <param name="LotInformation"></param>
        /// <param name="CurrentTime"></param>
        public void GenerateMachineType(List<Resource> MachineLists, List<Job_Operation_Index> UnrunningLots, List<Job_Operation_Index> RunningLots,
            Scheduler Scheduled, JobAndOperation LotInformation, DateTime CurrentTime,Accessory Accessory)//同時增加機台可作業的工件與作業標號
        {
            SetTesterAndHandlerIndex(RunningLots, MachineLists, Scheduled, LotInformation,CurrentTime);//set the index to all available tester and handler         
            RemoveNotAvailableTesterMachine();// Remove the tester which is not going to use and assign resource index to eligible tester sets.         
            SetEligibleHandler(UnrunningLots,Accessory);// set eligible handler                  
            foreach (Resource handler in eligibleHandlers)//Generate machine type and set machine type index to those elibible testers and eligible handlers.   
            {
                List<Resource> targer_testers = eligibleTesters.Where( eligibleTester => eligibleTester.WorkOrderNumber == handler.WorkOrderNumber 
                && eligibleTester.FTStep == handler.FTStep).ToList();// given mo and step, find eligible tester.
                foreach(Resource tester in targer_testers)
                {
                    SetMachineTypeIndex(tester.ResourceIndex, handler.ResourceIndex,tester.AvailableTime,handler.AvailableTime);//set machine type index
                }
            }
            Set_Available_MachineType_Job_Operation();//設定機台可執行的工件與作業
        }

        public void SetMachineTypeIndex(int TesterIndex, int HandlerIndex,double TesterAvailableTime,double HandlerAvailableTime)
        {
            int machine_type_index;
            var target_MachineType = machineTypes.FirstOrDefault(machineType => machineType.TesterIndex == TesterIndex
                        && machineType.HandlerIndex == HandlerIndex);//如果有不同的組合
            if (target_MachineType == null)
            {
                if (machineTypes.Count == 0)
                {
                    machine_type_index = 1;
                }
                else
                {
                    machine_type_index = machineTypes.Max(machineType => machineType.MachineTypeIndex) + 1;
                }                
                double availableTime = Math.Max(TesterAvailableTime, HandlerAvailableTime);
                machineTypes.Add(new MachineType() { MachineTypeIndex = machine_type_index, TesterIndex = TesterIndex, HandlerIndex = HandlerIndex, AvailableTime=availableTime });
            }
            else
            {
                return;
            }
        }

        public void Set_Available_MachineType_Job_Operation()
        {            
            foreach(Resource handler in eligibleHandlers)
            {
                List<Resource> target_eligibleTesters = eligibleTesters.Where(eligibleTester => eligibleTester.WorkOrderNumber == handler.WorkOrderNumber 
                && eligibleTester.FTStep == handler.FTStep).ToList();
                foreach(Resource tester in target_eligibleTesters)
                {
                    var target_machineType = machineTypes.FirstOrDefault(machineType => machineType.TesterIndex == tester.ResourceIndex && machineType.HandlerIndex == handler.ResourceIndex);
                    var duplicate = available_machineType_job_operation.FirstOrDefault(_=>_.MachineTypeIndex== target_machineType.MachineTypeIndex 
                    && _.WorkOrderNumber==tester.WorkOrderNumber&&_.FTStep==tester.FTStep);
                    if (duplicate == null) 
                    {
                        available_machineType_job_operation.Add(new MachineType()
                        {
                            MachineTypeIndex = target_machineType.MachineTypeIndex,
                            WorkOrderNumber = tester.WorkOrderNumber,
                            FTStep = tester.FTStep,
                        });
                    }                      
                }          
            }
        }

        public List<MachineType> Get_available_Machine_Job_Operation()
        {
            return available_machineType_job_operation;
        }

        public List<int> Get_MachineType_By_Job_Operation(string WorkOrder,string FTstep)//to get all available machine type by job and operation
        {
            List<int> availableMachineTypes = new List<int>();
            var target = available_machineType_job_operation.Where(_ => _.WorkOrderNumber == WorkOrder && _.FTStep == FTstep).ToList();
            foreach (MachineType machineType in target)
            {
                availableMachineTypes.Add(machineType.MachineTypeIndex);
            }
            return availableMachineTypes;
        }

        public void Remove_Machine_Type_of_Job_Operation(string WorkOrder, string FTstep,int MachineType)
        {
            var index = available_machineType_job_operation.FindIndex(_ => _.WorkOrderNumber == WorkOrder
            && _.FTStep == FTstep && _.MachineTypeIndex == MachineType);
            available_machineType_job_operation.RemoveAt(index);
        }
    
        /// <summary>
        /// To get correspond tester or handler name from Tester or Handler lists by resource index. 
        /// </summary>
        /// <param name="ResourceList"></param>
        /// <param name="ResourceIndex"></param>
        /// <returns></returns>
        public string Get_Resource_Name_By_Resource_Index(List<Resource> ResourceList,int ResourceIndex)
        {
            var target = ResourceList.FirstOrDefault(resource => resource.ResourceIndex == ResourceIndex);
            return target.ResourceName;             
        }

        /// <summary>
        /// Return tester or handler and its available time.
        /// </summary>
        /// <param name="TesterOrHandler">Input 'Tester' or 'Handler'</param>
        /// <returns></returns>
        public List<Resource> Get_Avaiable_Testers_or_Handlers(Machine_Type machine_type)
        {
            switch (machine_type)
            {
                case Machine_Type.Tester:
                    return testers;
                case Machine_Type.Handler:
                    return handlers;
                default:
                    return null;
            }
        }

        public List<MachineType> Get_MachineType_Tester_Handler()
        {
            return machineTypes;
        }

        public enum Machine_Type
        {
            Tester,
            Handler
        }
    }
}
