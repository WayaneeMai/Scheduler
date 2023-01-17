using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.IO;
using System.Data;
using System.Data.OleDb;

namespace TestingScheduling
{
    class Accessory
    {
        private List<Resource> required_accessories = new List<Resource>();//to save all accessory. including index, type and name.
        private List<Resource> require_Accessory_Job_Operation = new List<Resource>();// to set required accessory of job and operation
        private List<Resource> processing_accessory=new List<Resource>();//record on processing accessory and its next available time
        private List<AvailableAccessory> available_quantity_accessories = new List<AvailableAccessory>();//available quantity of accessory
        private List<SecondardResource> secondary_resources = new List<SecondardResource>();
        public List<Resource> eligibleHandlers;

        public void SetAllAccessoryList(string AccessoryName,string AccessoryType) //To set required accessories and assign accessory index
        {
            if (IsDuplicatedAccessoryExist(required_accessories, AccessoryName, AccessoryType) == false)
            {
                int accessoryIndex = required_accessories.Count() + 1;
                required_accessories.Add(new Resource() { ResourceIndex = accessoryIndex, ResourceType = AccessoryType, ResourceName = AccessoryName });                
            }
        }
        public int Get_Accessory_Index(string AccessoryName,string AccessoryType)
        {
            var target = required_accessories.FirstOrDefault(requireAccessory => requireAccessory.ResourceType == AccessoryType && requireAccessory.ResourceName == AccessoryName);
            return target.ResourceIndex;
        }

        //including accessoryIndexType, accessoryId, availableTIme
        /// <summary>
        /// To set available accessory. 不須處理InLine狀態的配件，該配件可用GetAvailableTime_Accessory_for_runningLot()的方式計算排程時可用的數量
        ///(1)Read the data and set available quantity of accessory when required accessory exist. 
        ///(2)to set avaiable quantity of accessory for 'Instock'.
        /// </summary>
        /// <param name="FileAddress"></param>
        /// <param name="SheetName"></param>
        public void ReadAvailableAccessoryData(string FileAddress,string SheetName,AccessoryType AccessoryCategory)            
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileAddress + ";Extended Properties='Excel 8.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from" + "[" + SheetName + "$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string accessoryStatus;
                        switch (AccessoryCategory)
                        {
                            case AccessoryType.Accessory_ACC:
                                accessoryStatus=dr[11].ToString();
                                break;
                            case AccessoryType.Accessory_Part:
                                accessoryStatus=dr[12].ToString();
                                break ;
                            default:
                                accessoryStatus = "InputError";
                                break ;
                        }
                        string accType = dr[1].ToString();
                        string accName = dr[2].ToString();
                        if (IsExistInSecondary(accType, accName) == true)
                        {
                            string primaryAccType = GetPrimaryResourceType(accType, accName);
                            string primaryAccName = GetPrimaryResourceName(accType, accName);
                            if ((IsAccessoryBeUsed(accType, accName) == true||IsAccessoryBeUsed(primaryAccType,primaryAccName)) && accessoryStatus == "InStock") //配件有存在allAccessoryList
                            {
                                SetAvailableAccessory(accType, accName);
                            }
                        }
                        else
                        {
                            if (IsAccessoryBeUsed(accType, accName) == true && accessoryStatus == "InStock") //配件有存在allAccessoryList
                                SetAvailableAccessory(accType, accName);
                        }
                    }
                }
            }
        }

        public void SetAvailableAccessory(string AccType, string AccName)
        {
            var target = available_quantity_accessories.Where(available_accessory => available_accessory.Accessory_Type_Index == Get_Accessory_Index(AccName, AccType)).ToList();//0524
            if (target.Count() != 0)//revise 0526
            {
                int accessory_index = target.Max(x => x.AvailableAccessoryIndex) + 1;
                available_quantity_accessories.Add(new AvailableAccessory()
                {
                    AccessoryType=AccType,
                    AccessoryName=AccName,
                    Accessory_Type_Index = Get_Accessory_Index(AccName, AccType),
                    AvailableAccessoryIndex = accessory_index,
                    Time = 0
                });
            }
            else
            {
                available_quantity_accessories.Add(new AvailableAccessory()
                {
                    Accessory_Type_Index = Get_Accessory_Index(AccName, AccType),
                    AvailableAccessoryIndex = 1,
                    Time = 0
                });
            }
        }


        public void SetAvailableAccessory_for_processingLot()
        {
            foreach(Resource accessory in processing_accessory)
            {
                var target_available_accessory = available_quantity_accessories.Where(x => x.Accessory_Type_Index == accessory.ResourceIndex).ToList();
                int accessory_index;
                if (target_available_accessory.Count() > 0)
                {
                    accessory_index = target_available_accessory.Max(x => x.AvailableAccessoryIndex) + 1;
                }
                else
                {
                    accessory_index=1;
                }
                for (int i = 0; i < accessory.ResourceQuantity; i++)
                {
                    available_quantity_accessories.Add(new AvailableAccessory()
                    {
                        AccessoryType = accessory.ResourceType,
                        AccessoryName = accessory.ResourceName,
                        Accessory_Type_Index = accessory.ResourceIndex,
                        AvailableAccessoryIndex = accessory_index,
                        Time = accessory.AvailableTime
                    });
                    accessory_index ++;
                }               
            }
        }


        public List<AvailableAccessory> GetAvailableAccessories()
        {
            return available_quantity_accessories;
        }

        public bool IsAccessoryBeUsed(string AccessoryType,string AccessoryName)//is accessory be used?
        {
            var target = required_accessories.FirstOrDefault(_ => _.ResourceType == AccessoryType && _.ResourceName == AccessoryName);
            if (target == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool IsExistInSecondary(string AccessoryType,string AccessoryName)
        {
            var target = secondary_resources.FirstOrDefault(_ => _.SecondaryResourceType == AccessoryType && _.SecondaryResourceName == AccessoryName);
            if (target == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool IsSecondaryAccessoryExist(string AccessoryType, string AccessoryName)
        {
            var target = secondary_resources.FirstOrDefault(_ => _.PrimaryResourceType == AccessoryType && _.PrimaryResourceName == AccessoryName);
            if (target == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public string GetSecondaryAccessoryType(string AccessoryType, string AccessoryName)
        {
            var target = secondary_resources.FirstOrDefault(_ => _.PrimaryResourceType == AccessoryType && _.PrimaryResourceName == AccessoryName);
            return target.SecondaryResourceType;
        }

        public string GetSecondaryAccessoryName(string AccessoryType, string AccessoryName)
        {
            var target = secondary_resources.FirstOrDefault(_ => _.PrimaryResourceType == AccessoryType && _.PrimaryResourceName == AccessoryName);
            return target.SecondaryResourceName;
        }

        public void ReadRequireAccessoryData(string FileAddress, string SheetName, List<Job_Operation_Index> UnrunningLots, List<Job_Operation_Index> RunningLots,DateTime ScheduleTime)
        {
            List<Resource> requireAccessoryLists=new List<Resource>();
            List<Job_Operation_Index> allLots=new List<Job_Operation_Index>(RunningLots);//to save umruunning lots and running lots
            foreach (Job_Operation_Index unrunning_lot in UnrunningLots)
            {
                allLots.Add(new Job_Operation_Index()
                {
                    DeviceName=unrunning_lot.DeviceName,
                    WorkOrderNumber=unrunning_lot.WorkOrderNumber,
                    FTStep=unrunning_lot.FTStep,
                    CurrentStatus= unrunning_lot.CurrentStatus
                });
            }
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileAddress + ";Extended Properties='Excel 8.0;HDR=Yes;'";
            eligibleHandlers = new List<Resource>();
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from" + "[" + SheetName + "$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        requireAccessoryLists = SetRequireAccessory(dr, requireAccessoryLists, allLots);//set require accessory to operations
                        SetEligibleHandlers(dr);//to set eligible handler type to operation from require acc list
                    }                                             
                }
            }            
            foreach(Job_Operation_Index unrunning_lot in UnrunningLots)//To set require accessory of job and operation for unrunning lots.
            {
                Set_Require_Accessory(unrunning_lot, requireAccessoryLists);
            }        
            foreach(Job_Operation_Index running_lot in RunningLots)//To set next available time for using accessory of running lots.
            {
                var target_require_accessory = requireAccessoryLists.Where(x => x.DeviceName == running_lot.DeviceName && x.FTStep == running_lot.FTStep).ToList();                
                if (target_require_accessory !=null)//it mean running lot and unrunning lot have same requirement of accessory
                {
                    foreach (Resource accessory in target_require_accessory)//update next available to the accessory
                    {                
                        Set_Next_Available_Time_To_IsUsing_Accessory(running_lot, accessory, ScheduleTime);
                    }
                }
            }
        }


        public void SetEligibleHandlers(OleDbDataReader Accessories)
        {
            string device = FixDeviceNameDataErrorInAccessoryFile(Accessories[1].ToString());
            Resource target_handler = eligibleHandlers.FirstOrDefault(handler => handler.DeviceName == device
                  && handler.FTStep == Accessories[4].ToString() && handler.ResourceType == Accessories[3].ToString() && handler.CustomerCode == "I" + Accessories[0].ToString());//Avoid duplicated
            if (target_handler == null)
            {
                eligibleHandlers.Add(new Resource()
                {
                    CustomerCode = "I"+Accessories[0].ToString(),
                    DeviceName = FixDeviceNameDataErrorInAccessoryFile(Accessories[1].ToString()),
                    FTStep = Accessories[4].ToString(),
                    ResourceType = Accessories[3].ToString(),
                }) ;
            }
        }

        public List<string> GetHandlerType(string CustCode, string Device, string Step)
        {
            List<string> handlerTypes = new List<string>();
            var target = eligibleHandlers.Where(_ => _.DeviceName == Device && _.FTStep == Step&&_.CustomerCode==CustCode).ToList();
            foreach (Resource handler in target)
            {
                handlerTypes.Add(handler.ResourceType);
            }
            return handlerTypes;
        }


        public List<Resource> SetRequireAccessory(OleDbDataReader RequireAccessories, List<Resource> requireAccessoryLists, List<Job_Operation_Index> AllLots)
        {
            string deviceName_accessory = FixDeviceNameDataErrorInAccessoryFile(RequireAccessories[1].ToString());// to fix the format of device name error.
            string step_accessory = RequireAccessories[4].ToString();
            Job_Operation_Index target_allLots = AllLots.FirstOrDefault(lot => lot.DeviceName == deviceName_accessory && lot.FTStep == step_accessory);
            if (target_allLots != null)//i.e This accessory have demand to the job and operation in running and unrunning lots.
            {
                string accessoryType = RequireAccessories[26].ToString();
                string accessoryName = RequireAccessories[27].ToString();
                int.TryParse(RequireAccessories[28].ToString(), out int requireAccessoryQuantity);
                SetAllAccessoryList(accessoryName, accessoryType); //To set required accessories and assign accessory index
                Resource target_accessory = requireAccessoryLists.FirstOrDefault(requireAccessory => requireAccessory.DeviceName == deviceName_accessory
                  && requireAccessory.FTStep == step_accessory && requireAccessory.ResourceType == accessoryType && requireAccessory.ResourceName == accessoryName);//Avoid duplicated
                if (target_accessory == null)
                {
                    requireAccessoryLists.Add(new Resource()
                    {
                        DeviceName = deviceName_accessory,
                        FTStep = step_accessory,
                        ResourceType = accessoryType,
                        ResourceName = accessoryName,
                        RequireQuantity = requireAccessoryQuantity
                    });
                }
                for (int loadBoardColumn = 9; loadBoardColumn < 21; loadBoardColumn = loadBoardColumn + 4)//load board
                {
                    if (RequireAccessories[loadBoardColumn].ToString() == "" || RequireAccessories[loadBoardColumn].ToString() == "NULL")
                        break;
                    string loadBoardType = RequireAccessories[loadBoardColumn].ToString();
                    string loadBoardName = RequireAccessories[loadBoardColumn + 1].ToString();
                    int.TryParse(RequireAccessories[loadBoardColumn + 3].ToString(), out int loadBoardQuantity);
                    SetAllAccessoryList(loadBoardName, loadBoardType);
                   
                    target_accessory = requireAccessoryLists.FirstOrDefault(requireAccessory => requireAccessory.DeviceName == deviceName_accessory
                    && requireAccessory.FTStep == step_accessory && requireAccessory.ResourceType == loadBoardType && requireAccessory.ResourceName == loadBoardName);
                    if (target_accessory == null)//Avoid duplicated
                    {
                        requireAccessoryLists.Add(new Resource()
                        {
                            DeviceName = deviceName_accessory,
                            FTStep = step_accessory,
                            ResourceType = loadBoardType,
                            ResourceName = loadBoardName,
                            RequireQuantity = loadBoardQuantity
                        });
                    }
                }
            }
            return requireAccessoryLists;
        }


        /// <summary>
        /// To set require accessory to job and operation. To create a Require Accessory Sets of job and operation.
        /// </summary>
        /// <param name="UnrunningLots"></param>
        /// <param name="RequireAccessories"></param>
        public void Set_Require_Accessory(Job_Operation_Index UnrunningLots, List<Resource> RequireAccessories)
        {
            List<Resource> target_require_accessory = RequireAccessories.Where(requiredAccessory => requiredAccessory.DeviceName == UnrunningLots.DeviceName
            && requiredAccessory.FTStep == UnrunningLots.FTStep).ToList();

            foreach (Resource accessory in target_require_accessory)
            {
                int accessoryIndex = Get_Accessory_Index(accessory.ResourceName, accessory.ResourceType);
                var target = require_Accessory_Job_Operation.FirstOrDefault(require_accessory_ik => require_accessory_ik.WorkOrderNumber == UnrunningLots.WorkOrderNumber 
                && require_accessory_ik.FTStep == UnrunningLots.FTStep && require_accessory_ik.ResourceIndex == accessoryIndex);
                if (target == null)//to add require accessory to job and operation
                {
                    require_Accessory_Job_Operation.Add(new Resource()
                    {
                        WorkOrderNumber = UnrunningLots.WorkOrderNumber,
                        FTStep=UnrunningLots.FTStep,
                        ResourceType = accessory.ResourceType,
                        ResourceName = accessory.ResourceName,
                        ResourceIndex = accessoryIndex,//accessory_type_index
                        RequireQuantity = accessory.RequireQuantity
                    });
                }
            }      
        }

        public void Set_Require_Accessory_Job_Operation(List<Job_Operation_Index> UnrunningLots)
        {
            foreach(Job_Operation_Index lot in UnrunningLots)//to set job and operation index to each accessory
            {
                List<Resource> target_require_accessories = require_Accessory_Job_Operation.Where(_ => _.WorkOrderNumber == lot.WorkOrderNumber&&_.FTStep==lot.FTStep).ToList();
                foreach(Resource require_accessory in target_require_accessories)
                {
                    require_accessory.JobIndex = lot.JobIndex;
                    require_accessory.OperationIndex=lot.OperationIndex;
                }
            }
        }


        public List<Resource> Get_Require_Accessory_Job_Operation()
        {
            return require_Accessory_Job_Operation;
        }

        public void Set_Next_Available_Time_To_IsUsing_Accessory(Job_Operation_Index RunningLots, Resource Accessory, DateTime ScheduleTime)
        {
            double availableTime=0;
            if (ScheduleTime < RunningLots.CompletionDate)//current schedule > estimation completion time, the accessory will be available
                availableTime = Time_Caculator.CaculateTimeSpan(RunningLots.CompletionDate, ScheduleTime);
            int accessoryIdex = Get_Accessory_Index(Accessory.ResourceName, Accessory.ResourceType);

            processing_accessory.Add(new Resource()
            {
                ResourceIndex = accessoryIdex,
                ResourceQuantity=Accessory.RequireQuantity,
                ResourceType=Accessory.ResourceType,
                ResourceName=Accessory.ResourceName,
                AvailableTime=availableTime
            });
        }

        /// <summary>
        /// To get using accessory in running lots. To update accessory quantity and next available time
        /// </summary>
        /// <returns></returns>
        public List<Resource> Get_Using_Accessory()
        {
            return processing_accessory;
        }

        public string FixDeviceNameDataErrorInAccessoryFile(string DeviceName)
        {
            int position = DeviceName.IndexOf(" ");
            if (position < 0)
            {
                return DeviceName;
            }
            else
            {
                string subAccessoryName = DeviceName.Substring(0, position);
                return subAccessoryName;
            }
        }


        public bool IsDuplicatedAccessoryExist(List<Resource> Accessory,string AccessoryName,string AccessoryType)
        {
            var target = Accessory.FirstOrDefault(_ => _.ResourceName == AccessoryName && _.ResourceType==AccessoryType);
            if (target == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// 僅處理1對1的關係的配件
        /// </summary>
        /// <param name="FileAddress"></param>
        /// <param name="SheetName"></param>
        public void SetSecondaryResource(string FileAddress, string SheetName)
        {
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileAddress + ";Extended Properties='Excel 8.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from" + "[" + SheetName + "$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string primaryAccessoryType = dr[0].ToString();
                        string primaryAccessoryName = dr[1].ToString();
                        string secondaryAccessoryType = dr[2].ToString();
                        string secondaryAccessoryName = dr[3].ToString();
                        int.TryParse(dr[4].ToString(), out int require_secondary_accessory_quantity);                       
                        if (IsAccessoryBeUsed(primaryAccessoryType, primaryAccessoryName) == true)//to creat substitute relationship to those be used accessory
                        {
                            if (primaryAccessoryName.Substring(0, 2) != "V_" && require_secondary_accessory_quantity == 1)//consider 1 to 1 substitute relationship
                            {
                                SetAllAccessoryList(secondaryAccessoryName, secondaryAccessoryType);
                                secondary_resources.Add(new SecondardResource()
                                {
                                    PrimaryResourceIndex = Get_Accessory_Index(primaryAccessoryName, primaryAccessoryType),
                                    PrimaryResourceType = primaryAccessoryType,
                                    PrimaryResourceName = primaryAccessoryName,
                                    SecondardResourceIndex=Get_Accessory_Index(secondaryAccessoryName,secondaryAccessoryType),
                                    SecondaryResourceType = secondaryAccessoryType,
                                    SecondaryResourceName = secondaryAccessoryName
                                });
                            }
                        }  
                    }
                }
            }   
        }

        public List<SecondardResource> GetSecondardResources()
        {
            return secondary_resources;
        }

        public string GetPrimaryResourceType(string SecondaryType,string SecondaryName)
        {
            var target = secondary_resources.FirstOrDefault(_ => _.SecondaryResourceType == SecondaryType && _.SecondaryResourceName == SecondaryName);
            return target.PrimaryResourceType;
        }

        public string GetPrimaryResourceName(string SecondaryType, string SecondaryName)
        {
            var target = secondary_resources.FirstOrDefault(_ => _.SecondaryResourceType == SecondaryType && _.SecondaryResourceName == SecondaryName);
            return target.PrimaryResourceName;
        }

        public enum AccessoryType 
        { 
            Accessory_ACC,
            Accessory_Part
        }
    }
}
