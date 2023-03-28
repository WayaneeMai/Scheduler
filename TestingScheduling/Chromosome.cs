using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Diagnostics;

namespace TestingScheduling
{
    [Serializable]
    public class Chromosome
    {
        double objective_value;
        bool objectiveCalculated;

        List<Job_Operation_Index> startSetupTime;
        List<Job_Operation_Index> startProcessingTime;
        List<Job_Operation_Index> finishTime;

        public int[] machine_assignment { get; set; }
        public int[] sequence_operation { get; set; }
        List<int> scheduled_Job;
        
        public ScheduleParameter environmentSetting;
        public GeneticSetting geneSetting;

        public Chromosome(int[] Sequence_Operation, int[] Machine_Assignment, ScheduleParameter EnvironmentSetting, GeneticSetting geneticSetting)
        {
            sequence_operation = Sequence_Operation;
            machine_assignment = Machine_Assignment;

            objectiveCalculated = false;
            objective_value = 0;
            startSetupTime = new List<Job_Operation_Index>();
            startProcessingTime = new List<Job_Operation_Index>();
            finishTime = new List<Job_Operation_Index>();
            environmentSetting =Copy.DeepClone(EnvironmentSetting);
            scheduled_Job = new List<int>();
            geneSetting = geneticSetting;
        }

        public double Fitness
        {
            get
            {
                return ObjectiveValue();
            }
        }

        public double ObjectiveValue()
        {
            if (objectiveCalculated)
                return objective_value;

            switch (geneSetting.GetObjectiveFunction())
            {
                case GeneticSetting.ObjectiveFunction.Makespan:
                    objective_value = Makespan_Decoding_FirstServe();
                    //Console.WriteLine("makespan");
                    break;
                case GeneticSetting.ObjectiveFunction.TotalWeightedTardiness:
                    objective_value = Tardiness_Decoding_FirstServe();
                    break;
            }
            objectiveCalculated = true;
            return objective_value;
        }

        public double Makespan_Decoding_FirstServe()
        {
            List<Job_Operation_Index> last_job_operation = new List<Job_Operation_Index>();//machine m's last operation
            last_job_operation = Copy.DeepClone(Chromosome.Data.Machine_Status_RunningLots);
            for (int i = 0; i < sequence_operation.Length; i++)//to get the gene from operation permutation vector
            {                
                int selectedJob = sequence_operation[i];//scheduled job and operation
                int selectedOperation = scheduled_Job.Where(scheduled_Job => scheduled_Job == selectedJob).ToList().Count() + 1;
                
                //to get the gene of assignment from machine assignment vector
                var machine_assignment_index = Chromosome.Data.Job_Operation.FindIndex(x => x.JobIndex == selectedJob && x.OperationIndex == selectedOperation);
                int assignRule = machine_assignment[machine_assignment_index];

                //to generate available machine types set by the assignment rule.
                MachineSelectionRule machineSelectionRule = Machine_Selection_Makespan(assignRule);
                int select_machine_index = GetCorrespondMachine(machineSelectionRule,selectedJob,selectedOperation,last_job_operation);

                //assign the operation to earlist available machine type.
                int selectedTester = environmentSetting.MachineTypes.FirstOrDefault(machine => machine.MachineTypeIndex == select_machine_index).TesterIndex;//assignment machine type's tester index
                int selectedHandler = environmentSetting.MachineTypes.FirstOrDefault(machine => machine.MachineTypeIndex == select_machine_index).HandlerIndex;//assignment machine type's handler index
                Resource available_tester = environmentSetting.Testers_availablity.FirstOrDefault(tester => tester.ResourceIndex == selectedTester);
                Resource available_handler = environmentSetting.Handlers_availablity.FirstOrDefault(handler => handler.ResourceIndex == selectedHandler);
                double machineAvailableTime = Math.Max(available_tester.AvailableTime, available_handler.AvailableTime);//available time of machine will be maximun value of tester and handler available time.

                //to get required accessory of job and operation
                var require_accessory = Chromosome.Data.Require_accessory_ik.Where(demand_acc => demand_acc.JobIndex == selectedJob && demand_acc.OperationIndex == selectedOperation).ToList();
                List<AvailableAccessory> assignment_resource = new List<AvailableAccessory>();
                double available_time_accessory;

                List<AvailableAccessory> candidate_accessory = new List<AvailableAccessory>();                
                assignment_resource = AssignAccessory(require_accessory, machineAvailableTime);//to assign accessory to operation
                available_time_accessory = assignment_resource.Max(accessory => accessory.Time);//earlist available time of accessory
                
                double precedence_operation_finishTime= Get_precedence_operation_finishTime(selectedJob);//define complement time of previous operation of same job
                double releaseTime = Chromosome.Data.Job_Operating_Data.FirstOrDefault(_ => _.JobIndex == selectedJob).ReleaseTime_of_Job;

                //start setup time will be complement time of last operation, release time, available time of machine and available time accessory
                double start_setup_time = Math.Max(available_time_accessory, Math.Max(machineAvailableTime, Math.Max(releaseTime, precedence_operation_finishTime)));
                double setup_time = GetSetupTime(selectedJob, selectedOperation, select_machine_index, last_job_operation);
                double start_processing_time = start_setup_time + setup_time;//start processing time equal to start setup time plus setup time
                double processing_time = Chromosome.Data.ProcessingTime_ikm.FirstOrDefault(operation => operation.JobIndex == selectedJob && operation.OperationIndex == selectedOperation
                && operation.MachineTypeIndex == select_machine_index).ProcessingTime;            
                double finish_time = start_processing_time + processing_time;//complection time is summation of start processing time + processing time
                SetSchedule(selectedJob, selectedOperation, select_machine_index, 
                    start_setup_time,start_processing_time,finish_time);// schedule the job and operation

                //define last job and operation on machine
                int machineLastJob = last_job_operation.FindIndex(_ => _.MachineTypeIndex == select_machine_index);//last job and operation on machine
                if (machineLastJob != -1)
                {
                    last_job_operation[machineLastJob].JobIndex = selectedJob;
                    last_job_operation[machineLastJob].OperationIndex = selectedOperation;
                    last_job_operation[machineLastJob].FTStep = Chromosome.Data.Job_Operation.FirstOrDefault(operation => operation.JobIndex == selectedJob 
                    && operation.OperationIndex == selectedOperation).FTStep;
                    //last_job_operation[machineLastJob].FamilyIndex = Chromosome.Data.Family_i.FirstOrDefault(operation => operation.JobIndex == selectedJob).FamilyIndex;
                    last_job_operation[machineLastJob].DeviceName = Chromosome.Data.Family_i.FirstOrDefault(operation => operation.JobIndex == selectedJob).DeviceName;
                }
                else
                {
                    last_job_operation.Add(new Job_Operation_Index()//first job and operation of machine
                    {
                        //FamilyIndex = Chromosome.Data.Family_i.FirstOrDefault(operation => operation.JobIndex == selectedJob).FamilyIndex,
                        DeviceName = Chromosome.Data.Family_i.FirstOrDefault(operation => operation.JobIndex==selectedJob).DeviceName,
                        FTStep=Chromosome.Data.Job_Operation.FirstOrDefault(operation=>operation.JobIndex==selectedJob&&operation.OperationIndex==selectedOperation).FTStep,
                        JobIndex = selectedJob,
                        OperationIndex = selectedOperation,
                        MachineTypeIndex = select_machine_index
                    }) ;
                }
                //update available time of accessory, tester and handler. the available time will be completion of the job and operation
                UpdateStatusOfResource(select_machine_index, selectedTester, selectedHandler, finish_time, assignment_resource);

                //Add to the scheduled sets once scheduled
                scheduled_Job.Add(selectedJob);
            }
            return finishTime.Max(finishTime => finishTime.Time);//definition of makespan
        }

        public double Tardiness_Decoding_FirstServe()
        {         
            List<Job_Operation_Index> last_job_operation = new List<Job_Operation_Index>();//machine m's last operation
            last_job_operation = Copy.DeepClone(Chromosome.Data.Machine_Status_RunningLots);
            for (int i = 0; i < sequence_operation.Length; i++)//to get the gene from operation permutation vector
            {
                int selectedJob = sequence_operation[i];//scheduled job and operation
                int selectedOperation = scheduled_Job.Where(scheduled_Job => scheduled_Job == selectedJob).Count() + 1;

                //to get the gene of assignment from machine assignment vector
                var machine_assignment_index = Chromosome.Data.Job_Operation.FindIndex(x => x.JobIndex == selectedJob && x.OperationIndex == selectedOperation);
                int assignRule = machine_assignment[machine_assignment_index];

                //to generate available machine types set by the assignment rule.
                MachineSelectionRule machineSelectionRule = Machine_Selection_Tardiness(assignRule);
                //List<MachineType> candidate_machines = Machine_Selection_Tardiness(assignRule, selectedJob, selectedOperation, last_job_operation);
                int select_machine_index = GetCorrespondMachine(machineSelectionRule,selectedJob,selectedOperation,last_job_operation);


                //assign the operation to earlist available machine type.
                //MachineType selected_Machine = candidate_machines[select_machine_index];
                MachineType selectedMachine = environmentSetting.MachineTypes.FirstOrDefault(machine => machine.MachineTypeIndex == select_machine_index);//assignment machine type index
                Resource available_tester = environmentSetting.Testers_availablity.FirstOrDefault(tester => tester.ResourceIndex == selectedMachine.TesterIndex);
                Resource available_handler = environmentSetting.Handlers_availablity.FirstOrDefault(handler => handler.ResourceIndex == selectedMachine.HandlerIndex);
                double machineAvailableTime = Math.Max(available_tester.AvailableTime, available_handler.AvailableTime);//available time of machine will be maximun value of tester and handler available time.

                //to get required accessory of job and operation
                //var require_accessory = Chromosome.Data.Require_accessory_ik.Where(demand_acc => demand_acc.JobIndex == selectedJob && demand_acc.OperationIndex == selectedOperation).ToList();
                var require_accessory = Chromosome.Data.Require_accessory_ik.Where(demand_acc => demand_acc.JobIndex == selectedJob && demand_acc.OperationIndex == selectedOperation);
                List<AvailableAccessory> assignment_resource = new List<AvailableAccessory>();
                double available_time_accessory;

                //to assign accessory to operation
                
                //assignment_resource = AssignAccessory(require_accessory, machineAvailableTime);
                assignment_resource=AssignAccessory_validate(require_accessory, machineAvailableTime);//validate only 0319
                available_time_accessory = assignment_resource.Max(accessory => accessory.Time);//earlist available time of accessory

                //define complement time of previous operation of same job
                double precedence_operation_finishTime = Get_precedence_operation_finishTime(selectedJob);
                double releaseTime = Chromosome.Data.Job_Operating_Data.FirstOrDefault(_ => _.JobIndex == selectedJob).ReleaseTime_of_Job;

                //start setup time will be complement time of last operation, release time, available time of machine and available time accessory
                double start_setup_time = Math.Max(available_time_accessory, Math.Max(machineAvailableTime, Math.Max(releaseTime, precedence_operation_finishTime)));               
                double setup_time = GetSetupTime(selectedJob, selectedOperation, select_machine_index, last_job_operation);
                double start_processing_time = start_setup_time + setup_time; //start processing time equal to start setup time plus setup time              
                double processing_time = Chromosome.Data.ProcessingTime_ikm.FirstOrDefault(operation => operation.JobIndex == selectedJob && operation.OperationIndex == selectedOperation
                && operation.MachineTypeIndex == select_machine_index).ProcessingTime;
                //complection time is summation of start processing time + processing time
                double finish_time = start_processing_time + processing_time;//setup time+processing time
                SetSchedule(selectedJob, selectedOperation, select_machine_index,
                    start_setup_time, start_processing_time, finish_time);// schedule the job and operation

                //define last job and operation on machine
                int machineLastJob = last_job_operation.FindIndex(_ => _.MachineTypeIndex == select_machine_index);//last job and operation on machine
                if (machineLastJob != -1)
                {
                    last_job_operation[machineLastJob].JobIndex = selectedJob;
                    last_job_operation[machineLastJob].OperationIndex = selectedOperation;
                    last_job_operation[machineLastJob].FTStep = Chromosome.Data.Job_Operation.FirstOrDefault(operation => operation.JobIndex == selectedJob
                    && operation.OperationIndex == selectedOperation).FTStep;
                    last_job_operation[machineLastJob].DeviceName = Chromosome.Data.Family_i.FirstOrDefault(operation => operation.JobIndex == selectedJob).DeviceName;
                }
                else
                {
                    last_job_operation.Add(new Job_Operation_Index()//first job and operation of machine
                    {
                        DeviceName = Chromosome.Data.Family_i.FirstOrDefault(operation => operation.JobIndex==selectedJob).DeviceName,
                        FTStep = Chromosome.Data.Job_Operation.FirstOrDefault(operation => operation.JobIndex == selectedJob && operation.OperationIndex == selectedOperation).FTStep,
                        JobIndex = selectedJob,
                        OperationIndex = selectedOperation,
                        MachineTypeIndex = select_machine_index
                    });
                }
                //update available time of accessory, tester and handler. the available time will be completion of the job and operation
                UpdateStatusOfResource(select_machine_index, selectedMachine.TesterIndex, selectedMachine.HandlerIndex, finish_time, assignment_resource);

                //Add to the scheduled sets once scheduled
                scheduled_Job.Add(selectedJob);
            }
            
            //tardiness=Max(completion time-due time,0)
            //total weighted of tardiness=total weight of tardiness + weight*tardiness
            double weightedTardiness = 0;
            foreach (Job_Operation_Index job in Chromosome.Data.Job_Operating_Data)
            {
                double weight = job.Weight;
                weightedTardiness = weightedTardiness + weight * GetTardiness(job);
            }
            return weightedTardiness;
        }

        public double GetTardiness(Job_Operation_Index Job)
        {
            double completion_time_job = finishTime.Where(operation => operation.JobIndex == Job.JobIndex).Max(x => x.Time);
            double tardiness=Math.Max((completion_time_job-Job.DueTime_of_Job),0);
            return tardiness;
        }

        public void SetSchedule(int Job,int Operation,int AssignMachineIndex,
            double StartSetupTime,double StartProcessingTime,double CompletionTime)
        {
            startSetupTime.Add(new Job_Operation_Index()
            {
                JobIndex = Job,
                OperationIndex = Operation,
                MachineTypeIndex = AssignMachineIndex,
                Time = StartSetupTime
            });

            startProcessingTime.Add(new Job_Operation_Index()
            {
                JobIndex = Job,
                OperationIndex = Operation,
                MachineTypeIndex = AssignMachineIndex,
                Time = StartProcessingTime
            });

            finishTime.Add(new Job_Operation_Index()
            {
                JobIndex = Job,
                OperationIndex = Operation,
                MachineTypeIndex = AssignMachineIndex,
                Time = CompletionTime
            });
        }

        public void UpdateStatusOfResource(int Machine_Index,int SelectedTester, int SelectedHandler,double FinishTime,List<AvailableAccessory> Assignment_Accessories)
        {
            //var index_tester = environmentSetting.Testers_availablity.FindIndex(tester => tester.ResourceIndex == SelectedTester);
            //environmentSetting.Testers_availablity[index_tester].AvailableTime = FinishTime;//update available time of tester  
            //environmentSetting.Testers_availablity[index_tester].MachineTypeIndex = Machine_Index;//定義目前測試機所屬的機器組合編號
            //var index_handler = environmentSetting.Handlers_availablity.FindIndex(handler => handler.ResourceIndex == SelectedHandler);
            //environmentSetting.Handlers_availablity[index_handler].AvailableTime = FinishTime;
            //environmentSetting.Handlers_availablity[index_handler].MachineTypeIndex = Machine_Index;
            var target_tester = environmentSetting.Testers_availablity.FirstOrDefault(tester => tester.ResourceIndex == SelectedTester);
            target_tester.AvailableTime = FinishTime;
            target_tester.MachineTypeIndex = Machine_Index;
            var target_handler = environmentSetting.Handlers_availablity.FirstOrDefault(handler => handler.ResourceIndex == SelectedHandler);
            target_handler.AvailableTime = FinishTime;
            target_handler.MachineTypeIndex = Machine_Index;
            

            //增加尋找目前所有可用的機器組合
            var machines = environmentSetting.MachineTypes.Where(machine => machine.TesterIndex == SelectedTester || machine.HandlerIndex == SelectedHandler);
            foreach (var machine in machines)
            {
                machine.AvailableTime = FinishTime;
            }
            //var index_machine_type = environmentSetting.MachineTypes.FindIndex(machine => machine.MachineTypeIndex == Machine_Index);
            //environmentSetting.MachineTypes[index_machine_type].AvailableTime = FinishTime;

            //update available time of accessory
            foreach (AvailableAccessory accessory in Assignment_Accessories)
            {
                var index_accessory = environmentSetting.AvailableQuantity_accessory.FindIndex(_ => _.AvailableAccessoryIndex == accessory.AvailableAccessoryIndex
                && _.Accessory_Type_Index == accessory.Accessory_Type_Index);
                if (index_accessory != -1)
                {
                    environmentSetting.AvailableQuantity_accessory[index_accessory].Time = FinishTime;//index out of range
                }
            }
        }

        public List<AvailableAccessory> AssignAccessory_validate(IEnumerable<Resource> Require_Accessory, double MachineAvailableTime)
        {
            List<AvailableAccessory> assignment_accessory = new List<AvailableAccessory>();
            foreach (Resource require_accessory in Require_Accessory)
            {
                List<AvailableAccessory> candidate_accessories = environmentSetting.AvailableQuantity_accessory.
                    Where(available_accessory => available_accessory.Accessory_Type_Index == require_accessory.ResourceIndex).ToList();
                var secondary_accessories_available = Chromosome.Data.SecondaryAccessoryRelation.FirstOrDefault(substitute => substitute.PrimaryResourceIndex == require_accessory.ResourceIndex);
                if (secondary_accessories_available != null)
                {
                    IEnumerable<AvailableAccessory> secondary_accessories = environmentSetting.AvailableQuantity_accessory.Where(available_accessory =>
                     available_accessory.Accessory_Type_Index == secondary_accessories_available.SecondardResourceIndex);
                    candidate_accessories.AddRange(secondary_accessories);
                }

                foreach (AvailableAccessory available_accessory in candidate_accessories)
                {
                    available_accessory.WaitingTime = Math.Abs(MachineAvailableTime - available_accessory.Time);
                    //進行排序
                }
                candidate_accessories = candidate_accessories.OrderBy(x => x.WaitingTime).ToList();

                //to assign accessory until its quantity meet the requirement.
                for (int k = 0; k < require_accessory.RequireQuantity; k++)
                {
                    assignment_accessory.Add(new AvailableAccessory()//Assign resource to job and operation
                    {
                        Accessory_Type_Index = candidate_accessories[k].Accessory_Type_Index,
                        AvailableAccessoryIndex = candidate_accessories[k].AvailableAccessoryIndex,
                        Time = candidate_accessories[k].Time
                    });
                }
            }
            return assignment_accessory;
        }

        public List<AvailableAccessory> AssignAccessory(List<Resource> Require_Accessory,double MachineAvailableTime)
        {
            List<AvailableAccessory> assignment_accessory = new List<AvailableAccessory>();           
            while (Require_Accessory.Count() != 0)
            {
                List<AvailableAccessory> candidate_accessory = new List<AvailableAccessory>();            
                candidate_accessory = environmentSetting.AvailableQuantity_accessory.
                    Where(available_accessory => available_accessory.Accessory_Type_Index == Require_Accessory[0].ResourceIndex).ToList();

                //to add secondary accessory to candidate accessory.
                int secondaryIndex;
                var target_secondary = Chromosome.Data.SecondaryAccessoryRelation.FirstOrDefault(substitute => substitute.PrimaryResourceIndex == Require_Accessory[0].ResourceIndex);
                if (target_secondary != null)
                {
                    secondaryIndex = target_secondary.SecondardResourceIndex;
                    var sec_accessories = environmentSetting.AvailableQuantity_accessory.Where(available_accessory => available_accessory.Accessory_Type_Index == secondaryIndex);
                    foreach (AvailableAccessory sec_accessory in sec_accessories)
                    {
                        candidate_accessory.Add(sec_accessory);
                    }
                }

                foreach (AvailableAccessory accessory in candidate_accessory)
                {
                    double waitingTime = Math.Abs(MachineAvailableTime - accessory.Time);
                    accessory.WaitingTime = waitingTime;
                }
                candidate_accessory = candidate_accessory.OrderBy(x => x.WaitingTime).ToList();

                //to assign accessory until its quantity meet the requirement.
                for (int k = 0; k < Require_Accessory[0].RequireQuantity; k++)
                {
                    assignment_accessory.Add(new AvailableAccessory()//Assign resource to job and operation
                    {
                        Accessory_Type_Index = candidate_accessory[k].Accessory_Type_Index,
                        AvailableAccessoryIndex = candidate_accessory[k].AvailableAccessoryIndex,
                        Time = candidate_accessory[k].Time
                    });
                }
                Require_Accessory.Remove(Require_Accessory[0]);//once finish the assignment of required accessory. Remove the accessory from Required Accessory sets.
            }
            return assignment_accessory;
        }

        public double Get_precedence_operation_finishTime(int Job)
        {
            int predecence_operation = scheduled_Job.Where(scheduled_Job => scheduled_Job == Job).Count();//last operation of same job
            double precedence_operation_finishTime=0;
            if (predecence_operation >= 1)
            {
                precedence_operation_finishTime = finishTime.FirstOrDefault(_ => _.JobIndex == Job && _.OperationIndex == predecence_operation).Time;
            }
            return precedence_operation_finishTime;
        }

        public double GetSetupTime(int Job, int Operation, int Machine_Index, List<Job_Operation_Index> Last_Job_Operation)
        {
            double setupTime = Chromosome.Data.SetupTime_ikm.FirstOrDefault(operation => operation.JobIndex == Job && operation.OperationIndex == Operation
                    && operation.MachineTypeIndex == Machine_Index).SetupTime;
            Job_Operation_Index last_job = Last_Job_Operation.FirstOrDefault(machine => machine.MachineTypeIndex == Machine_Index);
            //int family_selected_job = Chromosome.Data.Family_i.FirstOrDefault(job => job.JobIndex == Job).FamilyIndex;
            string device_selected_job = Chromosome.Data.Family_i.FirstOrDefault(job => job.JobIndex == Job).DeviceName;
            string step = Chromosome.Data.Job_Operation.FirstOrDefault(operation => operation.JobIndex == Job && operation.OperationIndex == Operation).FTStep;
            if (last_job != null && last_job.DeviceName == device_selected_job && step == last_job.FTStep)
            {
                int tester = environmentSetting.MachineTypes.FirstOrDefault(machine => machine.MachineTypeIndex == Machine_Index).TesterIndex;
                int handler = environmentSetting.MachineTypes.FirstOrDefault(machine => machine.MachineTypeIndex == Machine_Index).HandlerIndex;
                int testerBelongMachineType = environmentSetting.Testers_availablity.FirstOrDefault(machine => machine.ResourceIndex == tester).MachineTypeIndex;
                int handlerBelongMachineType = environmentSetting.Handlers_availablity.FirstOrDefault(machine => machine.ResourceIndex == handler).MachineTypeIndex;
                if (testerBelongMachineType == Machine_Index && handlerBelongMachineType == Machine_Index)
                {
                    setupTime = 0;                   
                }
            }
            return setupTime;
        }

        public int GetCorrespondMachine(MachineSelectionRule SelectionRule, int Job, int Operation, List<Job_Operation_Index> PreviousOperation)
        {
            int machineIndex;
            switch (SelectionRule)
            {
                case MachineSelectionRule.EarlistAvailableMachine:
                    machineIndex = SelectionRule_EarlistAvailableMachine(Job, Operation);
                    return machineIndex;
                case MachineSelectionRule.Shortest_ProcessingTime:
                    machineIndex = SelectionRule_ShortestProcessingTime(Job, Operation);
                    return machineIndex;
                case MachineSelectionRule.Shortest_SetupTime:
                    machineIndex = SelectionRule_ShortestSetupTime(Job, Operation);
                    return machineIndex;
                case MachineSelectionRule.EarlistCompletionTime:
                    machineIndex = SelectionRule_EarlistCompletionTime(Job, Operation, PreviousOperation);
                    return machineIndex;
                case MachineSelectionRule.Shortest_Process_and_SetupTime:
                    machineIndex = SelectionRule_ShortestSetup_ProcessingTime(Job, Operation);
                    return machineIndex;
                default:
                    Console.WriteLine("machine rule check");
                    return 0;
            }
        }

        public MachineSelectionRule Machine_Selection_Makespan(int AssignRule)
        {
            switch (AssignRule)
            {
                case 1:
                    return MachineSelectionRule.Shortest_ProcessingTime;
                case 2:
                    return MachineSelectionRule.Shortest_SetupTime;
                case 3:
                    return MachineSelectionRule.Shortest_Process_and_SetupTime;
                case 4:
                    return MachineSelectionRule.EarlistAvailableMachine;
                case 5:
                    return MachineSelectionRule.EarlistCompletionTime;
                default:
                    Console.WriteLine("machine selection number didn't define");
                    return MachineSelectionRule.Shortest_ProcessingTime;
            }
        }

        public int SelectionRule_EarlistAvailableMachine(int Job,int Operation)
        {
            List<Processing> target_machines=Chromosome.Data.ProcessingTime_ikm.Where(x=>x.JobIndex==Job&&x.OperationIndex==Operation).ToList();
            /*
            int select_machine_index = IndexOfMin_Available(target_machines);
            return target_machines[select_machine_index].MachineTypeIndex;
            */

            return IndexOfMin_Available(target_machines);
        }
        

        public MachineSelectionRule Machine_Selection_Tardiness(int AssignRule)
        {
            switch (AssignRule)
            {
                case 1:
                    return MachineSelectionRule.Shortest_ProcessingTime;
                case 2:
                    return MachineSelectionRule.Shortest_Process_and_SetupTime;
                case 3:
                    return MachineSelectionRule.EarlistCompletionTime;
                case 4:
                    return MachineSelectionRule.EarlistAvailableMachine;
                default:
                    Console.WriteLine("machine selection number didn't define");
                    return MachineSelectionRule.Shortest_ProcessingTime;
            }
        }

        public int SelectionRule_EarlistCompletionTime(int Job,int Operation, List<Job_Operation_Index> PreviousOperation)
        {
            List<Processing> target_earlist_completion = new List<Processing>();
            var target_machines = Chromosome.Data.ProcessingTime_ikm.Where(operation => operation.JobIndex == Job && operation.OperationIndex == Operation);
            foreach (Processing machine in target_machines)
            {
                double processingTime = target_machines.FirstOrDefault(operation => operation.MachineTypeIndex == machine.MachineTypeIndex).ProcessingTime;
                var setupTime = GetSetupTime(Job, Operation, machine.MachineTypeIndex, PreviousOperation);
                double machine_available_time = machine.AvailableTime;
                double completionTime = machine_available_time + processingTime + setupTime;
                target_earlist_completion.Add(new Processing()
                {
                    MachineTypeIndex = machine.MachineTypeIndex,
                    CompletionTime = completionTime
                });
            }
            //target_earlist_completion = target_earlist_completion.OrderBy(x => x.CompletionTime).ToList();
            //return target_earlist_completion[0].MachineTypeIndex;
            return GetMinEarlistCompletionMachine(target_earlist_completion);
        }

        public int GetMinEarlistCompletionMachine(List<Processing> Machines)
        {
            double minCompletionTime=Machines[0].CompletionTime;
            int selectedMachine = Machines[0].MachineTypeIndex;
            for (int i = 1; i < Machines.Count(); i++)
            {
                if (minCompletionTime > Machines[i].CompletionTime)
                {
                    minCompletionTime = Machines[i].CompletionTime;
                    selectedMachine = Machines[i].MachineTypeIndex;
                }
            }
            return selectedMachine;
        }


        public int SelectionRule_ShortestProcessingTime(int Job, int Operation)
        {
            /*
            var target_processingTime = Chromosome.Data.ProcessingTime_ikm.Where(_ => _.JobIndex == Job && _.OperationIndex == Operation).Min(_ => _.ProcessingTime);
            List<Processing> target_machines = Chromosome.Data.ProcessingTime_ikm.Where(x => x.JobIndex == Job && x.OperationIndex == Operation && x.ProcessingTime == target_processingTime).ToList();
            int select_machine_index = IndexOfMin_Available(target_machines);
            return target_machines[select_machine_index].MachineTypeIndex;
            */
            List<Processing> target_machines = new List<Processing>();
            target_machines = Get_Shortest_Processing_Time_Machines(Job, Operation);
            return IndexOfMin_Available(target_machines);
        }


        public List<Processing> Get_Shortest_Processing_Time_Machines(int Job, int Operation)
        {
            List<Processing> ProcessingTime_Job_Operation = new List<Processing>();
            ProcessingTime_Job_Operation=Chromosome.Data.ProcessingTime_ikm.Where(x => x.JobIndex == Job & x.OperationIndex == Operation).ToList();
            double minProcessingTime = ProcessingTime_Job_Operation[0].ProcessingTime;
            List<Processing> results = new List<Processing>();
            for (int i = 0; i < ProcessingTime_Job_Operation.Count(); i++)
            {
                if (ProcessingTime_Job_Operation[i].ProcessingTime < minProcessingTime)
                {
                    minProcessingTime = ProcessingTime_Job_Operation[i].ProcessingTime;
                    results.Clear();
                    results.Add(ProcessingTime_Job_Operation[i]);
                }
                else if (ProcessingTime_Job_Operation[i].ProcessingTime == minProcessingTime)
                {
                    results.Add(ProcessingTime_Job_Operation[i]);
                }
            }
            return results;
        }

        public int SelectionRule_ShortestSetupTime(int Job, int Operation)
        {
            /*
            var target_setupTime = Chromosome.Data.SetupTime_ikm.Where(_ => _.JobIndex == Job && _.OperationIndex == Operation).Min(_ => _.SetupTime);
            List<Processing> target_machines = Chromosome.Data.SetupTime_ikm.Where(_ => _.JobIndex == Job
            && _.OperationIndex == Operation && _.SetupTime == target_setupTime).ToList();
            int select_machine_index = IndexOfMin_Available(target_machines);
            return target_machines[select_machine_index].MachineTypeIndex;
            */

            List<Processing> target_machines = new List<Processing>();
            target_machines=Get_Shortest_Setup_Time_Machines(Job, Operation);
            return IndexOfMin_Available(target_machines);
        }

        public List<Processing> Get_Shortest_Setup_Time_Machines(int Job,int Operation)
        {
            List<Processing> SetupTime_Job_Operation = new List<Processing>();
            SetupTime_Job_Operation = Chromosome.Data.SetupTime_ikm.Where(x => x.JobIndex == Job & x.OperationIndex == Operation).ToList();
            double minSetupTime = SetupTime_Job_Operation[0].SetupTime;
            List<Processing> results = new List<Processing>();
            for (int i = 0; i < SetupTime_Job_Operation.Count(); i++)
            {
                if (SetupTime_Job_Operation[i].SetupTime < minSetupTime)
                {
                    minSetupTime = SetupTime_Job_Operation[i].SetupTime;
                    results.Clear();
                    results.Add(SetupTime_Job_Operation[i]);
                }
                else if (SetupTime_Job_Operation[i].SetupTime == minSetupTime)
                {
                    results.Add(SetupTime_Job_Operation[i]);
                }
            }
            return results;
        }

        public int SelectionRule_ShortestSetup_ProcessingTime(int Job, int Operation)
        {
            /*
            var target_processing_setupTime = Chromosome.Data.Sum_ProcessingTime_SetupTime_ikm.Where(_ => _.JobIndex == Job && _.OperationIndex == Operation).Min(_ => _.ProcessingTime);
            List<Processing> target_machines = Chromosome.Data.Sum_ProcessingTime_SetupTime_ikm.
                Where(x => x.JobIndex == Job && x.OperationIndex == Operation && x.ProcessingTime == target_processing_setupTime).ToList();
            int select_machine_index = IndexOfMin_Available(target_machines);

            return target_machines[select_machine_index].MachineTypeIndex;
            */
            List<Processing> target_machines = new List<Processing>();
            target_machines = Get_Shortest_Setup_Processing_Time_Machines(Job, Operation);
            return IndexOfMin_Available(target_machines);

        }

        public List<Processing> Get_Shortest_Setup_Processing_Time_Machines(int Job,int Operation)
        {
            List<Processing> Setup_ProcessingTime_Job_Operation = new List<Processing>();
            Setup_ProcessingTime_Job_Operation = Chromosome.Data.Sum_ProcessingTime_SetupTime_ikm.Where(x => x.JobIndex == Job && x.OperationIndex == Operation).ToList();
            double minProcessing_SetupTime = Setup_ProcessingTime_Job_Operation[0].ProcessingTime;
            List<Processing> results = new List<Processing>();
            for (int i = 0; i < Setup_ProcessingTime_Job_Operation.Count(); i++)
            {
                if (Setup_ProcessingTime_Job_Operation[i].ProcessingTime < minProcessing_SetupTime)
                {
                    minProcessing_SetupTime = Setup_ProcessingTime_Job_Operation[i].ProcessingTime;
                    results.Clear();
                    results.Add(Setup_ProcessingTime_Job_Operation[i]);
                }
                else if (Setup_ProcessingTime_Job_Operation[i].ProcessingTime == minProcessing_SetupTime)
                {
                    results.Add(Setup_ProcessingTime_Job_Operation[i]);
                }
            }
            return results;
        }


        public int IndexOfMin_Available(List<Processing> self)
        {
            if (self == null)
            {
                throw new ArgumentNullException("self");
            }

            if (self.Count == 0)
            {
                throw new ArgumentException("List is empty.", "self");
            }

            MachineType target_machine;
            double min = 0;
            //var target_machine = environmentSetting.MachineTypes.FirstOrDefault(x=>x.MachineTypeIndex == self[0].MachineTypeIndex);
            //double min = target_machine.AvailableTime;
            int minIndex = 0;
            for (int i = 0; i < self.Count; i++)
            {
                target_machine = environmentSetting.MachineTypes.FirstOrDefault(x => x.MachineTypeIndex == self[i].MachineTypeIndex);
                if (i == 0 || target_machine.AvailableTime < min)
                {
                    min = target_machine.AvailableTime;
                    minIndex = i;
                }
            }
            return self[minIndex].MachineTypeIndex;//改成回傳Machine type index
        }

        public List<Job_Operation_Index> OutputSchedule(List<Job_Operation_Index> UnrunningLots)
        {
            //save to runningLots
            List<Job_Operation_Index> outputs = new List<Job_Operation_Index>();
            for(int lot=0;lot< sequence_operation.Length; lot++)
            {
                var target_start_setup = startSetupTime.FirstOrDefault(_ => _.JobIndex == UnrunningLots[lot].JobIndex && _.OperationIndex == UnrunningLots[lot].OperationIndex).Time;
                var target_start_processing = startProcessingTime.FirstOrDefault(_ => _.JobIndex == UnrunningLots[lot].JobIndex && _.OperationIndex == UnrunningLots[lot].OperationIndex).Time;
                var target_finish = finishTime.FirstOrDefault(_ => _.JobIndex == UnrunningLots[lot].JobIndex && _.OperationIndex == UnrunningLots[lot].OperationIndex).Time;
                var target_release = Chromosome.Data.Job_Operating_Data.FirstOrDefault(x => x.JobIndex == UnrunningLots[lot].JobIndex).ReleaseTime_of_Job;
                var target_machine=finishTime.FirstOrDefault(_ => _.JobIndex == UnrunningLots[lot].JobIndex && _.OperationIndex == UnrunningLots[lot].OperationIndex).MachineTypeIndex;

                UnrunningLots[lot].TrackInDate = Chromosome.Data.currentTime.AddMinutes(target_start_setup);
                UnrunningLots[lot].TrackInDate_Processing = Chromosome.Data.currentTime.AddMinutes(target_start_processing);
                UnrunningLots[lot].CompletionDate = Chromosome.Data.currentTime.AddMinutes(target_finish);
                UnrunningLots[lot].MachineTypeIndex = target_machine;
                UnrunningLots[lot].ReleaseTime_of_Job = target_release;

            }
            outputs = UnrunningLots;
            return outputs;
        }


        public static class Data
        {
            public static List<Processing> ProcessingTime_ikm;
            public static List<Processing> SetupTime_ikm;
            public static List<Processing> Sum_ProcessingTime_SetupTime_ikm;
            //public static List<Job_Operation_Index> ReleaseTime_i;
            public static List<SecondardResource> SecondaryAccessoryRelation;
            public static List<Job_Operation_Index> Family_i;
            public static List<Resource> Require_accessory_ik;
            public static List<AvailableAccessory> AvailableQuantity_accessory;
            public static DateTime currentTime;
            public static List<Job_Operation_Index> Job_Operation;
            public static List<Job_Operation_Index> Job_Operating_Data;//儲存release time, weight and due time
            public static List<Job_Operation_Index> Machine_Status_RunningLots;//紀錄機台作業的品種、工件與站點等

        }

        public enum MachineSelectionRule
        {
            Shortest_ProcessingTime,
            Shortest_Process_and_SetupTime,
            Shortest_SetupTime,
            EarlistAvailableMachine,
            EarlistCompletionTime
        }

    }
}
