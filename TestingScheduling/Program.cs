using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;

namespace TestingScheduling
{
    class Program
    {
        static void Main(string[] args)
        {
            //Step 1.Parameter Setting
            int MAX_ITERIATION = 20;
            int POPULATION_RANDOM = 0;
            int POPULATION_HEURISTICS = 30;
            double CROSSOVER_RATE = 70;
            double MUTATION_RATE = 40;
            double STOP_CRITERIA_RATIO = 0.01;
            int NOT_IMPROVEMENT_TIMES = 20;
            List<string> SCHEDULE_CUSTOMER_CODE=new List<string>() { "IA9" };//schedule any operation meet customer code
            GeneticSetting.ObjectiveFunction objectiveFunction = GeneticSetting.ObjectiveFunction.TotalWeightedTardiness;//objective function
            GeneticSetting.InitialSolution initialSolutionGenerateMethod = GeneticSetting.InitialSolution.ShortestProcess_Setup;
            Chromosome.Data.currentTime = new DateTime(2022, 06, 06, 16, 00, 00);//schedule start time
            
            //Step 2.assign filePath
            string uphAddress= "UPH-A9-0705.xlsx";
            string sheet_uph = "Sheet1";
            string machineListAddress= "MachineList_20220606.xlsx";
            string sheet_machine_list = "工作表1";
            string wipListAddress= "Lot List & Job List(new)_2022060616_Rev0715.xlsx";
            string sheet_job_list = "JobList(new)";
            string sheet_lot_list = "LotList";
            string setupListAddress= "Microship setup time_0705.xlsx";
            string sheet_setup = "setup time";
            string accessoryListAddress = "ACC系統資料0606.xlsx";//accessoryList
            string sheet_require_accessory = "配件主檔定義";
            string sheet_acc_list= "AccList";
            string sheet_part_list= "PartList";
            string sec_acc_sheet_name= "替代料";
            string smart_table_dueDate= "SmartTable 0429.xlsx";
            string sheet_table_dueDate= "ProcessTime";

            //Step 3. Scheduling enviroment initialization            
            JobAndOperation jobAndOperation = new JobAndOperation();
            Machine machine=new Machine();           
            Scheduler scheduled=new Scheduler();

            //step 3-1 To set UPH and setup time
            jobAndOperation.ReadXlsx(uphAddress,JobAndOperation.ProcessingDataType.UPH, sheet_uph);
            jobAndOperation.ReadXlsx(setupListAddress, JobAndOperation.ProcessingDataType.SetupTime, sheet_setup);

            //step 3-2 To generate unrunning lot(未固定排程) and running lot(固定排程)
            switch (objectiveFunction)
            {
                case GeneticSetting.ObjectiveFunction.Makespan:
                    scheduled.ReadData_Makespan(wipListAddress, sheet_job_list, sheet_lot_list, SCHEDULE_CUSTOMER_CODE,smart_table_dueDate,sheet_table_dueDate);
                    break;
                case GeneticSetting.ObjectiveFunction.TotalWeightedTardiness:
                    scheduled.ReadData_Tardiness(wipListAddress, sheet_job_list, sheet_lot_list, SCHEDULE_CUSTOMER_CODE, smart_table_dueDate, sheet_table_dueDate);
                    break;
            }

            // step 3-3 To read the machine list for generating machineIndex and machine type 
            List<Resource> allMachineLists = new List<Resource>();
            allMachineLists=machine.ReadMachineListData(machineListAddress, sheet_machine_list);

            // step 3-3-1. To set eligible tester to each operation.
            List<Job_Operation_Index> unrunningLots = new List<Job_Operation_Index>();//Unschedule one
            unrunningLots = scheduled.GetAllUnRunningLot();       
            machine.ReadEligibleTesterData(wipListAddress, sheet_job_list, unrunningLots);//to generate eligible tester from jobList file

            //Step 3-4 To set accessory for processing lots       
            Accessory accessory = new Accessory();
            List<Job_Operation_Index> runningLots = new List<Job_Operation_Index>();
            runningLots = scheduled.GetAllRunngingLot();

            //step 3-4-1 To set Require Accessory(作業所需的配件) to job and operationjob operation count
            accessory.ReadRequireAccessoryData(accessoryListAddress, sheet_require_accessory, unrunningLots, runningLots, Chromosome.Data.currentTime);

            //step 3-4-2. To generate Machine type. i.e. each machine type is composed by tester and handler
            machine.GenerateMachineType(allMachineLists, unrunningLots,runningLots, scheduled, jobAndOperation, Chromosome.Data.currentTime,accessory);

            //Step 3-4-3 dealing with secondary resource relationship
            accessory.SetSecondaryResource(accessoryListAddress, sec_acc_sheet_name);

            //step 3-4-4 Available quantity of accessory- including acc list and part list            
            accessory.ReadAvailableAccessoryData(accessoryListAddress, sheet_acc_list, Accessory.AccessoryType.Accessory_ACC);
            accessory.ReadAvailableAccessoryData(accessoryListAddress, sheet_part_list, Accessory.AccessoryType.Accessory_Part);
            accessory.SetAvailableAccessory_for_processingLot();//set available quantity of accessory to the processing(running) lot. 
            List<AvailableAccessory> availableQuantity_accessory = new List<AvailableAccessory>(accessory.GetAvailableAccessories());

            //step 3-5 To validate the data in unrunning lots. And remove eligible machine type of job and operation from available_machineType_job_Operation lists if those operations
            //missing setup time or uph. 
            //Remove job from unrunning lots when there isn't any eligible machine to process.
            List<Resource> eligibleHandlers = new List<Resource>(machine.Get_Eligible_Handlers_List());//eligible and available machine of job and operation
            List<Resource> eligibleTesters = new List<Resource>(machine.Get_Elibible_Testers_List());//eligible and available
            scheduled.ValidateUnrunningLot(ref machine, eligibleTesters, eligibleHandlers, jobAndOperation,accessory);

            //step 3-6 To set job and operation index as input of scheduling
            unrunningLots =scheduled.Set_Job_Operation_Index();
            jobAndOperation.SetJob(unrunningLots);// Set Job count
            jobAndOperation.SetJobOperation(unrunningLots);// set number of operation for each job
            
            //step 3-6-1 setting require accessory to each operation
            accessory.Set_Require_Accessory_Job_Operation(unrunningLots);
            List<Resource> using_quantity_accessory = new List<Resource>(accessory.Get_Using_Accessory());
            List<MachineType> available_machine_ik = new List<MachineType>(machine.Get_available_Machine_Job_Operation());//get available machine type index of job and opertaion

            //step 4 Input of genetic algorithm: To set releated parameter for encoding and decoding- including processing time, setup time, release time, family, etc. 
            Chromosome.Data.ProcessingTime_ikm = jobAndOperation.Set_ProcessingTime_Job_Operation_Machine(unrunningLots, scheduled, machine, eligibleTesters, eligibleHandlers);//processing time ikm
            Chromosome.Data.SetupTime_ikm = jobAndOperation.Set_SetupTime_Job_Operation_Machine(unrunningLots, scheduled, machine, eligibleTesters, eligibleHandlers);//Setup ikm
            Chromosome.Data.Job_Operating_Data = jobAndOperation.SetJobInformation(unrunningLots, Chromosome.Data.currentTime);// define release time, due time and weight
            Chromosome.Data.Require_accessory_ik = accessory.Get_Require_Accessory_Job_Operation();//Rik, rquire accessory of job i and operation k
            Chromosome.Data.AvailableQuantity_accessory = availableQuantity_accessory;
            Chromosome.Data.Family_i = jobAndOperation.SetFamily_Of_Job(unrunningLots);// Family of job i 
            Chromosome.Data.SecondaryAccessoryRelation = accessory.GetSecondardResources();
            Chromosome.Data.Job_Operation = unrunningLots;
            
            //sum of processing time and setup time
            List<Processing> summation_processing_setupTime = new List<Processing>();
            summation_processing_setupTime = Copy.DeepClone(Chromosome.Data.ProcessingTime_ikm);
            foreach (Processing setupTime in Chromosome.Data.SetupTime_ikm)
            {
                int target_sorting_index = summation_processing_setupTime.FindIndex(x => x.JobIndex == setupTime.JobIndex
                  && x.OperationIndex == setupTime.OperationIndex && x.MachineTypeIndex == setupTime.MachineTypeIndex);
                summation_processing_setupTime[target_sorting_index].ProcessingTime = summation_processing_setupTime[target_sorting_index].ProcessingTime + setupTime.SetupTime;
            }
            summation_processing_setupTime=summation_processing_setupTime.OrderBy(operation => operation.ProcessingTime).ToList();//its processing time is summation of processing time and setup time
            Chromosome.Data.Sum_ProcessingTime_SetupTime_ikm = summation_processing_setupTime;
            List<int> jobs=new List<int>();
            for(int i = 1; i <= jobAndOperation.GetJob(); i++)
            {
                jobs.Add(i);
            }
            ScheduleParameter scheduleParameter = new ScheduleParameter()
            {
                MachineTypes= machine.Get_MachineType_Tester_Handler(),// machine type relationship with tester and handler
                Testers_availablity = machine.Get_Avaiable_Testers_or_Handlers(Machine.Machine_Type.Tester),//tester availablity
                Handlers_availablity = machine.Get_Avaiable_Testers_or_Handlers(Machine.Machine_Type.Handler),//handler availablity
                AvailableQuantity_accessory = availableQuantity_accessory, //Available quantity of accessory- including acc list and part list
                SecondardResources=accessory.GetSecondardResources()
            };

            //Step 5 Genetic algorithm
            //step 5-1 parameter setting of genetic algorithm
            GeneticSetting geneticSetting = new GeneticSetting()
            {
                initial_method= initialSolutionGenerateMethod,
                num_generation = MAX_ITERIATION,
                num_genes = unrunningLots.Count,
                num_population_by_heuristics = POPULATION_HEURISTICS,
                num_population_by_random= POPULATION_RANDOM,
                total_num_population= POPULATION_HEURISTICS+ POPULATION_RANDOM,
                selection_group_size=2,
                crossover_probability = CROSSOVER_RATE,
                mutation_probability = MUTATION_RATE,
                job = jobs,
                number_operation = jobAndOperation.GetNumberOfOperation(),
                job_operation_index = unrunningLots,
                stop_criteria_ratio= STOP_CRITERIA_RATIO,
                not_improvement_time= NOT_IMPROVEMENT_TIMES,
                max_iteriation = MAX_ITERIATION
            };
            geneticSetting.SetObjectiveFunction(objectiveFunction);

            GeneticAlgorithm geneticAlgorithm = new GeneticAlgorithm(geneticSetting);
            
            //Step 5-2 slove the scheduling problem by Genetic algorithm
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();      
            geneticAlgorithm.RunTheGeneticAlgorithm(scheduleParameter);
            stopwatch.Stop();

            //Step 5-3 Output scheduled best result
            var schedule_finish_Lots = geneticAlgorithm.Best.OutputSchedule(unrunningLots);


            //step 6 generating schedule report
            foreach (Job_Operation_Index lot in schedule_finish_Lots)
            {
                var target_machineType = machine.Get_MachineType_Tester_Handler().FirstOrDefault(_ => _.MachineTypeIndex == lot.MachineTypeIndex);
                string tester = machine.GetTesterName(target_machineType.TesterIndex);
                string handler= machine.GetHandlerName(target_machineType.HandlerIndex);
                lot.Tester=tester;
                lot.Handler=handler;
                scheduled.SetSchedule(lot);
            }

            runningLots = scheduled.GetAllRunngingLot();
            runningLots=runningLots.OrderBy(x => x.Tester).ThenBy(x=>x.TrackInDate).ThenBy(x=>x.WorkOrderNumber).ToList();//依據要求進行排序
            string time = DateTime.Now.ToString("MMddHHmm");
            scheduled.OutPutScheduleResult(runningLots, "ScheduleResult"+time, jobAndOperation);
            Console.Write("Running time: " + stopwatch.ElapsedMilliseconds+" milliseconds");
        }
    }
}
