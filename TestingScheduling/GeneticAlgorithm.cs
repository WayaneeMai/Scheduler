using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Diagnostics;

namespace TestingScheduling
{
    class GeneticAlgorithm
    {
        List<Chromosome> populations;// to save each population's fitness, schedule
        private Chromosome best_solution;
        int soultion_update;
        private int total_num_population;//number of solutions in the population.
        private bool valid_parameters = true;
        GeneticSetting geneticSetting;
        private ScheduleParameter environmentSetting;
        public List<Job_Operation_Index> genetic_space=new List<Job_Operation_Index>();//To save job Index;

        public List<int[]> population_operation=new List<int[]>();// initial population of operation sequence vector
        public List<int[]> population_machineSelection = new List<int[]>();// initial population of machine selection vector
        Random random;

        public GeneticAlgorithm(GeneticSetting Setting)//initialization of algorithm and parameter validation
        {
            geneticSetting = Setting;
            best_solution = null;
            soultion_update = 0;

            foreach (Job_Operation_Index jobIndex in Chromosome.Data.Job_Operation)
            {
                genetic_space.Add(new Job_Operation_Index()
                {
                    JobIndex = jobIndex.JobIndex,
                    OperationIndex = jobIndex.OperationIndex,
                });
            }
            Console.WriteLine("genetic_space " + genetic_space.Count);

            //validate the format
            if (geneticSetting.crossover_probability < 0 || geneticSetting.crossover_probability > 100)
            {
                valid_parameters = false;
                Console.WriteLine("The value sassigned to 'crossover rate' must be betweeb 0 and 100 but " + geneticSetting.crossover_probability + " found");
            }

            if (geneticSetting.mutation_probability < 0 || geneticSetting.mutation_probability > 100)
            {
                valid_parameters = false;
                Console.WriteLine("The value sassigned to 'mutation rate' must be betweeb 0 and 100 but " + geneticSetting.mutation_probability + " found");
            }
            
            this.total_num_population = geneticSetting.num_population_by_heuristics + geneticSetting.num_population_by_random;
            populations = new List<Chromosome>();

            //Validating the number of gene.
            if (geneticSetting.num_genes <= 0)
            {
                valid_parameters = false;
                Console.WriteLine("The number of genes cannot be less than 0 but " + geneticSetting.num_genes + " found");
            }
            random = new Random();      
        }

        public void RunTheGeneticAlgorithm(ScheduleParameter EnvironmentSetting)
        {            
            environmentSetting = EnvironmentSetting;

            int iteriation = 1;
            int not_improvement_time = 0;
            double last_iteriation_best_fitness=-1;
            //initializing population
            switch (geneticSetting.GetObjectiveFunction())
            {
                case GeneticSetting.ObjectiveFunction.Makespan:
                    for (int num_chromosome = 0; num_chromosome < geneticSetting.num_population_by_heuristics; num_chromosome++)
                    {
                        Chromosome chromosome = GenerateNewChromosome_makespan();
                        CalculatingFitness(chromosome);
                        populations.Add(chromosome);
                    }
                    for (int num_chromosome = 0; num_chromosome < geneticSetting.num_population_by_random; num_chromosome++)
                    {
                        Chromosome chromosome = GenerateNewChromosome_random();
                        CalculatingFitness(chromosome);
                        populations.Add(chromosome);
                    }
                    break;
                case GeneticSetting.ObjectiveFunction.TotalWeightedTardiness:                  
                    for (int num_chromosome = 0; num_chromosome < geneticSetting.num_population_by_heuristics; num_chromosome++)
                    {

                        Chromosome chromosome =GenerateNewChromosome_Tardiness() ;
                        CalculatingFitness(chromosome);
                        populations.Add(chromosome);
                    }
                    for (int num_chromosome = 0; num_chromosome < geneticSetting.num_population_by_random; num_chromosome++)
                    {
                        Chromosome chromosome = GenerateNewChromosome_random();
                        CalculatingFitness(chromosome);
                        populations.Add(chromosome);
                    }
                    break;
                default:
                    valid_parameters = false;
                    Console.WriteLine("The objective function shall be makespan or tardiness.");
                    break;    
            }
            last_iteriation_best_fitness = best_solution.Fitness;
            Console.WriteLine("iteriation 0 best fitness: " + best_solution.Fitness);

            //validate paramether
            if (valid_parameters == false)
            {
                Console.WriteLine("\nThe RunTheGeneticAlgorithm method cannot be executed with invalid parameters. Please check the parameters of GeneticAlgorithm\n");
                return;
            }

            do
            {
                List<Chromosome> offspring=new List<Chromosome>();
                for (int chromosome = 0; chromosome < geneticSetting.total_num_population; chromosome=chromosome+2)//to generate n's chromosomes
                {
                    int selectedChromosome1 = TournamentSelection();//select two parents for crossover and mutating by ussing binary tournament
                    int selectedChromosome2;
                    do
                    {
                        selectedChromosome2=TournamentSelection();//to pick another chromosome as parant 2.
                    }while (selectedChromosome1==selectedChromosome2);

                    Chromosome mother=Copy.DeepClone(populations[selectedChromosome1]);
                    Chromosome father= Copy.DeepClone(populations[selectedChromosome2]);
                    mother.geneSetting= geneticSetting;
                    father.geneSetting = geneticSetting;

                    Chromosome child1;
                    Chromosome child2;

                    if (random.Next(1, 100) <= geneticSetting.crossover_probability)//crossover to generate offspring
                    {
                        List<Chromosome> childs=new List<Chromosome>(Crossover(mother,father));
                        child1 =Copy.DeepClone(childs[0]);
                        child2 = Copy.DeepClone(childs[1]);
                    }
                    else
                    {
                        child1 = Copy.DeepClone(mother);
                        child2 = Copy.DeepClone(father);
                    }

                    if (random.Next(1, 100) <= geneticSetting.mutation_probability)
                    {
                        TwoPointMutation(child1);//mutation
                        TwoPointMutation(child2);
                    }

                    //populations.Add(child1);
                    //populations.Add(child2);
                    offspring.Add(child1);
                    offspring.Add(child2);

                    child1.geneSetting = geneticSetting;
                    child2.geneSetting = geneticSetting;
                    CalculatingFitness(child1);////measuring the fitness of each chromosome in the population and update best soluting
                    CalculatingFitness(child2);
                }
                Console.WriteLine("iteriation " + iteriation + " best fitness: " + best_solution.Fitness);
                //PrintChromosome("Chromosome iteriation"+ iteriation); 

                //update population
                populations.AddRange(offspring);//0314
                offspring.Clear();//0314
                populations = populations.OrderBy(chromosome => chromosome.Fitness).ToList();
                int remove_num=populations.Count()-geneticSetting.total_num_population;
                populations.RemoveRange(geneticSetting.total_num_population, remove_num);

                //stopping criterior: the best solution found so far has not changed during a predefined number of the last iterations.
                double ratio_improvement = (last_iteriation_best_fitness-best_solution.Fitness) / last_iteriation_best_fitness;
                if (ratio_improvement < geneticSetting.stop_criteria_ratio)
                {
                    not_improvement_time++;
                    if (not_improvement_time > geneticSetting.not_improvement_time)
                        break;
                }
                else
                {
                    last_iteriation_best_fitness = best_solution.Fitness;
                    Console.WriteLine("improvement ratio:" + ratio_improvement + "/improvement less than 0.01 for " + not_improvement_time + " times.");
                    not_improvement_time = 0;
                }
                iteriation++;//update generations
            } while (iteriation <= geneticSetting.max_iteriation);//stop criteria
        }

        Chromosome GenerateNewChromosome_random()
        {
            int[] sequence_operation=new int[geneticSetting.num_genes];
            int[] machine_assignment = new int[geneticSetting.num_genes];

            List<Job_Operation_Index> list = new List<Job_Operation_Index>(genetic_space);
            int column = 0;
            while (list.Count > 0)//operation sequence part
            {
                int index = random.Next(0, list.Count-1);
                int item = list[index].JobIndex;

                list.RemoveAt(index);
                sequence_operation[column] = item;
                machine_assignment[column] = random.Next(1, 6);
                column++;
            }
            return new Chromosome(sequence_operation, machine_assignment, environmentSetting,geneticSetting);
        }

        Chromosome GenerateNewChromosome_makespan()
        {
            int[] sequence_operation = new int[geneticSetting.num_genes];
            int[] machine_assignment = new int[geneticSetting.num_genes];

            int gene_count = 0;
            List<Job_Operation_Index> list = new List<Job_Operation_Index>(genetic_space);
            switch (geneticSetting.initial_method)
            {
                case GeneticSetting.InitialSolution.ShortestProcess_Setup:
                    foreach (Processing priority in Chromosome.Data.Sum_ProcessingTime_SetupTime_ikm)
                    {
                        var target = list.FirstOrDefault(x => x.JobIndex == priority.JobIndex && x.OperationIndex == priority.OperationIndex);
                        if (target != null)
                        {
                            sequence_operation[gene_count] = target.JobIndex;
                            machine_assignment[gene_count] = random.Next(1, 6);
                            list.Remove(target);//already had assigned
                            gene_count++;
                        }
                        if (list.Count() == 0)
                            break;
                    }
                    break;
                case GeneticSetting.InitialSolution.ReleaseTime:
                    List<Job_Operation_Index> ReleaseTime_Copy = new List<Job_Operation_Index>();
                    ReleaseTime_Copy = Copy.DeepClone(Chromosome.Data.Job_Operating_Data);
                    ReleaseTime_Copy = ReleaseTime_Copy.OrderBy(x => x.ReleaseTime_of_Job).ToList();
                    int operationCount = 0;
                    while (gene_count < list.Count())
                    {
                        List<Job_Operation_Index> removed_job = new List<Job_Operation_Index>();
                        operationCount++;
                        foreach(Job_Operation_Index priority in ReleaseTime_Copy)
                        {
                            var job_operation_count=list.Where(x=>x.JobIndex==priority.JobIndex).ToList().Count();
                            if (operationCount > job_operation_count)
                            {
                                removed_job.Add(priority); 
                                continue;                                
                            }
                            sequence_operation[gene_count] = priority.JobIndex;
                            machine_assignment[gene_count] = random.Next(1, 6);
                            gene_count++;
                        }
                        foreach(Job_Operation_Index job in removed_job)
                        {
                            ReleaseTime_Copy.Remove(job);
                        }
                    }                   
                    break;
            }
            return new Chromosome(sequence_operation, machine_assignment,environmentSetting,geneticSetting);
        }

        Chromosome GenerateNewChromosome_Tardiness()
        {
            int[] sequence_operation = new int[geneticSetting.num_genes];
            int[] machine_assignment=new int[geneticSetting.num_genes];
            int gene_count = 0;
            List<Job_Operation_Index> list = new List<Job_Operation_Index>(genetic_space);
            switch (geneticSetting.initial_method)
            {
                case GeneticSetting.InitialSolution.EarlistDueDate:
                    List<Job_Operation_Index> dueTime_sorted=new List<Job_Operation_Index>();
                    dueTime_sorted = Chromosome.Data.Job_Operating_Data;
                    dueTime_sorted = dueTime_sorted.OrderBy(x => x.DueTime_of_Job).ThenBy(x => x.Weight).ToList();
                    foreach(Job_Operation_Index priority in dueTime_sorted)
                    {
                        var target = list.Where(x => x.JobIndex == priority.JobIndex).ToList();
                        int jobCount = 0;
                        while (jobCount < target.Count())
                        {
                            sequence_operation[gene_count]=priority.JobIndex;
                            machine_assignment[gene_count]=random.Next(1,6);
                            jobCount++;
                            gene_count++;

                        }
                    }
                    break;
                case GeneticSetting.InitialSolution.ShortestProcessingTime:
                    List<Processing> ProcessingTime_ikm_Sort = new List<Processing>();
                    ProcessingTime_ikm_Sort= Copy.DeepClone(Chromosome.Data.ProcessingTime_ikm);
                    ProcessingTime_ikm_Sort = ProcessingTime_ikm_Sort.OrderBy(x => x.ProcessingTime).ToList();
                    foreach(Processing priority in ProcessingTime_ikm_Sort)
                    {
                        var target = list.FirstOrDefault(x => x.JobIndex == priority.JobIndex && x.OperationIndex == priority.OperationIndex);
                        if (target != null)
                        {
                            sequence_operation[gene_count] = target.JobIndex;
                            machine_assignment[gene_count] = random.Next(1, 5);

                            list.Remove(target);
                            gene_count++;
                        }

                        if (list.Count == 0)
                        {
                            break;
                        }
                    }
                    break;

                case GeneticSetting.InitialSolution.ShortestProcessingTime_Weighted:
                    List<Processing> Weighted_ProcessingTime_ikm_Sort = new List<Processing>();
                    for (int i = 0; i < genetic_space.Max(x => x.JobIndex); i++)
                    {
                        int operation_Index = Chromosome.Data.Job_Operation.Where(x => x.JobIndex == i + 1).Max(y=>y.OperationIndex);
                        for (int k = 0; k < operation_Index; k++)
                        {
                            double weight = Chromosome.Data.Job_Operating_Data.FirstOrDefault(x => x.JobIndex == i + 1).Weight;
                            double processingTime = Chromosome.Data.ProcessingTime_ikm.FirstOrDefault(x => x.JobIndex == i + 1 && x.OperationIndex == k + 1).ProcessingTime;
                            double temp = weight / processingTime;

                            Weighted_ProcessingTime_ikm_Sort.Add(new Processing()
                            {
                                JobIndex = i + 1,
                                OperationIndex = k+1,
                                ProcessingTime = temp
                            });
                        }            
                    }
                    Weighted_ProcessingTime_ikm_Sort = Weighted_ProcessingTime_ikm_Sort.OrderBy(x => x.ProcessingTime).ToList();

                    foreach (Processing priority in Weighted_ProcessingTime_ikm_Sort)
                    {
                        var target = list.FirstOrDefault(x => x.JobIndex == priority.JobIndex && x.OperationIndex == priority.OperationIndex);
                        if (target != null)
                        {
                            sequence_operation[gene_count] = target.JobIndex;
                            machine_assignment[gene_count] = random.Next(1, 5);

                            list.Remove(target);
                            gene_count++;
                        }

                        if (list.Count == 0)
                        {
                            break;
                        }
                    }
                    break;                 
            }
            return new Chromosome(sequence_operation, machine_assignment, environmentSetting, geneticSetting);
        }


        public void CalculatingFitness(Chromosome chromosome)
        {

            if (best_solution == null || best_solution.Fitness > chromosome.Fitness)
            {
                best_solution = chromosome;
                soultion_update++;
            }
            // each soluting have a fitness value      
            //return fitness value
        }
        int TournamentSelection()
        {            
            var selectNums = Enumerable.Range(0, geneticSetting.total_num_population-1).OrderBy(x => random.Next()).Take(geneticSetting.selection_group_size).ToList();
            int select_chromsome=BinaryTournament(selectNums);
            return select_chromsome;
        }

        int BinaryTournament(List<int> SelectNums)
        {
            double best_chromosome_Fitness=populations[SelectNums[0]].Fitness;
            int best_chromosome= SelectNums[0];
            for (int i = 0; i < SelectNums.Count - 1; i++)
            {
                if (best_chromosome_Fitness > populations[SelectNums[i + 1]].Fitness)
                {
                    best_chromosome_Fitness=populations[(SelectNums[i + 1])].Fitness;
                    best_chromosome = SelectNums[i + 1];
                }
            }
            return best_chromosome;
        }

        List<Chromosome> Crossover(Chromosome Parent1,Chromosome Parent2)// revised 0527
        {
            int[] sequence_operation = new int[geneticSetting.num_genes];
            int[] machine_assignment = new int[geneticSetting.num_genes];
            List<Chromosome> childs = new List<Chromosome>();
            Chromosome child1 = new Chromosome(Parent1.sequence_operation, Parent1.machine_assignment,environmentSetting,geneticSetting);
            Chromosome child2 = new Chromosome(Parent2.sequence_operation, Parent2.machine_assignment, environmentSetting,geneticSetting);
            int crossover_point1 = random.Next(0, geneticSetting.num_genes / 2 + 1);
            int crossover_point2 = crossover_point1 + random.Next(0, geneticSetting.num_genes / 2);
          
            child1.machine_assignment = TwoPointCrossover(Parent1, Parent2, crossover_point1, crossover_point2);
            child2.machine_assignment = TwoPointCrossover(Parent2, Parent1, crossover_point1, crossover_point2);

            int num_set1 = random.Next(1, geneticSetting.job.Max() );
            List<int> set1 = Enumerable.Range(1, geneticSetting.job.Max()).OrderBy(x => random.Next()).Take(num_set1).ToList();// to pick num_set1's job as set1

            List<int> set2 = Enumerable.Range(1, geneticSetting.job.Max()).ToList();
            set2 = set2.Where(x => !set1.Contains(x)).ToList();//自set2尋找不符合set1的值 test

            child1.sequence_operation = JobPrecedenceCorssover(Parent1, Parent2, set1, set2);
            child2.sequence_operation= JobPrecedenceCorssover(Parent2, Parent1, set2, set1);

            childs.Add(child1);
            childs.Add(child2);

            return childs;
        }

        public int[] JobPrecedenceCorssover(Chromosome Parent1, Chromosome Parent2,List<int> Set1, List<int> Set2) //debug 0527
        {
            List<int> offerspring = Parent1.sequence_operation.ToList();
            List<int> offerspring_parent2 = Parent2.sequence_operation.Where(set2_job => Set2.Contains(set2_job)).ToList();

            var index = offerspring.FindIndex(x => Set2.Contains(x));
            while(index != -1)
            {
                offerspring[index] = -1;
                index = offerspring.FindIndex(x => Set2.Contains(x));
            }
            
            int parent2_index = 0;
            for (int i = 0; i < geneticSetting.num_genes; i++)
            {
                if (offerspring[i] == -1)
                {
                    offerspring[i]= offerspring_parent2[parent2_index];
                    parent2_index++;
                }
            }
            return offerspring.ToArray();
        }

        public int[] TwoPointCrossover(Chromosome Parent1, Chromosome Parent2,int Crossover_point1, int Crossover_point2)
        {
            int[] offerspring=new int[geneticSetting.num_genes];

            for(int i = 0; i < Crossover_point1; i++)
            {
                offerspring[i]=Parent1.machine_assignment[i];
            }

            for(int i = Crossover_point2; i < geneticSetting.num_genes; i++)
            {
                offerspring[i] = Parent1.machine_assignment[i];
            }

            for(int i = Crossover_point1; i < Crossover_point2; i++)
            {
                offerspring[i] = Parent2.machine_assignment[i];
            }

            return offerspring;
        }
        public void TwoPointMutation(Chromosome child)
        {
            int x1 = random.Next(0, geneticSetting.num_genes / 2+1);
            int x2 = x1 + random.Next(0, geneticSetting.num_genes / 2);

            //job mutation
            int temp = child.sequence_operation[x1];
            child.sequence_operation[x1] = child.sequence_operation[x2];
            child.sequence_operation[x2] = temp;

            x1 = random.Next(0, geneticSetting.num_genes / 2+1);
            x2 = x1 + random.Next(0, geneticSetting.num_genes / 2);
            //machine mutation
            temp = child.machine_assignment[x1];
            child.machine_assignment[x1] = child.machine_assignment[x2];
            child.machine_assignment[x2] = temp;
        }
        
        public Chromosome Best
        {
            get { return best_solution; }
            set { best_solution = value; }
        }

        public void PrintChromosome(string FileName)
        {
            using (StreamWriter sw = new StreamWriter(FileName + ".csv", false, System.Text.Encoding.UTF8))
            {
                foreach (Chromosome a in populations)
                {
                    int i = 0;
                    while (i < a.sequence_operation.Length)
                    {
                        sw.Write(a.sequence_operation[i] + ",");
                        i++;
                    }
                    i = 0;
                    while (i < a.machine_assignment.Length)
                    {
                        
                        sw.Write(a.machine_assignment[i] + ",");
                        i++;
                    }
                    sw.WriteLine(a.Fitness);
                }
            }
        }
    }
}
