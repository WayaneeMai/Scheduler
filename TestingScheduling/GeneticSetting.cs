using System;
using System.Collections.Generic;
using System.Text;

namespace TestingScheduling
{
    public class GeneticSetting
    {       
        public int num_generation;
        public int num_genes;//chromosome length.
        public int total_num_population;//number of solutions in the population.
        public int num_population_by_random=0;
        public int num_population_by_heuristics=0;
        public double stop_criteria_ratio;//improvement ratio
        public int not_improvement_time;
        public int selection_group_size;
        public double crossover_probability;
        public double mutation_probability;
        public int max_iteriation;
        public List<int> job { get; set; }
        public int[] number_operation;
        public List<Job_Operation_Index> job_operation_index;

        private ObjectiveFunction objective_function;
        public InitialSolution initial_method;

        public void SetObjectiveFunction(ObjectiveFunction Objective)
        {
            objective_function = Objective;
        }

        public ObjectiveFunction GetObjectiveFunction()
        {
            return objective_function;
        }

        public enum ObjectiveFunction
        {
            Makespan,
            TotalWeightedTardiness
        }

        public enum InitialSolution
        {
            ShortestProcess_Setup,
            ReleaseTime
        }
    }
}
