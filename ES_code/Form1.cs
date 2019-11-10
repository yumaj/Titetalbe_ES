using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace ES_code
{
    public partial class Form1 : Form
    {
        teacher t;
        public Form1()
        {
           
            InitializeComponent();
            t = new teacher();
            

        }

        private void btnSelectPath_Click(object sender, EventArgs e)
        {


        }

        private void run_main()
        {

            int runtime = 10;
            int total = 0;
            int[] allt = new int[runtime];


                
            int populationnum = 1;
            int maxge = 300000;

            int[][] alllog = new int[runtime][];
            for (int i = 0; i < runtime; i++)
            {
                alllog[i] = new int[maxge];
            }


            int bestnum = 0;
            double crossoverrate = 1;
            double muterate = 0;
            double inidvi_muterate = 0.98;
            teacher[] population = new teacher[populationnum];
            for (int i = 0; i < populationnum; i++)
            {
                population[i] = new teacher();
            }
            //first init genertation 
            OpenFileDialog file = new OpenFileDialog();
            file.ShowDialog();
            //this.label6.Text = file.FileName;
            for (int forlo = 0; forlo < runtime; forlo++)
            {
                population[0].OpenExcel(file.FileName);
                for (int i = 0 ; i < populationnum; i++)
                {
                    population[i] = new teacher(population[0]);
                }
                for (int i = 0; i < populationnum; i++)
                {
                    population[i].getscore();

                }

                bool isbest = false;
                int gecounter = 0;

                int[] scorelist = new int[maxge];

                while (isbest == false)
                {
                    gecounter += 1;

                    //random selection two parents from old generation, generate two chlide, totoal generate N child
                    Random r = new Random();

                    teacher[] new_population = new teacher[populationnum];
                    
                   for (int i = 0; i < populationnum; i += 2)
                   {
                        // int rand_teacher = r.Next(0, population[0].get_teachernum());

                        //int parentsA = r.Next(0, populationnum);
                        //int parentsB = r.Next(0, populationnum);
                        teacher new_popA = new teacher(population[0]);
                        //teacher new_popA = new teacher(population[parentsA]);
                        //teacher new_popB = new teacher(population[parentsB]);

                        //crossover 
                        /*
                        if (r.NextDouble() > crossoverrate)
                       {
                           teacher temp = new_popA;
                           new_popA.crossover(rand_teacher, new_popB.get_sch(rand_teacher));
                           new_popB.crossover(rand_teacher, temp.get_sch(rand_teacher));
                       }
                       */
                       new_population[i] = new_popA;
                       //new_population[i + 1] = new_popB;

                   }
                   
                    //Child Mute
                    //for (int i = 0; i < populationnum; i++)
                    for (int i = 0; i < populationnum ; i++)
                    {
                        if (r.NextDouble() > muterate)
                        {
                            new_population[i].mute(inidvi_muterate);
                            //population[i].mute(inidvi_muterate);
                        }
                    }
                    //Select N population from old generation and child (top N) into next generation
                    for (int i = 0; i < populationnum; i++)
                    {
                        //new_population[i].getscore();
                        new_population[i].getscore();
                        //Console.WriteLine("pi = " + i.ToString() + "  "+ new_population[i].getscore().ToString());
                    }

                    teacher[] combine_population = new teacher[populationnum * 2];
                    
                    for (int i = 0; i < populationnum; i++)
                    {
                        combine_population[i] = population[i];
                    }
                    for (int i = populationnum; i < populationnum * 2; i++)
                    {
                        combine_population[i] = new_population[i - populationnum];
                    }
                    

                    for (int i = 0; i < populationnum * 2; i++)
                    {
                        combine_population[i].getscore();
                    }
                   


                    for (int i = 0; i < populationnum * 2; i++)
                    {
                        for (int j = i + 1; j < populationnum * 2; j++)
                        {
                            if (combine_population[i].get_score() > combine_population[j].get_score())
                            {
                                teacher temp = new teacher(combine_population[i]);
                                combine_population[i] = new teacher(combine_population[j]);
                                combine_population[j] = temp;
                            }
                        }
                    }
                     

                    for (int i = 0; i < populationnum; i++)
                    {
                        for (int j = i + 1; j < populationnum; j++)
                        {
                            if (combine_population[i].get_score() > combine_population[j].get_score())
                            {
                                teacher temp = new teacher(combine_population[i]);
                                combine_population[i] = new teacher(combine_population[j]);
                                combine_population[j] = temp;
                            }
                        }
                    }


                    for (int i = 0; i < populationnum; i++)
                    {
                        population[i] = new teacher(combine_population[i]);
                    }

                    int bestscore = population[0].getscore();

                    //find the best score (score = 0) or generationcounter = max_generationnumber, break the loop
                    alllog[forlo][gecounter -1] = bestscore;
                    if (maxge == gecounter || bestscore == bestnum)
                    {
                        isbest = true;
                        population[0].output_csv(1);
                    }
                    //label2.Text = gecounter.ToString();
                    //label3.Text = bestscore.ToString();
                    //Console.WriteLine("ga = " + gecounter.ToString());
                    //Console.WriteLine("bestscore = " + bestscore.ToString());
                    allt[forlo] = bestscore;
                }

                total += allt[forlo];
            }

            string filePath = "output/alllog_" + DateTime.Now.ToString("MMddyyyyHHmm") + "_result_" + ".csv";
            FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sr = new StreamWriter(fs, System.Text.Encoding.UTF8);

            for (int i = 0; i < maxge; i++)
            {
                for(int j = 0; j < runtime; j++)
                {
                    //Console.WriteLine("all = " + allt[i].ToString());
                    sr.Write(alllog[j][i] + ",");
                }
                sr.Write("\n");
            }
            sr.Close();
            //Console.WriteLine("avg = " + (total/runtime).ToString());
        }
        private void button1_Click(object sender, EventArgs e)
        {
            //start main function 

            Thread t1 = new Thread(run_main);
            t1.SetApartmentState(ApartmentState.STA);


            t1.Start();

        }

        private void button2_Click(object sender, EventArgs e)
        {

            OpenFileDialog file = new OpenFileDialog();
            file.ShowDialog();
            this.label6.Text = file.FileName;
            t.OpenExcel(file.FileName);
        }
    }










}
