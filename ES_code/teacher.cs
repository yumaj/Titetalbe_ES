using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.IO;




namespace ES_code
{
    class teacher
    {
        private int teachernum;
        private int classnum;
        private int score;
        private int classroomnum;
        private string[] classroomname;
        private string[][] data;
        private string[][] errordata;
        private string[][] subclass;
        private string[][] checkdata;
        private string[][] TimetalbOfclassroom;
        private string[] teachername;
        public teacher()
        {
            teachernum = 15;
            classnum = 45;
            score = 0;
            classroomnum = 6;
            data = new string[teachernum][];
            errordata = new string[teachernum][];
            subclass = new string[teachernum][];
            checkdata = new string[teachernum][];
            teachername = new string[teachernum];
            classroomname = new string[classroomnum];
            for (int i =0; i < teachernum;i++)
            {
                data[i] = new string[classnum];
                errordata[i] = new string[classnum];
                subclass[i] = new string[classnum];
                checkdata[i] = new string[classnum];


            }

            

        }
        public teacher(teacher or)
        {
            teachernum = 15;
            classnum = 45;
            score = 0;

            data = new string[teachernum][];
            errordata = new string[teachernum][];
            subclass = new string[teachernum][];
            checkdata = new string[teachernum][];
            teachername = new string[teachernum];
            for (int i = 0; i < teachernum; i++)
            {
                data[i] = new string[classnum];
                errordata[i] = new string[classnum];
                subclass[i] = new string[classnum];
                checkdata[i] = new string[classnum];


            }

            for(int i = 0; i < teachernum; i++)
            {
                for(int j = 0; j < classnum; j++)
                {
                    data[i][j] = string.Copy( or.get_data(i,j) );
                }
            }


        }
        public void output_csv(int times)
        {


            //the result of teacher
            string filePath = "output/teacher_"+ DateTime.Now.ToString("MMddyyyyHHmm") +"_result_" + times.ToString() + ".csv";
            FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sr = new StreamWriter(fs, System.Text.Encoding.UTF8);


            sr.Write(score.ToString() + ",\n");
            for (int i = 0; i < teachernum; i++)
            {
                sr.Write(teachername[i] + ",");
                for (int j = 0; j < classnum; j++)
                {
                    sr.Write(data[i][j] + ",");
                }
                sr.Write("\n");
            }
            sr.Close();
            //the result for student 
            string filePath2 = "output/grade_result_" + DateTime.Now.ToString("MMddyyyyHHmm")+"_"+ times.ToString() + ".csv";
            FileStream fs2 = new FileStream(filePath2, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sr2 = new StreamWriter(fs2, System.Text.Encoding.UTF8);


            string[][][] each_degree;
            each_degree = new string[4][][];

            for (int i = 0; i < 4; i++)
            {
                // each group
                each_degree[i] = new string[4][];

                for (int j = 0; j < 4; j++)
                {
                    each_degree[i][j] = new string[classnum];
                    for(int k = 0; k < classnum;k++)
                    {
                        each_degree[i][j][k] = "";
                    }
                }

                
            }


            for (int i = 0; i < teachernum; i++)
            {
                for (int j = 0; j < classnum; j++)
                {
                    if (data[i][j] != null && data[i][j] != "" && data[i][j] != "x")
                    {
                       each_degree[System.Convert.ToInt16(data[i][j][0].ToString()) - 1][System.Convert.ToInt16(data[i][j][6].ToString()) - 1][j] = data[i][j];
                    }
                }
            }

            string[] groupname = {"甲", "乙", "丙", "丁"};
            for (int i = 0; i < 4; i++)
            {
                sr2.Write("大" + (i + 1).ToString());
                for (int j  = 0; j < 4; j++)
                {
                    sr2.Write(groupname[j]  + "班\n");
                    sr2.Write("第1節,第2節,第3節,第4節,第5節,第6節,第7節,第8節,第9節,\n");
                    for (int k = 0; k < classnum; k++)
                    {
                        sr2.Write(each_degree[i][j][k] + ",");
                        if (k % 9 == 8)
                        {
                            sr2.Write("\n");
                        }

                    }
                }

            }
            sr2.Close();
            //error result 
            string filePath3 = "output/error_result_" + DateTime.Now.ToString("MMddyyyyHHmm") +"_"+times.ToString() + ".csv";
            FileStream fs3 = new FileStream(filePath3, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sr3 = new StreamWriter(fs3, System.Text.Encoding.UTF8);



            for (int i = 0; i < teachernum; i++)
            {
                sr3.Write(teachername[i] + ",");
                for (int j = 0; j < classnum; j++)
                {
                    sr3.Write(errordata[i][j] + ",");
                }
                sr3.Write("\n");
            }

            sr3.Close();

        }

        public string[] get_sch(int i)
        {
            return data[i];
        }
        //set function
        public void set_teachernum(int num)
        {
            teachernum = num;
        }
        public void set_classnum(int num)
        {
            classnum = num;
        }
        public void set_score(int num)
        {
            score = num;
        }
        public void set_data(int i,int j,string str)
        {
            data[i][j] = string.Copy( str);
        }
        public void set_errordata(int i, int j , string str)
        {
            errordata[i][j] = string.Copy(str);
        }
        public void set_subclass(int i,int j , string str)
        {
            subclass[i][j] = string.Copy(str);
        }
        public void set_checkdata(int i, int j , string str)
        {
            checkdata[i][j] = string.Copy(str);
        }
        //get function
        public int get_teachernum()
        {
            return teachernum;
        }
        public int get_classnum()
        {
            return classnum;
        }
        public int get_score()
        {
            return score;
        }
        public string get_data(int i, int j)
        {
            return data[i][j];
        }
        
        public string get_errordata(int i, int j)
        {
            return errordata[i][j];
        }
        public string get_subclass(int i, int j)
        {
            return subclass[i][j];
        }
        public string get_checkdata(int i, int j)
        {
            return checkdata[i][j];
        }

        public void printdata()
        {
            for(int i = 0;i < teachernum; i++)
            {
                for(int j = 0;j < classnum; j++)
                {
                    //System.Console.Write(data[i][j]);
                   // if (j % 9 == 0) System.Console.WriteLine();
                }
               // System.Console.WriteLine();
            }
        }
        public void score_clear()
        {
            score = 0;
        }
        public double  crossover_num()
        { 
            Random random = new Random();
           
            return random.NextDouble();
        }
        public void crossover(int num,string[] str)
        {
            for(int i = 0; i < str.Length; i++)
            {
                data[num][i] = string.Copy(str[i]);
            }
            
        }
        public void mute(double rate)
        {
            Random random = new Random();


            for(int i = 0; i < teachernum; i++)
            {
                for(int j = 0; j < classnum; j++)
                {
                    if(rate < random.NextDouble())
                    {
                        int randindex = random.Next(0,classnum);
                        if(data[i][j] == "x" || data[i][randindex] == "x")
                        {
                            continue;
                        }
                        if (data[i][j] == "" && data[i][randindex] == "")
                        {
                            continue;
                        }
                        if (data[i][j] == "" && data[i][randindex] != "")
                        {
                            if(data[i][randindex][4] != '2')
                            {
                                string temp = string.Copy(data[i][j]);
                                data[i][j] = string.Copy( data[i][randindex]);
                                data[i][randindex] = string.Copy( temp);
                            }
                            

                        }
                        else if(data[i][randindex] == "" && data[i][j] != "")
                        {
                            if (data[i][j][4] != '2')
                            {
                                string temp = string.Copy(data[i][j]);
                                data[i][j] = string.Copy(data[i][randindex]);
                                data[i][randindex] = string.Copy(temp);
                               
                            }
                          
                        }
                        else 
                        {
                            if (data[i][randindex][4] != '2' && data[i][j][4] != '2')
                            {
                                string temp = string.Copy(data[i][j]);
                                data[i][j] = string.Copy(data[i][randindex]);
                                data[i][randindex] = string.Copy(temp);
                              
                            }
                           

                        }
                    }
                }
            }




        }

        
        
        
        
        
        //fitnessfunciton 

        private int fitnessfunction1()//rule 1
        {
            int s = 0;
            for(int i = 0; i < teachernum; i++)
            {
                for(int j = 4; j < 8; j++)
                {
                    if(data[i][j] != ""  && data[i][j] != "x" )
                    {
                        s += 400;
                        errordata[i][j] += "1,";
                    }
                }
            }
            return s;
        }

        private int fitnessfunction2()// rule 8
        {
            int s = 0;
            for(int j = 0; j < classnum; j++)
            {
                bool[][] used;
                used = new bool[4][];
                for(int i = 0; i < 4; i++)
                {
                    used[i] = new bool[4];
                }
                for(int k = 0; k < 4; k++)
                {
                    for(int i = 0; i < 4; i++)
                    {
                        used[k][i] = false;
                    }
                }
                for(int i = 0; i < teachernum; i++)
                {
                    if(data[i][j] != "" && data[i][j] != "x")
                    {
                        if (data[i][j][2] == '1' || data[i][j][2] == '2')
                        {
                            int g = Convert.ToInt16(data[i][j][0].ToString());
                            int c = Convert.ToInt16(data[i][j][6].ToString());
                            if (used[g - 1][c - 1] == true)
                            {
                                s += 200;
                                errordata[i][j] += "2,";
                                errordata[i][j] += data[i][j];

                            }
                            else
                            {
                                used[g - 1][c - 1] = true;
                            }

                        }
                    }

                }
            }

            return s;
        }
        private int fitnessfunction3()//rule 10
        {
            int s = 0;
            int mcounter = 0;
            int ecounter = 0;
            for(int i = 0; i < teachernum; i++)
            {
                for(int j = 0; j < classnum; j++)
                {
                    if(j % 9 == 0)
                    {
                        mcounter = 0;
                        ecounter = 0;
                    }
                    if (data[i][j] != "" && data[i][j]!= "x")
                    {
                        mcounter += 1;
                        if (mcounter > 4)
                        {
                            errordata[i][j] += "3,";
                            s += 200;
                        }
                    }
                }
                

            }
            return s;
        }
        private int fitnessfunction4()//rule 12
        {
            int s = 0;
            for(int j = 0; j < classnum; j++)
            {
                bool[] used;
                used = new bool[4];
                for(int i = 0; i < 4; i++)
                {
                    used[i] = false;
                }
                for(int i = 0; i < teachernum; i++)
                {
                    if(data[i][j] != "" && data[i][j] != "x")
                    {
                        if (data[i][j][2] == '2')
                        {
                            int g = Convert.ToInt16(data[i][j][0]) - 49;// - 51;
                            if (used[g] == true)
                            {
                                s += 100;
                                errordata[i][j] += "4,";

                            }
                            else
                            {
                                used[g] = true;
                            }
                        }
                    }

                }
                
            }

            return s;
        }
        private int fitnessfunction5()//rule 13
        {
            int s = 0;
            for(int i = 0; i < teachernum; i++)
            {
                int daycounter = 0;
                bool classcounter = false;
                for(int j = 0; j < classnum; j++)
                {
                    if(data[i][j] != "" && data[i][j] != "x" )
                    {
                        classcounter = true;
                    }
                    if(j % 9 == 8)
                    {
                        if(classcounter == true)
                        {
                            daycounter += 1;
                            classcounter = false;
                        }
                    }


                }

                if (daycounter == 5)
                {
                    errordata[i][44] += "55,";
                    s += 50;
                }
                else if (daycounter == 4)
                {
                    errordata[i][44] += "45,";
                    s += 20;
                }

            }



            return s;
        }
        private int fitnessfunction6()//rule 14
        {
            int s = 0;
            int mcounter = 0;
            int acounter = 0;
            for(int i = 0; i < teachernum; i++)
            {
                for(int j = 0; j < classnum; j++)
                {
                    if( j % 9 == 0)
                    {
                        mcounter = 0;
                        acounter = 0;
                    }
                    if(data[i][j] != "" && data[i][j] != "x" && (j % 9 ) < 4)
                    {
                        mcounter += 1;
                    }
                    if(data[i][j] != "" && data[i][j] != "x" && (j % 9) >= 4)
                    {
                        acounter += 1;

                    }
                    if(mcounter + acounter >= 2 && j % 9  == 8)//fix the m + a number 
                    {
                        if(mcounter == 0 || acounter == 0)
                        {
                            continue;
                        }
                        /*
                        else if(Math.Abs(mcounter - acounter) >= m)
                        {
                            errordata[i][j] += "6,";
                            s += 10;
                        }
                        else if(Math.Abs(mcounter - acounter) <= 1 && j % 9 == 8)
                        {
                            errordata[i][j] += "6,";
                            s += 20;
                        }
                        */
                        else
                        {
                            errordata[i][j] += "6,";
                            s += 20;
                        }
                    }
                }
            }
            return s;
        }
        private int fitnessfunction7()
        {
            int s = 0;
            int now_classc_num = 0;
            int now_acces_classc_num = 0;
            for(int i = 0; i < teachernum; i++)
            {
                string[] classc = new string[1000];
                int[] acces_classc = new int[1000];
                for(int j = 0; j < classnum; j++)
                {
                    if( data[i][j] != "x" && data[i][j] != null && data[i][j] != "" && j % 9 != 3 && j % 9 != 8)
                    {
                        bool us = false;
                        for(int k = 0; k < now_classc_num;k++)
                        {
                            if (classc[k] == data[i][j]) us = true;
                        }
                        if(us == false)
                        {
                            classc[now_classc_num++] = data[i][j];

                            if(data[i][j] != data[i][j + 1])
                            {
                                acces_classc[now_acces_classc_num++] = 1;
                            }
                            else
                            {
                                acces_classc[now_acces_classc_num++] = 0;
                            }

                        }

                        if(us == true)
                        {
                            for(int k = 0; k < now_classc_num; k++)
                            {
                                if(classc[k] == data[i][j])
                                {
                                    if(acces_classc[k] == 1 && data[i][j] != data[i][j + 1])
                                    {
                                        s += 1000;
                                        checkdata[i][j] += "71, ";
                                        errordata[i][j] += "71, ";
                                    }
                                    else
                                    {
                                        acces_classc[k] = 0;
                                    }
                                }
                            }
                        }
                    }
                    if (data[i][j] != "x" && data[i][j] != null && data[i][j] != "" && (j % 9 == 3 || j % 9 == 8))
                    {
                        bool us = false;
                        for (int k = 0; k < now_classc_num; k++)
                        {
                            if (classc[k] == data[i][j]) us = true;
                        }
                        if (us == false)
                        {
                            classc[now_classc_num++] = data[i][j];

                            if (data[i][j] != data[i][j - 1])
                            {
                                acces_classc[now_acces_classc_num++] = 1;
                            }   
                            else
                            {
                                acces_classc[now_acces_classc_num++] = 0;
                            }

                        }

                        if (us == true)
                        {
                            for (int k = 0; k < now_classc_num; k++)
                            {
                                if (classc[k] == data[i][j])
                                {
                                    if (acces_classc[k] == 1 && data[i][j] != data[i][j - 1])
                                    {
                                        //change 200 to 400
                                        s += 1000;
                                        checkdata[i][j] += "72, ";
                                        errordata[i][j] += "72, ";
                                    }
                                    else
                                    {
                                        acces_classc[k] = 0;
                                    }
                                }
                            }
                        }
                    }
                }


            }


            return s;
        }
        private int fitnessfunction8()
        {
            int s = 0;
            for(int i = 0; i < teachernum; i++)
            {
                for (int j = 0; j < classnum; j++)
                {
                    if (data[i][j] != "" && data[i][j] != "x" && data[i][j] != null)
                    {
                        if (Convert.ToInt16(data[i][j][0].ToString()) < 4)
                        {
                            if (subclass[Convert.ToInt16(data[i][j][0].ToString()) - 1][j] != "" && subclass[Convert.ToInt16(data[i][j][0].ToString()) - 1][j] != null && subclass[Convert.ToInt16(data[i][j][0].ToString()) - 1][j][Convert.ToInt16(data[i][j][6].ToString()) * 2 - 2] == '1')
                            {
                                s += 200;
                                errordata[i][j] += ",8";
                            }
                        }

                    }

                }
            }
            return s;
        }
        private int fitnessfunction9()
        {
            int s = 0;
            int next_m = 1;
            int next_af = 5;
            for (int i = 0; i < teachernum;i ++)
            {
                for(int j = 0; j < classnum; j++)
                {
                    if(j == next_m && data[i][j] != "" && data[i][j] != "x" && data[i][j] != null)
                    {
                        if(data[i][j] == data[i][j + 1] && data[i][j + 1] == data[i][j + 2])
                        {
                            s += 200;
                        }
                        next_m += 8;
                    }
                }
                for (int j = 0; j < classnum; j++)
                {
                    if (j == next_af && data[i][j] != "" && data[i][j] != "x" && data[i][j] != null)
                    {
                        if (data[i][j] == data[i][j + 1] && data[i][j + 1] == data[i][j + 2])
                        {
                            s += 200;
                        }
                        next_af += 8;
                    }
                }
            }

            return s;
        }
        private int fitnessfunction10()
        {
            
            int s = 0;
            int now_classc_num = 0;
            int now_acces_classc_num = 0;
            for (int i = 0; i < teachernum; i++)
            {
                string[] classc = new string[1000];
                int[] acces_classc = new int[1000];
                for (int j = 0; j < classnum; j++)
                {
                    if (j % 9 == 0)
                    {
                        now_acces_classc_num = 0;
                        now_classc_num = 0;
                    }
                    if (data[i][j]!= "x" && data[i][j] != null && data[i][j] != "" && Convert.ToInt16(data[i][j][4].ToString()) != 2)
                    {
                        bool us = false;
                        for (int k = 0; k < now_classc_num; k++)
                        {
                            if (classc[k] == data[i][j]) us = true;
                        }
                        if (us == false)
                        {
                            classc[now_classc_num++] = data[i][j];
                            acces_classc[now_acces_classc_num++] = 0;


                        }

                        if (us == true)
                        {
                            for (int k = 0; k < now_classc_num; k++)
                            {
                                if (classc[k] == data[i][j])
                                {
                                    acces_classc[k] += 1;

                                    if (acces_classc[k] == 3)//????  && i == 10
                                    {
                                        s += 150;

                                        errordata[i][j] += "10, ";
                                    }

                                }
                            }
                        }
                    }
                }


            }


            return s;
        }
        private int fitnessfunction11()// do not put all same class in one day
        {
            int s = 0;
            Dictionary<string, int> Class_dic = new Dictionary<string, int>( );
            for (int i = 0; i < teachernum; i++)
            {
                for (int j = 0; j < classnum; j++)
                {

                    if (data[i][j] != "x" && data[i][j] != null && data[i][j] != "")
                    {

                        if (true == (Class_dic.ContainsKey(data[i][j])))
                        {
                            Class_dic[data[i][j]] ++;
                        }
                        else
                        {
                            Class_dic.Add(data[i][j],1);
                        }

                    }
                    if (j % 9 == 8)
                    {

                        foreach (KeyValuePair<string, int> item in Class_dic)
                        {
                            if (item.Value >= 3)
                            {
                                s += 30;
                                errordata[i][j] += "11, ";
                            }
                        }
                        Class_dic = null;
                        Class_dic = new Dictionary<string, int>();

                    }
                }
            }

            return s;
        }
        public int getscore()
        {
            score = 0;
            subclass = new string[teachernum][];
            checkdata = new string[teachernum][];

            for (int i = 0; i < teachernum; i++)
            {
                subclass[i] = new string[classnum];
                checkdata[i] = new string[classnum];
            }

            score += fitnessfunction1();
            score += fitnessfunction2();
            score += fitnessfunction3();
            score += fitnessfunction4();
            score += fitnessfunction5();
            score += fitnessfunction6();
            score += fitnessfunction7();
            score += fitnessfunction8();
            score += fitnessfunction9();
            score += fitnessfunction10();
            score += fitnessfunction11();
            return score;
        }
        /*
        public load_unable()
        {
            for(int i = 0; i < teachernum;i++)
            {
                for(int j = 0; j < classnum; j++)
                {
                    if(data[i][j] != null)
                    {
                        if(data[i][j][4] == '2' || data[i][j][0] == 'x')
                        {
                            
                        }
                    }
                }
            }
        }
        */
        public void loadsch(string filepath)
        {

        }


        public void OpenExcel(string strFileName)
        {
            object missing = System.Reflection.Missing.Value;
            Application excel = new Application();//lauch excel application
            
            if (excel == null)
            {
                Console.Out.Write("Can't access excel");
            }
            else
            {
                excel.Visible = false; excel.UserControl = true;
                // 以只讀的形式打開EXCEL文檔
                Workbook wb = excel.Application.Workbooks.Open(strFileName, missing, true, missing, missing, missing,
                 missing, missing, missing, true, missing, missing, missing, missing, missing);
                //取得第一個工作薄
                Worksheet ws = (Worksheet)wb.Worksheets.get_Item(1);

                Excel.Range xlRange = ws.UsedRange;
                //Console.Out.WriteLine("Row = " + ws.UsedRange.Rows.Count);
                //Console.Out.WriteLine("Col = " + ws.UsedRange.Columns.Count);
                teachernum = Convert.ToInt16( xlRange.Cells[1, 1].Value2 );

                data = new string[teachernum][];
                errordata = new string[teachernum][];
                subclass = new string[teachernum][];
                checkdata = new string[teachernum][];
                teachername = new string[teachernum];
                for (int i = 0; i < teachernum; i++)
                {
                    data[i] = new string[classnum];
                    errordata[i] = new string[classnum];
                    subclass[i] = new string[classnum];
                    checkdata[i] = new string[classnum];


                }


                string[][] sa = new string[ws.UsedRange.Rows.Count][];
                for(int i = 0;i < ws.UsedRange.Rows.Count;i++)
                {
                    sa[i] = new string[ws.UsedRange.Columns.Count];
                }
                for (int i = 1; i <= ws.UsedRange.Rows.Count; i++)
                {
                    for (int j = 1; j <= ws.UsedRange.Columns.Count; j++)
                    {
                        //new line
                        //if (j == 1)
                         //   Console.Write("\r\n");

                        //write the value to the console
                       // if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        //    Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

                        //add useful things here!   
                        if( i > 1)
                        {
                            if (xlRange.Cells[i, j].Value2 == null) sa[i - 2][j - 1] = "";
                            else sa[i - 2][j - 1] = xlRange.Cells[i, j].Value2.ToString();
                        }
                         
                        
                    }
                }

                for(int i = 0; i < teachernum; i++)
                {
                    teachername[i] = sa[i][0];
                    for(int j = 1 ; j <= classnum; j++)
                    {
                        string class_string = sa[i][j];
                        data[i][j - 1] = class_string;
                    }

                }
                for(int i = 0; i < 3; i++)
                {
                    for(int j = 1; j <= classnum; j++)
                    {
                          subclass[i][j - 1] = sa[teachernum + i][j];
                    }
                }

            }
            excel.Quit(); excel = null;
          
            GC.Collect();
        }

    }
}
