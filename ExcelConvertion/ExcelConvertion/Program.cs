using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Windows.Forms;

namespace ExcelConvertion
{
    class Program
    {
        static void Main(string[] args)
        {

            // For Friends Series Enable this Line

            string InputNameLines = System.IO.File.ReadAllText(@"D:\Project\delete\New Text Document1.txt").ToString();
            string[] test = InputNameLines.Split(new string[] { "\n\n" }, StringSplitOptions.None);


            // For Movies Srt to enable this line

            //string InputNameLines = System.IO.File.ReadAllText(@"D:\Project\delete\La La Land.srt");
            //string[] test = InputNameLines.Split(new string[] { "\r" + "\n" + "\r" }, StringSplitOptions.RemoveEmptyEntries);


            excel.Application oxl;
            excel.Workbook owb;
            excel.Worksheet osheet;
            excel.Range orng;

            oxl = new excel.Application();

            oxl.Visible = true;


            object misvalue = System.Reflection.Missing.Value;
            try
            {
                oxl = new excel.Application();
                oxl.Visible = true;

                owb = (excel.Workbook)(oxl.Workbooks.Add(""));
                osheet = (excel.Worksheet)owb.ActiveSheet;
                osheet.Cells[1, 1] = "FROM";
                osheet.Cells[1, 2] = "TO";
                osheet.Cells[1, 3] = "TEXT";
                osheet.Cells[1, 4] = "ACTOR";
                //osheet.get_Range("A1", "c1").Font.Bold = true;
               // osheet.get_Range("A1", "c1").VerticalAlignment = excel.XlVAlign.xlVAlignCenter;
                osheet.get_Range("A1", "c1").HorizontalAlignment = excel.XlVAlign.xlVAlignCenter;

                //  osheet.get_Range()


                //For Friends Series
                for (int i = 0; i <= test.Length - 1; i++)
                {
                    int k = i + 1;

                  string[] test1 = test[i].Split(new string[] { "\n" }, StringSplitOptions.None);
                  //  string[] test1 = test[i].Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);

                    if (test1[0] != string.Empty)
                    {
                        string[] test2 = test1[1].Split(new string[] { "--> " }, StringSplitOptions.None);

                      
                        dynamic val = 0;
                        dynamic val1 = 0;
                        //dynamic val2 = 0;
                        //dynamic val3 = 0;

                        for (dynamic m = 0; m <= test2.Length - 1; m++)
                        {
                            if (m == 0)
                            {
                            //   val= test2[m].Replace(",",":");

                                val = test2[m].Substring(0, test2[m].Length - 5);

                            }
                            if (m == 1)
                            {
                              //  val2 = test2[m].Replace(",", ":");
                                val1 = test2[m].Substring(0, test2[m].Length - 4);

                              //  val3 = val2.Substring(0, val2.Length - 5);
                            }

                        }
                        
                        
                            TimeSpan ts = TimeSpan.Parse(val);
                        var millisec = ts.TotalMilliseconds;
                        TimeSpan ts1 = TimeSpan.Parse(val1);
                        var millisec1 = ts1.TotalMilliseconds;

                        osheet.Cells[1][k + 1] = millisec;
                        osheet.Cells[2][k + 1] = millisec1;



                        //osheet.Cells[1][k + 1] = test2[0];
                        //osheet.Cells[2][k + 1] = test2[1];

                        string temp = string.Empty;
                        for (int j = 2; j <= test1.Length - 1; j++)
                        {
                            temp = temp + test1[j];
                        }
                        osheet.Cells[3][k + 1] = temp;
                    }
                }

                //For Movies                
                //for (int i = 0; i <= test.Length-1; i++)
                //{

                //    int k = i + 1;



                //    string[] test1 = test[i].Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);

                //    if (test1[0] != string.Empty)
                //    {
                //        string[] test2 = test1[1].Split(new string[] { "-->" }, StringSplitOptions.None);
                //        osheet.Cells[1][k + 1] = test2[0];
                //        osheet.Cells[2][k + 1] = test2[1];
                //        string temp = string.Empty;
                //        for (int j = 2; j <= test1.Length - 1; j++)
                //        {
                //            temp = temp + test1[j];
                //        }
                //        osheet.Cells[3][k + 1] = temp;
                //    }
                //}

                Thread.Sleep(5000);
                orng = osheet.get_Range("A1", "c1");
                orng.EntireColumn.AutoFit();
                oxl.Visible = false;
               

                //For Movies

                //owb.SaveAs(@"D:\Project\delete\Movies\Output3.xlsx", excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                //    false, false, excel.XlSaveAsAccessMode.xlNoChange,
                //    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //owb.Close();


                //For Friends Series

                owb.SaveAs(@"D:\Project\delete\Friends\Output3.xlsx", excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                   false, false, excel.XlSaveAsAccessMode.xlNoChange,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                owb.Close();
                
                oxl.Application.Quit();

                Console.WriteLine("Conversion Completed SuccessFully");

                Environment.Exit(0);
                

            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception" + ex);

                oxl.Application.Quit();

                Environment.Exit(0);
               
            }

        }

    }
}
