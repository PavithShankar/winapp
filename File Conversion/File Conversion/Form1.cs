using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using Syncfusion.XlsIO;
using Newtonsoft.Json;

namespace File_Conversion
{
    public partial class Form1 : Form
    {
        dynamic Filelist, FileListNames, SaveFilename, SampleSaveFileName, SaveFilepath, Newfilename;

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.ProgBar.Increment(10);
        }

        public Form1()
        {
            InitializeComponent();

        }

        private void Open_Click(object sender, EventArgs e)
        {
            OpenFileDialog opfd = new OpenFileDialog { Multiselect = true };

            if (opfd.ShowDialog() == DialogResult.OK)
            {
                Filelist = opfd.FileNames;

                foreach (string item in Filelist)
                {
                    OpenFileNames.Items.Add(item);


                }
                FileListNames = opfd.SafeFileNames;

            }

            FolderBrowserDialog fobd = new FolderBrowserDialog();

            if (fobd.ShowDialog() == DialogResult.OK)
            {
                SaveFilename = FileListNames;

                foreach (string item in SaveFilename)
                {
                    SaveFilepath = fobd.SelectedPath;

                    // FileName = item;

                    SampleSaveFileName = item.Substring(0, item.Length - 4);

                    SaveFilepath += ("\\" + SampleSaveFileName + ".xlsx");

                    FormatFileNames.Items.Add(SaveFilepath);

                }
            }
        }

        private void DramaToExcel_Click(object sender, EventArgs e)
        {
            this.timer1.Start();

            if (OpenFileNames.Items.Count==0)
            {

                MessageBox.Show("Plz Select the Srt files");

            }
            else
            {

                foreach (string openpath in Filelist)
                {

                    var outputdirectory = Path.GetDirectoryName(SaveFilepath);

                    var Fileformat = Path.GetFileName(openpath).Replace(".srt", ".xlsx");

                    Newfilename = Path.Combine(outputdirectory, Fileformat);


                    string InputNameLines = System.IO.File.ReadAllText(openpath).ToString();
                    string[] test = InputNameLines.Split(new string[] { "\n\n" }, StringSplitOptions.None);

                    excel.Application oxl;
                    excel.Workbook owb;
                    excel.Worksheet osheet;
                    excel.Range orng;

                    oxl = new excel.Application();

                    oxl.Visible = true;


                    object misvalue = System.Reflection.Missing.Value;
                    try
                    {
                        owb = (excel.Workbook)(oxl.Workbooks.Add(""));
                        osheet = (excel.Worksheet)owb.ActiveSheet;
                        osheet.Cells[1, 1] = "FROM";
                        osheet.Cells[1, 2] = "TO";
                        osheet.Cells[1, 3] = "TEXT";
                        osheet.Cells[1, 4] = "ACTOR";

                        osheet.get_Range("A1", "c1").HorizontalAlignment = excel.XlVAlign.xlVAlignCenter;


                        //For Friends Series
                        for (int i = 0; i <= test.Length - 1; i++)
                        {
                            int k = i + 1;

                            string[] test1 = test[i].Split(new string[] { "\n" }, StringSplitOptions.None);


                            if (test1[0] != string.Empty)
                            {
                                string[] test2 = test1[1].Split(new string[] { "--> " }, StringSplitOptions.None);


                                dynamic val = 0;
                                dynamic val1 = 0;


                                for (dynamic m = 0; m <= test2.Length - 1; m++)
                                {
                                    if (m == 0)
                                    {


                                        val = test2[m].Substring(0, test2[m].Length - 5);

                                    }
                                    if (m == 1)
                                    {

                                        val1 = test2[m].Substring(0, test2[m].Length - 4);


                                    }

                                }


                                TimeSpan ts = TimeSpan.Parse(val);
                                var millisec = ts.TotalMilliseconds;
                                TimeSpan ts1 = TimeSpan.Parse(val1);
                                var millisec1 = ts1.TotalMilliseconds;

                                osheet.Cells[1][k + 1] = millisec;
                                osheet.Cells[2][k + 1] = millisec1;


                                string temp = string.Empty;
                                for (int j = 2; j <= test1.Length - 1; j++)
                                {
                                    temp = temp + test1[j];
                                }
                                osheet.Cells[3][k + 1] = temp;
                            }
                        }



                        Thread.Sleep(5000);
                        orng = osheet.get_Range("A1", "c1");
                        orng.EntireColumn.AutoFit();
                        oxl.Visible = false;

                        owb.SaveAs(Newfilename, excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                           false, false, excel.XlSaveAsAccessMode.xlNoChange,
                           Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        owb.Close();

                        oxl.Application.Quit();

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Exception" + ex);

                        oxl.Application.Quit();

                        Environment.Exit(0);

                    }


                }
                MessageBox.Show("Conversion Completed Successfully");
                Environment.Exit(0);
            }

        }

        private void MoviesToExcel_Click(object sender, EventArgs e)
        {
            this.timer1.Start();

            if (OpenFileNames.Items.Count==0)
            {

                MessageBox.Show("Plz Select the Srt files");

            }

            else
            {


                foreach (string openpath in Filelist)
                {

                    var outputdirectory = Path.GetDirectoryName(SaveFilepath);

                    var Fileformat = Path.GetFileName(openpath).Replace(".srt", ".xlsx");

                    Newfilename = Path.Combine(outputdirectory, Fileformat);


                    string InputNameLines = System.IO.File.ReadAllText(openpath);
                    string[] test = InputNameLines.Split(new string[] { "\r" + "\n" + "\r" + "\n" }, StringSplitOptions.RemoveEmptyEntries);

                    excel.Application oxl;
                    excel.Workbook owb;
                    excel.Worksheet osheet;
                    excel.Range orng;

                    oxl = new excel.Application();

                    oxl.Visible = true;
                    owb = (excel.Workbook)(oxl.Workbooks.Add(""));
                    osheet = (excel.Worksheet)owb.ActiveSheet;

                    object misvalue = System.Reflection.Missing.Value;
                    try
                    {

                        osheet.Cells[1, 1] = "FROM";
                        osheet.Cells[1, 2] = "TO";
                        osheet.Cells[1, 3] = "TEXT";
                        osheet.Cells[1, 4] = "ACTOR";

                        osheet.get_Range("A1", "c1").HorizontalAlignment = excel.XlVAlign.xlVAlignCenter;



                        for (int i = 0; i <= test.Length - 1; i++)
                        {
                            int k = i + 1;


                            string[] test1 = test[i].Split(new string[] { "\n" }, StringSplitOptions.None);

                            if (test1[0] != string.Empty)
                            {
                                string[] test2 = test1[1].Split(new string[] { "--> " }, StringSplitOptions.None);


                                dynamic val = 0;
                                dynamic val1 = 0;

                                for (dynamic m = 0; m <= test2.Length - 1; m++)
                                {
                                    if (m == 0)
                                    {


                                        val = test2[m].Substring(0, test2[m].Length - 5);

                                    }
                                    if (m == 1)
                                    {

                                        val1 = test2[m].Substring(0, test2[m].Length - 5);

                                    }

                                }


                                TimeSpan ts = TimeSpan.Parse(val);
                                var millisec = ts.TotalMilliseconds;
                                TimeSpan ts1 = TimeSpan.Parse(val1);
                                var millisec1 = ts1.TotalMilliseconds;

                                osheet.Cells[1][k + 1] = millisec;
                                osheet.Cells[2][k + 1] = millisec1;

                                string temp = string.Empty;
                                for (int j = 2; j <= test1.Length - 1; j++)
                                {
                                    temp = temp + test1[j];
                                }
                                osheet.Cells[3][k + 1] = temp;
                            }
                        }

                        Thread.Sleep(5000);
                        orng = osheet.get_Range("A1", "c1");
                        orng.EntireColumn.AutoFit();
                        oxl.Visible = false;

                        owb.SaveAs(Newfilename, excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                        false, false, excel.XlSaveAsAccessMode.xlNoChange,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        owb.Close();

                        oxl.Application.Quit();

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Exception" + ex);

                        owb.Close();

                        oxl.Application.Quit();

                        Environment.Exit(0);

                    }

                }

                MessageBox.Show("Conversion Completed Successfully");
                Environment.Exit(0);

            }
        }

        private void ExcelToJson_Click(object sender, EventArgs e)
        {
            this.timer1.Start();
            if (OpenFileNames.Items.Count==0)
            {

                MessageBox.Show("Plz Select the Excel files");

            }
            
            else
            {

                foreach (string openpath in Filelist)
                {
                    
                    var outputdirectory = Path.GetDirectoryName(SaveFilepath);

                    var Fileformat = Path.GetFileName(openpath).Replace(".xlsx", ".json");

                    Newfilename = Path.Combine(outputdirectory, Fileformat);


                    using (ExcelEngine excelEngine = new ExcelEngine())
                    {
                        IApplication application = excelEngine.Excel;


                        FileStream fileStream = new FileStream(openpath, FileMode.Open);

                        IWorkbook workbook = application.Workbooks.Open(fileStream, ExcelOpenType.Automatic);
                        IWorksheet worksheet = workbook.Worksheets[0];

                        IList<ExcelData> Xldata = worksheet.ExportData<ExcelData>(1, 1, worksheet.UsedRange.LastRow, workbook.Worksheets[0].UsedRange.LastColumn);

                        using (StreamWriter file = File.CreateText(Newfilename))
                        {
                            JsonSerializer serializer = new JsonSerializer();


                            serializer.Serialize(file, Xldata);


                        }
                    }

                }

                MessageBox.Show("Process Completed Successfully");
                Environment.Exit(0);

            }
        }


    }
}


        



