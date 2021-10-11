using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using Syncfusion.XlsIO;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Excel_To_Json_Converter_WinForms_
{
    public partial class Form1 : Form
    {
        dynamic openfilelist, XLFileName, XLSaveFilename, Fileformat, FileName, XLFileSavepath, Newfilename;

        

        public Form1()
        {
            InitializeComponent();
        }

        private void Openbut_Click(object sender, EventArgs e)
        {
            OpenFileDialog opfd = new OpenFileDialog { Multiselect = true };

            if(opfd.ShowDialog()==DialogResult.OK)
            {
                openfilelist = opfd.FileNames;

                foreach(string item in openfilelist)
                {
                    InputListBox.Items.Add(item);
                }
                XLFileName = opfd.SafeFileNames;

            }

            FolderBrowserDialog fobd = new FolderBrowserDialog();
          
            if(fobd.ShowDialog()==DialogResult.OK)
            {
                XLSaveFilename = XLFileName;

                foreach(string item in XLSaveFilename)
                {
                    XLFileSavepath = fobd.SelectedPath;

                   // FileName = item;

                    FileName = item.Substring(0, item.Length - 5);

                    XLFileSavepath += ("\\" + FileName + ".json");

                    OutputListBox.Items.Add(XLFileSavepath);

                }
            }
        }
        private void Convert_Click(object sender, EventArgs e)
        {
            foreach (string openpath in openfilelist)
            {

                var outputdirectory = Path.GetDirectoryName(XLFileSavepath);

                var Fileformat = Path.GetFileName(openpath).Replace(".xlsx", ".json");

                Newfilename = Path.Combine(outputdirectory, Fileformat);

                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;

                    //The workbook is opened.
                    FileStream fileStream = new FileStream(openpath, FileMode.Open);
               //     FileStream fileStream = new FileStream()


                    IWorkbook workbook = application.Workbooks.Open(fileStream, ExcelOpenType.Automatic);
                    IWorksheet worksheet = workbook.Worksheets[0];

                    //Export worksheet data into CLR Objects
                    IList<ExcelData> Xldata = worksheet.ExportData<ExcelData>(1, 1, worksheet.UsedRange.LastRow, workbook.Worksheets[0].UsedRange.LastColumn);

                    //  IList<Customer> customers = worksheet.ExportData<Customer>(1, 1, worksheet.UsedRange.LastRow, workbook.Worksheets[0].UsedRange.LastColumn);


                    //open file stream
                    using (StreamWriter file = File.CreateText(Newfilename))
                    {
                        JsonSerializer serializer = new JsonSerializer();

                        //serialize object directly into file stream
                        serializer.Serialize(file, Xldata);
                    }
                }


                
            }

            MessageBox.Show("Process Completed Successfully");
            
          
        }
    }
}
