using Spire.Xls;
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

namespace XlsToXls
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.MaximizeBox = false;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            if (File.Exists(Application.StartupPath + @"\Remark.txt"))
            {
                using (StreamReader sr = new StreamReader(Application.StartupPath + @"\Remark.txt"))
                {
                    txt_remark.Text = sr.ReadToEnd();
                }
            }
        }

        private void btn_run_Click(object sender, EventArgs e)
        {
            Thread mission = new Thread(Run);
            mission.Start();
        }
        private void Run()
        {
            this.Invoke((MethodInvoker)delegate ()
            {
                label3.ForeColor = Color.Red;
                label3.Text = "執行中...";
            });
            // Before catefiltergory
            List<string> RawData = new List<string>();
            // Be written XLS
            List<string> intoData = new List<string>();
            //After filter
            List<string> Data = new List<string>();
            List<string> SortedData = new List<string>();

            XlsReader xlsReader = new XlsReader();
            string FilePath = "";
            string IntoFilePath = "";
            string target_ = "";

            this.Invoke((MethodInvoker)delegate ()
            {
                target_ = target.Text;
                FilePath = Application.StartupPath + $@"\Sample\{sample.Text}.xls";
                IntoFilePath = Application.StartupPath + $@"\Sample\{target.Text}.xls";
            });


            // Get xls raw data List<string>
            RawData = xlsReader.ReadExcel(FilePath);
            intoData = xlsReader.ReadExcel(IntoFilePath);
            // Get sending data
            foreach (var item in RawData)
            {
                string[] dataSplit = item.Split('_');
                if (!string.IsNullOrEmpty(dataSplit[11]))
                {
                    // comanpy name
                    Data.Add(dataSplit[6] + "_" + dataSplit[11]);
                }
            }
            foreach (var item in intoData)
            {
                string[] dataSplit = item.Split('_');
                if (!string.IsNullOrEmpty(dataSplit[6]) && !string.IsNullOrEmpty(dataSplit[9]) && !string.IsNullOrEmpty(dataSplit[10]))
                {
                    // comanpy name
                    SortedData.Add(dataSplit[6] + "_" + dataSplit[dataSplit.Length - 1]);
                }
            }

            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(IntoFilePath);
                Worksheet sheet = workbook.Worksheets[0];
                foreach (var item in Data)
                {
                    string[] listedValue = item.Split('_');

                    foreach (var ss in SortedData)
                    {
                        string[] sortSplit = ss.Split('_');

                        if (listedValue[0].Equals(sortSplit[0]))
                        {
                            InsertXLS(listedValue[1], $"L{sortSplit[1]}", workbook, sheet, target_);
                        }
                    }
                }
                Console.WriteLine("done");
            }
            catch (Exception ee)
            {
                this.Invoke((MethodInvoker)delegate ()
                {
                    label3.ForeColor = Color.Red;
                    label3.Text = ee.Message;
                });
            }
            this.Invoke((MethodInvoker)delegate ()
            {
                label3.ForeColor = Color.DarkGreen;
                label3.Text = "執行成功!";
            });
        }
        private void InsertXLS(string addStr, string columnNumber, Workbook workbook, Worksheet sheet, string filename)
        {
            //設定文字
            string comment = addStr;

            //設定字型
            ExcelFont font = workbook.CreateFont();
            font.FontName = "Calibri";
            font.Color = Color.Black;

            //新增到指定單元格
            Spire.Xls.CellRange range = sheet.Range[columnNumber];
            range.Text = comment;

            //儲存文件
            workbook.SaveToFile(Application.StartupPath + $"/Sample/{filename}.xls", ExcelVersion.Version97to2003);
            //System.Diagnostics.Process.Start("AddComment.xlsx");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            using (StreamWriter sw = new StreamWriter(Application.StartupPath + @"\Remark.txt"))
            {
                sw.WriteLine(txt_remark.Text);
            }
        }
    }
}
