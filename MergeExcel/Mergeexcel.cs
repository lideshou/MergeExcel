using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;

namespace MergeExcel
{
    public partial class Mergeexcel : Form
    {

        //本地打开文件夹路径
        public string localOpenPath;

        //本地存储路径
        public string localSavePath;

        public Mergeexcel()
        {
            InitializeComponent();
        }
        
        //打开文件按钮监听函数
        private void btnOpen_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                openPath.Text = folderBrowserDialog1.SelectedPath;
            }
            localOpenPath = folderBrowserDialog1.SelectedPath;
        }

        //保存本地按钮监听函数
        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel 工作簿（*.xlsx）|*.xlsx|Excel 启动宏的工作簿（*.xlsm）|*.xlsm|Excel 97-2003工作簿（*.xls）|*.xls";
            saveFileDialog1.FileName = "默认文件名";//设置默认文件名
            saveFileDialog1.RestoreDirectory = true;//保存对话框是否记忆上次打开的目录

            saveFileDialog1.CheckPathExists = true;//检查目录

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                savePath.Text = saveFileDialog1.FileName;
            }
            localSavePath = saveFileDialog1.FileName;//文件路径
        }

        //开始操作按钮监听函数
        private void btnOpreat_Click(object sender, EventArgs e)
        {
            FileInfo newFile = new FileInfo(localSavePath);
            if(newFile.Exists)
            {
                newFile.Delete();
            }
            using (var excel = new ExcelPackage(newFile))
            {
                var ws = excel.Workbook.Worksheets.Add("report");
                ws.Row(1).Height = 30;
                ws.Column(1).Width = 20;
                ws.Column(2).Width = 15;
                ws.Column(3).Width = 15;
                ws.Column(25).Width = 15;
                //单元格自动适应大小
                ws.Cells.Style.ShrinkToFit = true;
                //水平居中
                ws.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //垂直居中
                ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["A1:A3,B2:B3,C2:C3,Y2:Y3"].Merge = true;
                ws.Cells["D2:F2,G2:I2,J2:L2,M2:O2,P2:R2,S2:U2,V2:X2"].Merge = true;
                ws.Cells["Z2:AB2,AC2:AE2,AF2:AH2,AI2:AK2,AL2:AN2,AO2:AQ2,AR2:AT2"].Merge = true;
                ws.Cells["B1:X1"].Merge = true;
                ws.Cells["Y1:AT1"].Merge = true;
                ws.Cells["A1"].Value = "SN";
                ws.Cells["B1"].Value = "GSM 900";
                ws.Cells["Y1"].Value = "DCS 1800";
                ws.Cells["B2"].Value = "AFC Init DAC";
                ws.Cells["C2"].Value = "PCL5 APC DAC";
                ws.Cells["D2"].Value = "Output Power (dBm)";
                ws.Cells["G2"].Value = "Peak_Phase_Error (deg.)";
                ws.Cells["J2"].Value = "Frequency_Error (Hz)";
                ws.Cells["M2"].Value = "RMS_Phase_Error (deg.)";
                ws.Cells["P2"].Value = "ORFS_SW_-400 kHz (dBm)";
                ws.Cells["S2"].Value = "ORFS_SW_+400 kHz (dBm)";
                ws.Cells["V2"].Value = "BER @ -102 dBm RX Level (%)";
                ws.Cells["Y2"].Value = "PCL0 APC DAC";
                ws.Cells["Z2"].Value = "Output Power (dBm)";
                ws.Cells["AC2"].Value = "Peak_Phase_Error (deg.)";
                ws.Cells["AF2"].Value = "Frequency_Error (Hz)";
                ws.Cells["AI2"].Value = "RMS_Phase_Error (deg.)";
                ws.Cells["AL2"].Value = "ORFS_SW_-400 kHz (dBm)";
                ws.Cells["AO2"].Value = "ORFS_SW_+400 kHz (dBm)";
                ws.Cells["AR2"].Value = "BER @ -102 dBm RX Level (%)";
                for (int i = 4; i <= 46; i++)
                {
                    if (i < 26)
                    {
                        switch (i % 3)
                        {
                            case 1: ws.Cells[3, i].Value = "CH975"; break;
                            case 2: ws.Cells[3, i].Value = "CH62"; break;
                            case 0: ws.Cells[3, i].Value = "CH124"; break;
                        }
                    }
                    else
                    {
                        switch (i % 3)
                        {
                            case 1: ws.Cells[3, i].Value = "CH885"; break;
                            case 2: ws.Cells[3, i].Value = "CH512"; break;
                            case 0: ws.Cells[3, i].Value = "CH698"; break;
                        }
                    }
                }

                DirectoryInfo dir = new DirectoryInfo(localOpenPath);
                FileInfo[] fileInfo = dir.GetFiles();
                List<string> fileNames = new List<string>();
                foreach (FileInfo item in fileInfo)
                {
                    fileNames.Add(item.FullName);
                }
                for (int i = 0; i < fileNames.Count; i++)
                {
                    OperateTable ot = new OperateTable();
                    ot.strPath = fileNames[i];
                    ws.Cells[i + 4, 1].Value = ot.getSN().Substring(0, 12);
                    ws.Cells[i + 4, 2].Value = ot.getAFC_Init_DAC();
                    List<string> Pad = ot.getPCL5_APC_DAC();
                    ws.Cells[i + 4, 3].Value = Pad[0];
                    ws.Cells[i + 4, 25].Value = Pad[1];
                    Pad.Clear();
                    List<string> ch975 = ot.getData("975");
                    ws.Cells[i + 4, 4].Value = ch975[0];
                    ws.Cells[i + 4, 7].Value = ch975[1];
                    ws.Cells[i + 4, 10].Value = ch975[2];
                    ws.Cells[i + 4, 13].Value = ch975[3];
                    ws.Cells[i + 4, 16].Value = ch975[4];
                    ws.Cells[i + 4, 19].Value = ch975[5];
                    ch975.Clear();
                    List<string> ch62 = ot.getData("62");
                    ws.Cells[i + 4, 5].Value = ch62[0];
                    ws.Cells[i + 4, 8].Value = ch62[1];
                    ws.Cells[i + 4, 11].Value = ch62[2];
                    ws.Cells[i + 4, 14].Value = ch62[3];
                    ws.Cells[i + 4, 17].Value = ch62[4];
                    ws.Cells[i + 4, 20].Value = ch62[5];
                    ch62.Clear();
                    List<string> ch124 = ot.getData("124");
                    ws.Cells[i + 4, 6].Value = ch124[0];
                    ws.Cells[i + 4, 9].Value = ch124[1];
                    ws.Cells[i + 4, 12].Value = ch124[2];
                    ws.Cells[i + 4, 15].Value = ch124[3];
                    ws.Cells[i + 4, 18].Value = ch124[4];
                    ws.Cells[i + 4, 21].Value = ch124[5];
                    ch124.Clear();
                    List<string> ch512 = ot.getData("512");
                    ws.Cells[i + 4, 26].Value = ch512[0];
                    ws.Cells[i + 4, 29].Value = ch512[1];
                    ws.Cells[i + 4, 32].Value = ch512[2];
                    ws.Cells[i + 4, 35].Value = ch512[3];
                    ws.Cells[i + 4, 38].Value = ch512[4];
                    ws.Cells[i + 4, 41].Value = ch512[5];
                    ch512.Clear();
                    List<string> ch698 = ot.getData("698");
                    ws.Cells[i + 4, 27].Value = ch698[0];
                    ws.Cells[i + 4, 30].Value = ch698[1];
                    ws.Cells[i + 4, 33].Value = ch698[2];
                    ws.Cells[i + 4, 36].Value = ch698[3];
                    ws.Cells[i + 4, 39].Value = ch698[4];
                    ws.Cells[i + 4, 42].Value = ch698[5];
                    ch698.Clear();
                    List<string> ch885 = ot.getData("885");
                    ws.Cells[i + 4, 28].Value = ch885[0];
                    ws.Cells[i + 4, 31].Value = ch885[1];
                    ws.Cells[i + 4, 34].Value = ch885[2];
                    ws.Cells[i + 4, 37].Value = ch885[3];
                    ws.Cells[i + 4, 40].Value = ch885[4];
                    ws.Cells[i + 4, 43].Value = ch885[5];
                    ch885.Clear();
                }
                for(int r=1;r<=ws.Dimension.End.Row;r++)
                {
                    for (int c = 1; c <= ws.Dimension.End.Column; c++)
                    {
                        ws.Cells[r, c].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(0, 0, 0));
                    }
                }
                    excel.Save();
            }
        }

        private void mergeexcel_Load(object sender, EventArgs e)
        {

        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
    }
}
