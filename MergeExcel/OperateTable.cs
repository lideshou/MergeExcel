using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace MergeExcel
{
    public class OperateTable
    {

        /// <summary>
        /// 文件路径
        /// </summary>
        public string strPath = string.Empty;

        /// <summary>
        /// 获取SN(表格名称)的方法
        /// </summary>
        /// <returns></returns>
        public string getSN()
        {
            return Path.GetFileNameWithoutExtension(strPath);
        }

        /// <summary>
        /// 获取AFC Init DAC的方法
        /// </summary>
        /// <returns></returns>
        public string getAFC_Init_DAC()
        {
            string aid = string.Empty;
            FileStream fs = new FileStream(strPath, FileMode.Open);
            using (ExcelPackage pck = new ExcelPackage(fs))
            {
                for (int i = 1; i <= pck.Workbook.Worksheets.Count; ++i)
                {
                    ExcelWorksheet sheet = pck.Workbook.Worksheets[i];
                    for (int r = sheet.Dimension.Start.Row, mr = sheet.Dimension.End.Row; r <= mr; r++)
                    {
                        if (sheet.Cells[r, 1].Value != null && sheet.Cells[r, 1].Value.ToString().Contains("Init DAC"))
                        {
                            aid = sheet.Cells[r + 1, 1].Value.ToString();
                        }
                    }
                }
            }
            fs.Close();
            return aid;
        }

        /// <summary>
        /// 获取PLC5 APC DAC的方法
        /// </summary>
        /// <returns></returns>
        public List<string> getPCL5_APC_DAC()
        {
            List<string> pad = new List<string>();
            FileStream fs = new FileStream(strPath, FileMode.Open);
            using (ExcelPackage pck = new ExcelPackage(fs))
            {
                for (int i = 1; i <= pck.Workbook.Worksheets.Count; ++i)
                {
                    ExcelWorksheet sheet = pck.Workbook.Worksheets[i];
                    int endrow = sheet.Dimension.End.Row;
                    if (endrow <= 4)
                    {
                        pad.Add("MISS");
                        pad.Add("MISS");
                    }
                    else
                    {
                        for (int r = sheet.Dimension.Start.Row, mr = sheet.Dimension.End.Row; r <= mr; r++)
                        {
                            if (sheet.Cells[r, 4].Value != null && sheet.Cells[r, 4].Value.ToString().Contains("APC DAC"))
                            {
                                pad.Add(sheet.Cells[r + 1, 4].Value.ToString());
                            }
                        }
                    }
                }
            }
            fs.Close();
            return pad;
        }

        /// <summary>
        /// 获取每张表格有用数据的方法
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public List<string> getData(string s)
        {
            List<string> datalist = new List<string>(6);
            FileStream fs = new FileStream(strPath, FileMode.Open);
            using (ExcelPackage pck = new ExcelPackage(fs))
            {
                for (int i = 1; i <= pck.Workbook.Worksheets.Count; ++i)
                {
                    ExcelWorksheet sheet = pck.Workbook.Worksheets[i];
                    int endrow = sheet.Dimension.End.Row;
                    if (endrow <= 4)
                    {
                        datalist.Add("MISS");
                        datalist.Add("MISS");
                        datalist.Add("MISS");
                        datalist.Add("MISS");
                        datalist.Add("MISS");
                        datalist.Add("MISS");
                    }
                    else
                    {
                        for (int r = sheet.Dimension.Start.Row, mr = sheet.Dimension.End.Row; r <= mr; r++)
                        {
                            if (sheet.Cells[r, 2].Value != null && sheet.Cells[r, 2].Value.ToString() == s)
                            {
                                datalist.Add(sheet.Cells[r + 1, 2].Value.ToString());
                                datalist.Add(sheet.Cells[r + 4, 2].Value.ToString());
                                datalist.Add(sheet.Cells[r + 3, 2].Value.ToString());
                                datalist.Add(sheet.Cells[r + 5, 2].Value.ToString());
                                datalist.Add(sheet.Cells[r + 8, 1].Value.ToString());
                                datalist.Add(sheet.Cells[r + 8, 2].Value.ToString());
                            }
                        }
                    }
                }
            }
            fs.Close();
            return datalist;
        }
    }
}
