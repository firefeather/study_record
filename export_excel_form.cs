using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.POIFS.FileSystem;
using NPOI.HSSF.Util;

namespace wpa_hjp_001
{
    public partial class MainForm : Form
    {
        static IWorkbook m_workbook;
        static String fileName;

        public MainForm()
        {
            InitializeComponent();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.W))
            {
                // Do what you want to do here
                MessageBox.Show(@"这是一个测试！");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);

        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            fileName = "JS75不良区域统计表_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xls";
            if(!File.Exists(fileName))
            {
                MemoryStream ms = new MemoryStream();    // 创建内存流用于写入文件       
                IWorkbook workbook = new HSSFWorkbook();   // 创建Excel工作部   
                ISheet sheet = workbook.CreateSheet("数据表1");// 创建工作表

                // 写标题文本     
                ICell cellTitle = sheet.CreateRow(0).CreateCell(0);
                cellTitle.SetCellValue("JS75不良区域统计表");

                // 合并标题行
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 3);
                sheet.AddMergedRegion(region);

                // 设置标题行样式
                ICellStyle style = workbook.CreateCellStyle();
                style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                IFont font = workbook.CreateFont();
                font.FontHeight = 20 * 20;
                style.SetFont(font);
                cellTitle.CellStyle = style;

                HSSFSheet isheet = workbook.GetSheet("数据表1") as HSSFSheet;
                for (int i = region.FirstRow; i <= region.LastRow; i++)
                {
                    IRow irow = HSSFCellUtil.GetRow(i, isheet);
                    for (int j = region.FirstColumn; j <= region.LastColumn; j++)
                    {
                        ICell singleCell = HSSFCellUtil.GetCell(irow, (short)j);
                        singleCell.CellStyle = style;
                    }
                }

                // 设置第一列列宽
                sheet.SetColumnWidth(0, 20 * 256);
                sheet.SetColumnWidth(1, 40 * 256);
                sheet.SetColumnWidth(2, 100 * 256);
                sheet.SetColumnWidth(3, 15 * 256);

                // 表头数据
                IRow row = sheet.CreateRow(1);

                // 单元格样式
                ICellStyle cellStyle = workbook.CreateCellStyle();
                cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

                // 对齐
                cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

                // 设置字体
                font = workbook.CreateFont();
                font.FontHeightInPoints = 18;
                font.FontName = "微软雅黑";
                cellStyle.SetFont(font);

                ICell cell = row.CreateCell(0);
                cell.SetCellValue("日期");
                cell.CellStyle = cellStyle;

                cell = row.CreateCell(1);
                cell.SetCellValue("Serials Number");
                cell.CellStyle = cellStyle;

                cell = row.CreateCell(2);
                cell.SetCellValue("Defect Code");
                cell.CellStyle = cellStyle;

                cell = row.CreateCell(3);
                cell.SetCellValue("区域");
                cell.CellStyle = cellStyle;

                // 将Excel写入流
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                workbook = null;

                FileStream dumpFile = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                ms.WriteTo(dumpFile);// 将流写入文件
            }

            // 初始化UI    
            tb_date.Text = DateTime.Now.Date.ToString("MM/dd");

            string[] array = { "Product点状物异常[Product Foreign Dot-Material]",
                "Product纤维状异物[Product Foreign Fiber-Material]",
                "TP污染[TP Dirt]",
                "TP刺伤不良[TP Surface Prick]",
                "CG可视区刮伤[CGScratch]",
            "TP凹凸点不良[TP Surface Burr]",
            "TP崩边[TP Edge Broken]",
            "CG异色[CG Discoloration]",
            "TP崩脚[TP Comer Broken]",
            "CG油墨不良[PoorInk]"};
            cb_defectCode.DataSource = array;
            cb_defectCode.Text = "";
        }

        private void button_Click(object sender, EventArgs e)
        {
            if (!File.Exists(fileName))
            {
                MessageBox.Show("请重启软件！");
                return;
            }

            FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);// 读取流
            fs.Seek(0, SeekOrigin.Begin);
            POIFSFileSystem ps = new POIFSFileSystem(fs);// 需using NPOI.POIFS.FileSystem;
            IWorkbook workbook = new HSSFWorkbook(ps);
            ISheet sheet = workbook.GetSheetAt(0); // 获取工作表
            IRow row = sheet.GetRow(1); // 得到表头
            row = sheet.CreateRow((sheet.LastRowNum + 1));// 在工作表中添加一行

            ICell cell = row.CreateCell(0);

            // 创建单元格样式
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

            // 对齐
            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

            // 设置字体
            IFont font = workbook.CreateFont();
            font.FontHeightInPoints = 18;
            font.FontName = "微软雅黑";
            cellStyle.SetFont(font);


            string strDate = tb_date.Text;
            string strSN = tb_sn.Text;
            string strDefectCode = cb_defectCode.Text;
            string strArea = ((Button)sender).Text.Substring(4);
            //设置值
            cell = row.CreateCell(0);
            cell.SetCellValue(strDate);
            cell.CellStyle = cellStyle;

            cell = row.CreateCell(1);
            cell.SetCellValue(strSN);
            cell.CellStyle = cellStyle;

            cell = row.CreateCell(2);
            cell.SetCellValue(strDefectCode);
            cell.CellStyle = cellStyle;

            cell = row.CreateCell(3);
            cell.SetCellValue(strArea);
            cell.CellStyle = cellStyle;

            m_workbook = workbook;
            WriteToFile();
        }

        static void WriteToFile()
        {
            FileStream fout = new FileStream(fileName, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);// 写入流
            fout.Flush();
            m_workbook.Write(fout);//写入文件
            m_workbook = null;
            fout.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show(DateTime.Now.Date.ToString("MM/dd"));
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        { 
            if(0 == textBox1.Text.Length)
            {
                return;
            }
            fileName = textBox1.Text + ".xls";
            if (!File.Exists(fileName))
            {
                MemoryStream ms = new MemoryStream();    // 创建内存流用于写入文件       
                IWorkbook workbook = new HSSFWorkbook();   // 创建Excel工作部   
                ISheet sheet = workbook.CreateSheet("数据表1");// 创建工作表

                // 写标题文本     
                ICell cellTitle = sheet.CreateRow(0).CreateCell(0);
                cellTitle.SetCellValue("JS75不良区域统计表");

                // 合并标题行
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 3);
                sheet.AddMergedRegion(region);

                // 设置标题行样式
                ICellStyle style = workbook.CreateCellStyle();
                style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                IFont font = workbook.CreateFont();
                font.FontHeight = 20 * 20;
                style.SetFont(font);
                cellTitle.CellStyle = style;

                HSSFSheet isheet = workbook.GetSheet("数据表1") as HSSFSheet;
                for (int i = region.FirstRow; i <= region.LastRow; i++)
                {
                    IRow irow = HSSFCellUtil.GetRow(i, isheet);
                    for (int j = region.FirstColumn; j <= region.LastColumn; j++)
                    {
                        ICell singleCell = HSSFCellUtil.GetCell(irow, (short)j);
                        singleCell.CellStyle = style;
                    }
                }

                // 设置第一列列宽
                sheet.SetColumnWidth(0, 20 * 256);
                sheet.SetColumnWidth(1, 40 * 256);
                sheet.SetColumnWidth(2, 100 * 256);
                sheet.SetColumnWidth(3, 15 * 256);

                // 表头数据
                IRow row = sheet.CreateRow(1);

                // 单元格样式
                ICellStyle cellStyle = workbook.CreateCellStyle();
                cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

                // 对齐
                cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

                // 设置字体
                font = workbook.CreateFont();
                font.FontHeightInPoints = 18;
                font.FontName = "微软雅黑";
                cellStyle.SetFont(font);

                ICell cell = row.CreateCell(0);
                cell.SetCellValue("日期");
                cell.CellStyle = cellStyle;

                cell = row.CreateCell(1);
                cell.SetCellValue("Serials Number");
                cell.CellStyle = cellStyle;

                cell = row.CreateCell(2);
                cell.SetCellValue("Defect Code");
                cell.CellStyle = cellStyle;

                cell = row.CreateCell(3);
                cell.SetCellValue("区域");
                cell.CellStyle = cellStyle;

                // 将Excel写入流
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                workbook = null;

                FileStream dumpFile = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                ms.WriteTo(dumpFile);// 将流写入文件
            }
        }
    }
}
