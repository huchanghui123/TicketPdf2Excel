using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Threading;

namespace Pdf2Excel
{
    class ExcelFormat
    {
        public static string ExcelNamed = "发票汇总.xlsx";
        static string AppPath = AppDomain.CurrentDomain.BaseDirectory;

        public static void CreateSinglelExcel(TicketItem ticket)
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //设置Excel文档的某些属性
                excelPackage.Workbook.Properties.Author = "QOTOM";
                excelPackage.Workbook.Properties.Title = "发票数据";
                excelPackage.Workbook.Properties.Created = DateTime.Now;

                //创建工作表
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("发票数据");

                workSheet.Cells[1, 1].Value = "发票代码";
                workSheet.Cells[1, 2].Value = "发票号码";
                workSheet.Cells[1, 3].Value = "开票日期";
                workSheet.Cells[1, 4].Value = "货品名称和型号";
                workSheet.Cells[1, 5].Value = "总金额";
                workSheet.Cells[1, 6].Value = "公司名称";

                workSheet.Cells[2, 1].Value = ticket.Code;
                workSheet.Cells[2, 2].Value = ticket.Number;
                workSheet.Cells[2, 3].Value = ticket.Date;
                workSheet.Cells[2, 4].Value = ticket.Project.ToString();
                workSheet.Cells[2, 5].Value = ticket.Sum;
                workSheet.Cells[2, 6].Value = ticket.Company;

                //自适应宽度
                workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
                //workSheet.Row(1).CustomHeight = true;//自动换行
                workSheet.Column(4).Style.WrapText = true;
                string filename = String.Format("{0}-{1}-{2}-{3}.xlsx", ticket.Date, 
                    ticket.Sum.Substring(1), ticket.Number, ticket.Company);
                //保存Excel文件
                string filePath = AppPath + filename;
                Console.WriteLine("save to {0}", filePath);
                FileInfo fi = new FileInfo(filePath);
                excelPackage.SaveAs(fi);
            }
        }

        public static void CreateNewExcel(TicketItem ticket)
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //设置Excel文档的某些属性
                excelPackage.Workbook.Properties.Author = "QOTOM";
                excelPackage.Workbook.Properties.Title = "发票数据";
                excelPackage.Workbook.Properties.Created = DateTime.Now;

                //创建工作表
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("发票数据");
                
                workSheet.Cells[1, 1].Value = "发票代码";
                workSheet.Cells[1, 2].Value = "发票号码";
                workSheet.Cells[1, 3].Value = "开票日期";
                workSheet.Cells[1, 4].Value = "货品名称和型号";
                workSheet.Cells[1, 5].Value = "总金额";
                workSheet.Cells[1, 6].Value = "公司名称";

                workSheet.Cells[2, 1].Value = ticket.Code;
                workSheet.Cells[2, 2].Value = ticket.Number;
                workSheet.Cells[2, 3].Value = ticket.Date.Insert(4, "年").Insert(7, "月").Insert(10, "日");
                workSheet.Cells[2, 4].Value = ticket.Project.ToString();
                workSheet.Cells[2, 5].Value = ticket.Sum;
                workSheet.Cells[2, 6].Value = ticket.Company;

                //自适应宽度
                workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
                //workSheet.Row(1).CustomHeight = true;//自动换行
                workSheet.Column(4).Style.WrapText = true;

                //保存Excel文件
                string filePath = AppPath + ExcelNamed;
                Console.WriteLine("save to {0}", filePath);
                FileInfo fi = new FileInfo(filePath);
                excelPackage.SaveAs(fi);
            }
        }

        public static void AddData2Excel(TicketItem ticket)
        {
            int row = 1;
            int cell = 1;
            string filePath = AppPath + ExcelNamed;
            //创建一个列表以保存所有值
            List<string> excelData = new List<string>();
            byte[] bin = new byte[1024];
            //读取Excel文件为字节数组
            try
            {
                bin = File.ReadAllBytes(filePath);
            }
            catch (Exception)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                {
                    MessageBox.Show("添加数据失败，请先关闭"+ ExcelNamed, "提示", 
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }));
                return;
            }
            
            //byte[] bin = File.ReadAllBytes(Server.MapPath("test.xlsx"));
            //在内存流中创建一个新的Excel包
            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                //循环所有工作表
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    //循环所有行
                    for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                    {
                        //循环每列
                        for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                        {
                            //将单元格数据添加到列表
                            if (worksheet.Cells[i, j].Value != null)
                            {
                                excelData.Add(worksheet.Cells[i, j].Value.ToString());
                            }
                        }
                    }
                }
            }

            if (excelData.IndexOf(ticket.Number) > 0)
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                {
                    MessageBox.Show("发票重复", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }));
                return;
            }
            //添加数据到最后一行
            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
                
                row = worksheet.Dimension.End.Row;
                cell = worksheet.Dimension.End.Column;
                Console.WriteLine("max row:{0} max cell:{1}", row, cell);

                //将单元格数据添加到列表
                worksheet.Cells[row + 1, 1].Value = ticket.Code;
                worksheet.Cells[row + 1, 2].Value = ticket.Number;
                worksheet.Cells[row + 1, 3].Value = ticket.Date.Insert(4, "年").Insert(7, "月").Insert(10, "日");
                worksheet.Cells[row + 1, 4].Value = ticket.Project.ToString();
                worksheet.Cells[row + 1, 5].Value = ticket.Sum;
                worksheet.Cells[row + 1, 6].Value = ticket.Company;

                //excelPackage.Save();
                FileInfo fi = new FileInfo(filePath);
                excelPackage.SaveAs(fi);
            }
            
        }
        
    }
}
