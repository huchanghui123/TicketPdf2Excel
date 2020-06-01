using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Pdf2Excel
{
    class PdfFormat
    {
        public static StringBuilder GetFdfText(string filename)
        {
            StringBuilder content = new StringBuilder();
            PdfDocument document = new PdfDocument();
            document.LoadFromFile(filename);

            //提取PDF所有页面的文本
            foreach (PdfPageBase page in document.Pages)
            {
                content.Append(page.ExtractText());
            }
            if (content.ToString().IndexOf("发票代码") < 0)
            {
                content.Remove(0, content.Length);
            }
            return content;
        }

        public static TicketItem GetTicketItem(String[] strArr, int[] arr, List<int> pro_list)
        {
            TicketItem item = new TicketItem();
            //第一行发票代码
            string code = "发票代码";
            string tickets_code = strArr[MainWindow.index_code].Substring(strArr[MainWindow.index_code].IndexOf(code) + code.Length + 1);
            //Console.WriteLine(tickets_code);
            //第二行发票号码
            string number = "发票号码";
            string tickets_number = strArr[MainWindow.index_number].Substring(strArr[MainWindow.index_number].IndexOf(number) + number.Length + 1);
            //Console.WriteLine(tickets_number);
            //第三行开票日期
            string date = "开票日期";
            string result = strArr[MainWindow.index_date].Substring(strArr[MainWindow.index_date].IndexOf(date) + date.Length + 1);
            string date_str = System.Text.RegularExpressions.Regex.Replace(result, @"[^0-9]+", "");
            string tickets_date = date_str.Insert(4,"年").Insert(7,"月").Insert(10, "日");
            //Console.WriteLine(tickets_date);

            item.Code = tickets_code;
            item.Number = tickets_number;
            item.Date = date_str;

            var product = new StringBuilder();
            foreach (var index in pro_list)
            {
                String content = strArr[index];
                Console.WriteLine("货品："+content);
                String[] project_info = { "货品名称 型号"};
                //有时候解析出的行全是空格
                String check = Regex.Replace(content, @"\s", "");
                if(check.Length > 0)
                {
                    project_info = content.ToString().Split(new string[] { " " },
                        StringSplitOptions.RemoveEmptyEntries);
                    //Console.WriteLine("project_info.Length：" + project_info.Length + " content.Length:"+ content.Length + " check:"+ check);
                    product.Append(project_info[0]).Append(" ").Append(project_info[1]).Append("\r\n");
                }
            }
            if (product.Length>2)
            {
                product.Remove(product.Length - 2, 2);
            }
            //Console.WriteLine(product);
            item.Project = product;

            string sum = "小写";
            string tickets_sum = strArr[MainWindow.index_sum].Substring(strArr[MainWindow.index_sum].IndexOf(sum));
            if (strArr[MainWindow.index_sum].IndexOf("￥") > -1)
            {
                tickets_sum = strArr[MainWindow.index_sum].Substring(strArr[MainWindow.index_sum].IndexOf("￥"));
            }
            else if (strArr[MainWindow.index_sum].IndexOf("¥") > -1)
            {
                tickets_sum = strArr[MainWindow.index_sum].Substring(strArr[MainWindow.index_sum].IndexOf("¥"));
            }

            //string rex = "名　　　　称";
            String temp_company = strArr[MainWindow.index_name].Replace("：",":");
            string tickets_company = temp_company.Split(':')[1].Trim().Split(' ')[0];
            //Console.WriteLine(tickets_company);

            item.Sum = tickets_sum;
            item.Company = tickets_company;

            return item;
        }

        public static void RenamePdf(TicketItem ticket, String filePath)
        {
            string filename = String.Format("{0}-{1}-{2}.pdf", ticket.Date, ticket.Sum.Substring(1), ticket.Company);
            //保存Excel文件
            string newfilepath = AppDomain.CurrentDomain.BaseDirectory + filename;
            FileInfo fi = new FileInfo(filePath);
            //fi.MoveTo(newfilepath);
            if (!System.IO.File.Exists(newfilepath))
            {
                fi.CopyTo(newfilepath);
            }
            
        }
    }
}
