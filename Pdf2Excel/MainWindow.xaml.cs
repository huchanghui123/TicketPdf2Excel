/**
 * ********************************************************************
 * PDF解析使用Free Spire.PDF开源组件                                                                                
 * 免费版有 10 页的页数限制，在创建和加载 PDF 文档时要求文档不超过 10 页                                    
 * 将 PDF 文档转换为图片时，仅支持转换前 3 页       
 * EXCEL使用EPPLUS组件，个人用途
 * ********************************************************************
 */
using OfficeOpenXml;
using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;

namespace Pdf2Excel
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool isDrop = false;
        private int row = 0;
        private string fileName = String.Empty;
        private int[] arr = new int[2];
        private List<int> pro_list = new List<int>();
        private TicketItem ticket;
        private StringBuilder content = new StringBuilder();

        public static int index_code = 0;
        public static int index_number = 0;
        public static int index_date = 0;
        public static int index_sum = 0;
        public static int index_name = 0;

        public MainWindow()
        {
            InitializeComponent();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        private void OpenFileClick(object sender, RoutedEventArgs e)
        {
            row = 0;
            ticket = null;
            Array.Clear(arr,0,arr.Length);
            pro_list.Clear();
            content.Remove(0,content.Length);

            if (!isDrop)
            {
                var openFileDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "Pdf Files (*.pdf)|*.pdf"
                };
                if (openFileDialog.ShowDialog() == true)
                {
                    fileName = openFileDialog.FileName;
                    this.text_file_name.Text = fileName;
                }
                else
                {
                    return;
                }
            }
            
            try
            {
                content = PdfFormat.GetFdfText(fileName);
                //Console.WriteLine(content.ToString());
                if (content.Length == 0)
                {
                    MessageBox.Show("...发票!发票!发票!", "提示");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
            //StreamWriter sw = File.AppendText(@"C:\\Users\\16838\\Desktop\\test2323.txt");
            //sw.Write(content.ToString());
            //sw.Flush();
            //sw.Close();

            String[] contentArry = content.ToString().Split(new string[] { "\r\n" },
                StringSplitOptions.RemoveEmptyEntries);
            if (contentArry.Length == 0)
            {
                Console.WriteLine("split failed!");
                return;
            }
            foreach (String str in contentArry)
            {
                //Console.WriteLine("{0}, {1}", row, str);
                if (str.IndexOf("发票代码") > -1)
                {
                    index_code = row;
                }
                else if (str.IndexOf("发票号码") > -1)
                {
                    index_number = row;
                }
                else if (str.IndexOf("开票日期") > -1)
                {
                    index_date = row;
                }
                //这里需要知道货品的行数，用规格型号行和合计行推算出来
                else if (str.IndexOf("规格型号") > -1)
                {
                    var start = row;
                    arr[0] = start;
                    Console.WriteLine("start----------" + start);
                }
                //合计行容易冲突，这里用价税合计代替，行数要-1
                else if (str.IndexOf("价税合计") > -1)
                {
                    index_sum = row;
                    var end = row;
                    arr[1] = end - 1;
                    Console.WriteLine("end----------" + end);
                }
                else if (str.LastIndexOf("称:") > -1)
                {
                    index_name = row;
                    Console.WriteLine("index_name----------" + index_name);
                }
                row++;
            }
            for (int i=arr[0]+1;i<arr[1];i++)
            {
                pro_list.Add(i);
            }
            //Console.WriteLine("index_code:{0} index_number:{1} index_date:{2} index_name:{3}",
            //    index_code, index_number, index_date, index_name);

            ticket = PdfFormat.GetTicketItem(contentArry, arr, pro_list);
            Console.WriteLine(ticket.ToString());
            PdfFormat.RenamePdf(ticket, fileName);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //ExcelFormat.CreateSinglelExcel(ticket);
            string filePath = AppDomain.CurrentDomain.BaseDirectory + ExcelFormat.ExcelNamed;
            if (File.Exists(filePath))
            {
                ExcelFormat.AddData2Excel(ticket);
            }
            else
            {
                ExcelFormat.CreateNewExcel(ticket);
            }
            isDrop = false;
            
        }

        private void OnDrop(object sender, DragEventArgs e)
        {
            fileName = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            //只允许DPF文件
            if (System.IO.Path.GetExtension(fileName).ToUpperInvariant()!=".PDF")
            {
                MessageBox.Show("不是PDF!","提示");
                return;
            }
            isDrop = true;
            this.text_file_name.Text = fileName;
            OpenFileClick(btn, new RoutedEventArgs());
        }

        private void OnDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.Text))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
        }

    }
}
