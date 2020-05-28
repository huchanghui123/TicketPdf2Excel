using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pdf2Excel
{
    class TicketItem
    {
        //发票代码
        public string Code { get; set; }
        //发票号码
        public string Number { get; set; }
        //开票日期
        public string Date { get; set; }
        //产品名称
        public StringBuilder Project { get; set; }
        //总金额
        public string Sum { get; set; }
        //公司名称
        public string Company { get; set; }

        public TicketItem() { }

        public TicketItem(string code, string number, string date, StringBuilder project, string sum, string company)
        {
            Code = code;
            Number = number;
            Date = date;
            Project = project;
            Sum = sum;
            Company = company;
        }

        public override string ToString()
        {
            String result = String.Format("发票代码:{0}\r\n发票号码:{1}\r\n开票日期:{2}\r\n产品名称:\r\n{3}\r\n总金额:{4}\r\n公司名称:{5}\r\n",
                Code, Number, Date, Project.ToString(), Sum, Company);
            return result;
        }
    }
}
