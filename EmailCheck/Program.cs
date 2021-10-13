using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace EmailCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length==0)
            {
                Console.WriteLine("缺少excel檔案路徑");
                return;
            }

            Console.WriteLine("Email檢查開始");

            // 取得Excel Email
            var excelEmails = GetExcelEmails(args[0]);
            // 驗證Email 略過非正確的
            var failEmails = GetFailEmail(excelEmails);

            if (failEmails.Count > 0)
            {
                Console.WriteLine("錯誤Email如下:");
                foreach (var item in failEmails)
                {
                    Console.WriteLine(item);
                }
            }
            else
            {
                Console.WriteLine("無錯誤Email");
            }

            Console.WriteLine("Email檢查結束");
        }


        public static List<string> GetExcelEmails(string filePath)
        {
            // 取 email
            List<string> excelEmails = new List<string>();
            var workbook = new XLWorkbook(filePath);
            var sheet = workbook.Worksheets.FirstOrDefault();
            var rowsCount = sheet.RangeUsed().RowsUsed().Count();
            for (int i = 2; i <= rowsCount; i++)
            {
                var cellValue = sheet.Row(i).Cell(1).Value.ToString();
                if (string.IsNullOrWhiteSpace(cellValue))
                {
                    continue;
                }

                excelEmails.Add(cellValue.ToLower());
            }

            return excelEmails;
        }

        public static List<string> GetFailEmail(List<string> excelEmails)
        {
            List<string> failEmails = new List<string>();
            foreach (var email in excelEmails)
            {
                bool isEmail = Regex.IsMatch(email, @"[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?[^,]+$");
                if (!isEmail)
                {
                    // 不正確的加入list
                    failEmails.Add(email);
                }
            }
            return failEmails;
        }
    }
}
