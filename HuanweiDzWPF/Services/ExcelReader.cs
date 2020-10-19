using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace HuanweiDzWPF.Services
{
    public class ExcelReader
    {
        public string ExcelFilePath { get; set; }
        public string Side { get; set; }
        public ExcelReader(string path, string side)
        {
            ExcelFilePath = path;
            Side = side;
        }

        public void Run()
        {
            //开始之前记录所有的excel程序进程
#if DEBUG

            Process[] excelInstances = Process.GetProcessesByName("EXCEL");
            List<int> oldPID = new List<int>();
            foreach (Process p in excelInstances)
            {
                Debug.Print("{0}\t{1}\t{2}\t{3}", p.Id, p.ProcessName, p.StartTime, p.WorkingSet64);
                oldPID.Add(p.Id);
            }
#endif

            //打开表格，开始判定账本位置
            object misVal = System.Reflection.Missing.Value;
            EXCEL.Application app = new EXCEL.Application();
            EXCEL.Workbooks books = app.Workbooks;
            EXCEL.Workbook book = books.Open(ExcelFilePath,
        0, false, 5, "", "", false, EXCEL.XlPlatform.xlWindows, "",
        true, false, 0, true, false, false);

            //开始运行
            switch (Side)
            {
                case "Company":
                    //执行公司侧账目同步逻辑
                    break;
                default:
                    break;
            }



            //关闭表格，垃圾处理

            book.Close(false, misVal, misVal);
            app.Quit();

            GC.Collect();
            GC.WaitForPendingFinalizers();

#if DEBUG
            Process[] newExcelInstances = Process.GetProcessesByName("EXCEL");
            IEnumerable<Process> query =
                from p in newExcelInstances
                where !oldPID.Contains(p.Id)
                select p;

            Debug.Print("旧进程：");
            foreach (Process p in excelInstances)
            {
                Debug.Print("{0}\t{1}\t{2}\t{3}", p.Id, p.ProcessName, p.StartTime, p.WorkingSet64);
                

            }
            Debug.Print("新出现的Excel进程：");
            foreach (Process p in query)
            {
                Debug.Print("{0}\t{1}\t{2}\t{3}", p.Id, p.ProcessName, p.StartTime, p.WorkingSet64);
                p.Kill();

            }

#endif
        }
    }
}
