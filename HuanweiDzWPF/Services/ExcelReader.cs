using ExcelDataReader;
using HuanweiDzWPF.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace HuanweiDzWPF.Services
{
    public class ExcelReader
    {
        public string filePath { get; set; }
        public string Side { get; set; }
        public ExcelReader(string path, string side)
        {
            this.filePath = path;
            Side = side;
        }

        public ObservableCollection<LedgerItem> Read()
        {
            var res = new ObservableCollection<LedgerItem>();
            //Trace.Listeners.Clear();
            string LogFileName = string.Format(".\\{0}_Trace.log"
                , DateTime.Now.ToString("yyyy_MMdd_HHmmss"));
            //TextWriterTraceListener traceListener = new TextWriterTraceListener(LogFileName);
            //Trace.Listeners.Add(traceListener);
            //TraceWrapper("正在开始读取文件" + filePath);

            FileInfo file = new FileInfo(filePath);
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))//当文件被占用时会报错。
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream)) //需要加入文件后缀和类型的判定
                    {
                        //仅对bank：在RowContent中的第0项寻找xxxx年字段，提取年份
                        Regex rxGetYear = new Regex(@"\d{4}(?=年)");
                        int BankYear = 0;
                        double LastRemain = 0d;
                        do
                        {
                            while (reader.Read())
                            {
                                string[] RowContent = new string[9];
                                int NonEmptyLength = 0;
                                for (int r = 0; r < 9; r++)
                                {
                                    object raw = reader.GetValue(r);
                                    if (raw == null)
                                    {
                                        RowContent[r] = string.Empty;
                                    }
                                    else
                                    {
                                        RowContent[r] = reader.GetValue(r).ToString();
                                        NonEmptyLength++;
                                    }
                                }
                                if (BankYear == 0) //在没有找到bankyear的情况下，试图找出
                                {
                                    Match match = rxGetYear.Match(RowContent[0]);
                                    if (match.Success) BankYear = int.Parse(match.Value);
                                    
                                }

                                //TraceWrapper("非空单行元素判定: " + NonEmptyLength);
                                //string RowContents = string.Join(",", RowContent);

                                //TraceWrapper(RowContents);
                                //以下代码生成 LedgerItem 对象
                                //两侧符合要求的对象：第0-4不为空，5或6至少有一位不为空，不为空时可以被转换为double。7为平、借、或者贷，余额为double
                                LedgerItem ledgerItem = null;
                                //建立组成会计条目的参数
                                string preDateString = "", LedgerInfo = "";
                                DateTime DateIncurred;
                                int CredDebNotSatisfy = 0;
                                double CreditParm = 0d, DebitParm = 0d, RemainParm = 0d;
                                Regex rx = new Regex(@"\d+(\.\d{1,2}){0,1}");
                                //新加一项检查：检查账目是否连续：
                                //即：本次remain不为0时，检查是否成立：
                                //本次remain + 本次debit - 本次credit = 上次remain
                                switch (Side)
                                {
                                    case "Company":
                                        
                                        //从0-2试图组成datetime
                                        preDateString = RowContent[0] + "-" + RowContent[1] + "-" + RowContent[2];
                                        if (!DateTime.TryParse(preDateString, out DateIncurred))
                                        {
                                            break;
                                        }
                                        //从3，4组成信息
                                        LedgerInfo = RowContent[3] + "," + RowContent[4];
                                        //从5组成借方，判定其是否满足regex = \d+(\.\d{1,2}){0,1}
                                        if (rx.IsMatch(RowContent[5]))
                                        {
                                            double.TryParse(RowContent[5], out CreditParm);
                                        }
                                        else CredDebNotSatisfy++;
                                        if (rx.IsMatch(RowContent[6]))
                                        {
                                            double.TryParse(RowContent[6], out DebitParm);
                                        }
                                        else CredDebNotSatisfy++;
                                        if (rx.IsMatch(RowContent[8]))
                                        {
                                            double.TryParse(RowContent[8], out RemainParm);
                                        }
                                        else break;

                                        //Credit and Debit 必须至少满足一项，如果有CredDebNotSatisfy > 1, break
                                        if (CredDebNotSatisfy > 1) break;
                                        string Direction = RowContent[7];
                                        ledgerItem = new LedgerItem
                                        (DateIncurred, LedgerInfo, CreditParm,
                                        DebitParm, Direction, RemainParm);
                                        //检查账目是否连续：
                                        if (LastRemain != 0d)
                                        {
                                            double calculatedThisRemain = LastRemain + ledgerItem.Credit - ledgerItem.Debit;
                                            Debug.Print("正在比较上次余额{0}+本次贷方{1}-本次借方{2}={3}与本次余额{4}, 进行相等比较，差值：{5}",
                                                LastRemain, ledgerItem.Credit,
                                                ledgerItem.Debit, calculatedThisRemain,
                                                ledgerItem.RemainingFund,
                                                calculatedThisRemain - ledgerItem.RemainingFund);
                                            //在比较double的时候有精确度问题，需要有容忍度
                                            double diff = calculatedThisRemain - ledgerItem.RemainingFund;
                                            if (Math.Abs(diff) > 0.1d)
                                            {
                                                ledgerItem = null;
                                                break;
                                            }
                                        }
                                        LastRemain = ledgerItem.RemainingFund;
                                        break;
                                    case "Bank":
                                        //从BankYear, 0, 1试图组成datetime
                                        preDateString = BankYear + "-" + RowContent[0] + "-" + RowContent[1];
                                        if (!DateTime.TryParse(preDateString, out DateIncurred))
                                        {
                                            break;
                                        }
                                        //从2, 3，4组成信息
                                        LedgerInfo = RowContent[2] + "," + RowContent[3] + "," + RowContent[4];
                                        //从5, 6, 8组成借方，判定其是否满足regex = \d+(\.\d{1,2}){0,1}
                                        if (rx.IsMatch(RowContent[5]))
                                        {
                                            double.TryParse(RowContent[5], out CreditParm);
                                        }
                                        else CredDebNotSatisfy++;
                                        if (rx.IsMatch(RowContent[6]))
                                        {
                                            double.TryParse(RowContent[6], out DebitParm);
                                        }
                                        else CredDebNotSatisfy++;
                                        if (rx.IsMatch(RowContent[8]))
                                        {
                                            double.TryParse(RowContent[8], out RemainParm);
                                        }
                                        else break;

                                        //Credit and Debit 必须至少满足一项，如果有CredDebNotSatisfy > 1, break
                                        if (CredDebNotSatisfy > 1) break;
                                        Direction = RowContent[7];
                                        ledgerItem = new LedgerItem
                                        (DateIncurred, LedgerInfo, CreditParm,
                                        DebitParm, Direction, RemainParm);
                                        //检查账目是否连续：
                                        if (LastRemain != 0d)
                                        {
                                            double calculatedThisRemain = LastRemain + ledgerItem.Credit - ledgerItem.Debit;
                                            Debug.Print("正在比较上次余额{0}+本次贷方{1}-本次借方{2}={3}与本次余额{4}, 进行相等比较，差值：{5}",
                                                LastRemain, ledgerItem.Credit,
                                                ledgerItem.Debit, calculatedThisRemain, 
                                                ledgerItem.RemainingFund, 
                                                calculatedThisRemain - ledgerItem.RemainingFund);
                                            //在比较double的时候有精确度问题，需要有容忍度
                                            double diff = calculatedThisRemain - ledgerItem.RemainingFund;
                                            if (Math.Abs(diff) > 0.1d)
                                            {
                                                ledgerItem = null;
                                                break;
                                            }
                                        }
                                        LastRemain = ledgerItem.RemainingFund;
                                        break;
                                    default:
                                        break;
                                }
                                
                                if (ledgerItem != null) //需要多加一项判断：不可以导入重复项
                                        //已经添加逻辑，如果为不连续则为null;
                                {
                                    
                                    res.Add(ledgerItem);
                                    //TraceWrapper(ledgerItem.ToString());
                                }

                            }
                        } while (reader.NextResult());
                    }
                }
            }
            catch (IOException exp)
            {

                //TraceWrapper("遇到了文件读写错误：");
                //TraceWrapper(exp.Message);

                MessageBox.Show(
                    "文件读写遇到了错误。\r\n请检查目标工作簿是否已经打开或者被其他程序占用。\r\n请释放工作簿后再次尝试。",
                    "发生了错误：目标文件被程序占用。\r\n" + exp.Message,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                    );

            }
            finally
            {
                MessageBox.Show("读取文件完成。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //Trace.Flush();
            }
            ObservableCollection<LedgerItem> finalRes = new ObservableCollection<LedgerItem>();
            foreach (LedgerItem item in res)
            {
                finalRes.Add(item);
            }
            return finalRes;
        }
        /*
        [Conditional ("DEBUG")]
        private void TraceWrapper(string message)
        {
            string DebugMessage = string.Format("[{0}] @ <{1}>: {2}, EOL"
                , DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
                , "ExcelReaderXLSReader"
                , message);
            Trace.WriteLine(DebugMessage);
        }
        */
    }
}
