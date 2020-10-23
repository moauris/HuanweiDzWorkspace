using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HuanweiDzWPF.ViewModels;
using HuanweiDzWPF.Models;
using System.Collections.ObjectModel;

namespace HuanweiDzWPF.Services
{
    public class Consolidator
    {


        /// <summary>
        /// A method that consolidates a viewmodel
        /// Default Mode, 1 item to 1 item
        /// </summary>
        /// <param name="model">
        /// Target viewmodel to be consolidated
        /// </param>
        public static void Consolidate(MainWindowViewModels model)
        {
            foreach (LedgerItem bankitem in model.BankLedger)
            {
                foreach (LedgerItem companyitem in model.CompanyLedger)
                {
                    //Console.WriteLine("正在比较");
                    //Console.WriteLine(bankitem);
                    //Console.WriteLine(companyitem);
                    if (companyitem.Paired) //如果已经被同步了，跳过。
                    {
                        continue;
                    }
                    if (
                            (Math.Abs(bankitem.DebitRemain - companyitem.DebitRemain) < 0.00001d)
                        )
                    { //如果两者的差值小于0.00001则视为一样
                        ConsolidatedPair pair = new ConsolidatedPair(
                            new ObservableCollection<LedgerItem> { companyitem }, 
                            new ObservableCollection<LedgerItem> { bankitem });//
                        model.MatchedCollection.Add(pair); //添加
                        //model.BankLedger.Remove(bankitem); //删减 ！ 在循环中不可以修改被循环集合的内容
                        //model.BankLedger.Remove(companyitem); //删减
                        //TODO: 需要在 LedgerItem 中添加两个属性，第一个贷方余额，第二个是 paired 布尔函数。
                        companyitem.Paired = true;
                        bankitem.Paired = true;

                    }
                }

            }
            //清理所有Paired为true的对象
            var bankitemNotPaired = from i in model.BankLedger
                                    where i.Paired = true
                                    select i;
            for (int i = 0; i < bankitemNotPaired.Count(); i++)
            {
                model.BankLedger.Remove(bankitemNotPaired.ToArray()[i]);
            }
            var companyitemNotPaired = from i in model.CompanyLedger
                                    where i.Paired = true
                                    select i;
            for (int i = 0; i < companyitemNotPaired.Count(); i++)
            {
                model.CompanyLedger.Remove(companyitemNotPaired.ToArray()[i]);
            }


        }
    }
}
