using System;
using HuanweiDzModels;

namespace HuanweiDzWorkspace
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            object[] parameters =
            {
                null, //new DateTime(1993, 3, 27),
                "测试用的 Ledger 对象",
                0D,
                3997D,
                "贷方",
                322209.27D
            };
            LedgerItem ledger = new LedgerItem(LedgerSides.FromBank, parameters);

            Console.WriteLine(ledger.Credit);
            Console.WriteLine(ledger.Debit);
            Console.WriteLine(ledger.IncuredOn);
            Console.WriteLine(ledger.Info);

        }
    }
}
