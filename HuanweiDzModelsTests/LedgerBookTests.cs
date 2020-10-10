using Microsoft.VisualStudio.TestTools.UnitTesting;
using HuanweiDzModels;
using System;
using System.Collections.Generic;
using System.Text;

namespace HuanweiDzModels.Tests
{
    [TestClass()]
    public class LedgerBookTests
    {
        public LedgerItem GenerateRandom()
        {
            Random rand = new Random();

            DateTime dateIncured = new DateTime(2019, rand.Next(3,7),rand.Next(1, 30));
            string InfoStringFormat = "付给{0}{1}的{2}";
            string[] Level0Seed = {
                "天上天下有限责任公司","疼逊公司",
                "宝贝猫","神探狄仁杰","黑崎一户","梁诗诗",
                "隔壁小朋友","派派","堡血公司","文丑丑",
                "脏狗子",
                "汪毅公司",
                "好奇鼠",
                "小张"
            };
            string[] Level1Seed = {
                "当月",
                "目前工程进度兑现",
                "上月差旅结算",
                "拆家",
                "搬运设备",
                "去年"
            };
            string[] Level2Seed = {
                "工资",
                "奖金",
                "医疗补贴",
                "货款",
                "工程款",
                "设备购入"
            };
            string infoString = string.Format(InfoStringFormat
                , Level0Seed[rand.Next(0, Level0Seed.Length)]
                , Level1Seed[rand.Next(0, Level1Seed.Length)]
                , Level2Seed[rand.Next(0, Level2Seed.Length)]);

            double credit = 0; double debit = 0;
            if (rand.Next(0, 1) == 0)
            {
                //Generate Debit
                debit = Math.Round(100000d * rand.NextDouble(), 2);

            }
            else
            {
                //Generate Credit
                credit = Math.Round(100000d * rand.NextDouble(), 2);
            }
            string direction = "贷方";
            double remain = Math.Round(100000d * rand.NextDouble(), 2); ;
            LedgerItem ledger = new LedgerItem(dateIncured, infoString, credit, debit, direction, remain);
            return ledger;
        }
        [TestMethod()]
        public void AddTest()
        {
            LedgerBook book = new LedgerBook(LedgerSides.FromBank);
            for (int i = 0; i < 15; i++)
            {
                book.Add(GenerateRandom());
            }

            foreach (LedgerItem item in book)
            {
                Console.WriteLine(item.ToString());
            }
        }

        [TestMethod()]
        public void ClearTest()
        {
            LedgerBook book = new LedgerBook(LedgerSides.FromBank);
            for (int i = 0; i < 15; i++)
            {
                book.Add(GenerateRandom());
            }

            foreach (LedgerItem item in book)
            {
                Console.WriteLine(item.ToString());
            }

            Console.WriteLine("测试 book 生成完毕，正在开始 Clear() 测试，当前 book 的 Count 为：{0}", book.Count);
            book.Clear();
            Console.WriteLine("Clear() 完成，当前 book 的 Count 为：{0}", book.Count);
            Console.WriteLine("进行清除后的循环");
            foreach (LedgerItem item in book)
            {
                Console.WriteLine(item.ToString());
            }
            Console.WriteLine("循环完毕");
        }

        [TestMethod()]
        public void ContainsTest()
        {
            Random rand = new Random();
            int CaptureIndex = rand.Next(0, 15);
            LedgerItem CapturedLedger = null;
            Console.WriteLine("生成的随机捕获Index为{0}", CaptureIndex);
            LedgerBook book = new LedgerBook(LedgerSides.FromBank);
            for (int i = 0; i < 15; i++)
            {
                LedgerItem ledger = GenerateRandom();
                if (i == CaptureIndex)
                {
                    CapturedLedger = ledger;
                }
                book.Add(ledger);
            }

            foreach (LedgerItem item in book)
            {
                Console.WriteLine(item.ToString());
            }
            Console.WriteLine("测试 book 生成完毕，正在开始 Contains() 肯定测试");
            bool DoesContain = book.Contains(CapturedLedger);
            Console.WriteLine("测试完毕，book 是否含有\r\n{0}\r\n => {1}",
                CapturedLedger.ToString(),
                DoesContain);

            CapturedLedger = GenerateRandom();
            Console.WriteLine("正在开始 Contains() 否定测试");
            DoesContain = book.Contains(CapturedLedger);
            Console.WriteLine("测试完毕，book 是否含有\r\n{0}\r\n => {1}",
                CapturedLedger.ToString(),
                DoesContain);

        }

        [TestMethod()]
        public void CopyToTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void GetEnumeratorTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void IndexOfTest()
        {
            Random rand = new Random();
            int CaptureIndex = rand.Next(0, 15);
            LedgerItem CapturedLedger = null;
            Console.WriteLine("生成的随机捕获Index为{0}", CaptureIndex);
            LedgerBook book = new LedgerBook(LedgerSides.FromBank);
            for (int i = 0; i < 15; i++)
            {
                LedgerItem ledger = GenerateRandom();
                if (i == CaptureIndex)
                {
                    CapturedLedger = ledger;
                }
                book.Add(ledger);
            }

            foreach (LedgerItem item in book)
            {
                Console.WriteLine(item.ToString());
            }
            Console.WriteLine("测试 book 生成完毕，正在开始 IndexOf() 肯定测试");
            int CapturedIndex = book.IndexOf(CapturedLedger);
            Console.WriteLine("测试完毕，book 中的\r\n{0}\r\n 所在的序列号为 {1}",
                CapturedLedger.ToString(),
                CapturedIndex);

            Console.WriteLine("正在开始 IndexOf() 否定测试");
            CapturedLedger = GenerateRandom();
            CapturedIndex = book.IndexOf(CapturedLedger);
            Console.WriteLine("测试完毕，book 中的\r\n{0}\r\n 所在的序列号为 {1}",
                CapturedLedger.ToString(),
                CapturedIndex);
        }

        [TestMethod()]
        public void InsertTest()
        {
            LedgerBook book = new LedgerBook(LedgerSides.FromBank);
            for (int i = 0; i < 15; i++)
            {
                book.Add(GenerateRandom());
            }

            foreach (LedgerItem item in book)
            {
                Console.WriteLine(item.ToString());
            }
            Random rnd = new Random();
            int InsertIndex = rnd.Next(0, 15);
            LedgerItem InsertedItem = GenerateRandom();
            Console.WriteLine("测试 book 生成完毕，正在开始 Insert() 测试，在序列号{0}中插入\r\n{1}"
                , InsertIndex, InsertedItem);

            book.Insert(InsertIndex, InsertedItem);

            Console.WriteLine("再次进行book循环");
            foreach (LedgerItem item in book)
            {
                Console.WriteLine(item.ToString());
            }
            Console.WriteLine("循环完毕");
        }

        [TestMethod()]
        public void RemoveTest()
        {
            LedgerBook book = new LedgerBook(LedgerSides.FromBank);
            for (int i = 0; i < 15; i++)
            {
                book.Add(GenerateRandom());
            }

            foreach (LedgerItem item in book)
            {
                Console.WriteLine(item.ToString());
            }
            Random rnd = new Random();
            int RemoveIndex = rnd.Next(0, 15);

            Console.WriteLine("测试 book 生成完毕，正在开始 Remove() 测试，在序列号{0}中移除"
                , RemoveIndex);

            book.Remove(book[RemoveIndex]);

            Console.WriteLine("再次进行book循环");
            foreach (LedgerItem item in book)
            {
                Console.WriteLine(item.ToString());
            }
            Console.WriteLine("循环完毕");
        }

        [TestMethod()]
        public void RemoveAtTest()
        {
            LedgerBook book = new LedgerBook(LedgerSides.FromBank);
            for (int i = 0; i < 15; i++)
            {
                book.Add(GenerateRandom());
            }

            foreach (LedgerItem item in book)
            {
                Console.WriteLine(item.ToString());
            }
            Random rnd = new Random();
            int RemoveIndex = rnd.Next(0, 15);

            Console.WriteLine("测试 book 生成完毕，正在开始 Remove() 测试，在序列号{0}中移除"
                , RemoveIndex);

            book.RemoveAt(RemoveIndex);

            Console.WriteLine("再次进行book循环");
            foreach (LedgerItem item in book)
            {
                Console.WriteLine(item.ToString());
            }
            Console.WriteLine("循环完毕");
        }
    }
}