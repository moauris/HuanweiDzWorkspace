using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using HuanweiDzWPF.Models;

namespace HuanweiDzWPF.Tests
{
    public class LedgerGenerator
    {
        public static LedgerItem GetRandomItem(Random rand)
        {
            DateTime incurDate = new DateTime(rand.Next(2007, 2019), rand.Next(3, 7), rand.Next(2, 29));
            Debug.Print("生成的日期是{0}：", incurDate.ToShortDateString());
            string[] Seed0 =
            {
                "宝宝猫", "入云龙公孙胜", "菜花", "留香", "小李飞刀", "大于然", "基努李维斯", "昆丁大坏蛋", "黑崎一护", "埼玉老师"
            };
            string[] Seed1 =
            {
                "上个月", "本月", "本月", "本月", "本月", "本月", "本月", "本月", "本月", "今年", "去年", "往年", "总共", "以往", "前个月"
            };
            string[] Seed2 =
            {
                "医疗费报销", "工资", "工程款", "项目结算", "赔付", "工资", "奖金", "年终奖"
            };

            string infoString = string.Format("{0}{1}的{2}"
                , Seed0[rand.Next(0, Seed0.Length - 1)]
                , Seed1[rand.Next(0, Seed1.Length - 1)]
                , Seed2[rand.Next(0, Seed2.Length - 1)]);
            Debug.Print("生成的随机信息是：{0}", infoString);
            double credit = 0;
            double debit = 0;

            if (rand.Next(0, 1) == 1)
            {
                credit = Math.Round(100000 * rand.NextDouble(), 2);
            }
            else
            {
                debit = Math.Round(100000 * rand.NextDouble(), 2);
            }
            Debug.Print("生成的随机贷方是：{0}", debit);
            Debug.Print("生成的随机借方是：{0}", credit);
            double remain = Math.Round(1000000 * rand.NextDouble(), 2);
            Debug.Print("生成的随机余额是：{0}", remain);
            LedgerItem result = new LedgerItem(incurDate, infoString, credit, debit, "贷方", remain);
            return result;
        }
        public static LedgerItem GetRandomItem()
        {
            Random rand = new Random();
            return GetRandomItem(rand);
        }

        public static LedgerBook GetRandomBook(int Size, LedgerSides Side)
        {
            LedgerBook book = new LedgerBook(Side);
            
            for (int i = 0; i < Size; i++)
            {
                Thread.Sleep(100); //random 对象被产生的瞬间小于系统时钟时不会有变化。此处需要等待。
                book.Add(GetRandomItem());
            }
            return book;
        }
    }
}
