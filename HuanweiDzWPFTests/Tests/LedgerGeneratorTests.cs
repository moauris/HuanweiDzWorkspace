using Microsoft.VisualStudio.TestTools.UnitTesting;
using HuanweiDzWPF.Tests;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HuanweiDzWPF.Models;

namespace HuanweiDzWPF.Tests.Tests
{
    [TestClass()]
    public class LedgerGeneratorTests
    {
        [TestMethod()]
        public void NextTest()
        {
            LedgerBook book = LedgerGenerator.GetRandomBook(30, LedgerSides.FromBank);
            foreach (LedgerItem item in book)
            {
                Console.WriteLine(item);
            }
        }
    }
}