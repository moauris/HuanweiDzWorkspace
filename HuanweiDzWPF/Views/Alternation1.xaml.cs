using HuanweiDzWPF.Services;
using HuanweiDzWPF.Tests;
using HuanweiDzWPF.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace HuanweiDzWPF.Views
{
    /// <summary>
    /// Interaction logic for Alternation1.xaml
    /// </summary>
    public partial class Alternation1 : Window
    {
        private MainWindowViewModels ViewModel = null;
        public Alternation1()
        {

            InitializeComponent();
            ViewModel = Resources["ViewModels"] as MainWindowViewModels;
            if (ViewModel == null)
            {
                throw new NullReferenceException("ViewModels 不可以为 NULL");
            }
        }


        private void CanExecute_AddRandomLedgerItem(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void Excuted_AddRandomLedgerItem(object sender, ExecutedRoutedEventArgs e)
        {
            switch (e.Parameter)
            {
                case "Bank":
                    MessageBox.Show("唉，还没做呢，还没做！");
                    break;
                case "Company":
                    //执行添加公司侧逻辑
                    ViewModel.LedgerItemCollection.Add(LedgerGenerator.GetRandomItem());
                    break;
                default:
                    break;
            }
        }

        private void CanExecute_ReadExcel(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void Excuted_ReadExcel(object sender, ExecutedRoutedEventArgs e)
        {
            //TODO: 编写读取Excel的行为
            var reader = new ExcelReader(@"C:\Users\40137\source\repos\Huanwei_Account\Samples\1月明细账.xls", (string)e.Parameter);
            ViewModel.LedgerItemCollection = reader.Read();
        }
    }
}
