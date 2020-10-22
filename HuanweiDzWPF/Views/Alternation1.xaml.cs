using HuanweiDzWPF.Models;
using HuanweiDzWPF.Services;
using HuanweiDzWPF.Tests;
using HuanweiDzWPF.ViewModels;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
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
                    ViewModel.CompanyLedger.Add(LedgerGenerator.GetRandomItem());
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
            //跳出弹窗，选取文件:
            OpenFileDialog diag = new OpenFileDialog();
            switch (e.Parameter)
            {
                case "Company":
                    diag.Title = "请选择【公司方】账本的 Excel 文件";
                    break;
                case "Bank":
                    diag.Title = "请选择【银行方】账本的 Excel 文件";
                    break;
                default:
                    break;
            }
            diag.Filter = "含有账本的 Excel 文件(.xls)|*.xls";

            if (!(bool)diag.ShowDialog()) return; //如果没有进行选择，退出
            if (!diag.CheckFileExists) return; //如果文件不存在，退出
            string[] filenames = diag.FileNames;
            if (filenames.Count() > 1) return; //如果选中多于一个文件退出。
            var reader = new ExcelReader(filenames[0], (string)e.Parameter);
            switch (e.Parameter)
            {
                case "Company":
                    ViewModel.CompanyLedger = reader.Read();
                    break;
                case "Bank":
                    diag.Title = "请选择【银行方】账本的 Excel 文件";

                    ViewModel.BankLedger = reader.Read();
                    break;
                default:
                    break;
            }
        }

        private void CanExecute_Consolidate(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = ViewModel?.CompanyLedger.Count > 0 && ViewModel?.BankLedger.Count > 0;
            //If the exception throws not referenced to an instance, use the ?. to judge if that exist.


        }

        private void Excuted_Consolidate(object sender, ExecutedRoutedEventArgs e)
        {
            //MessageBox.Show("开始执行对账逻辑");
            //首先进行简单的对比，单笔对单笔

        }
    }
}
