using HuanweiDzWPF.Models;
using HuanweiDzWPF.Tests;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace HuanweiDzWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private LedgerBook bookCompany;

        public LedgerBook BookCompany
        {
            get { return bookCompany; }
            set { bookCompany = value; OnPropertyChanged("BookCompany"); }
        }
        private LedgerBook bookBank;

        public LedgerBook BookBank
        {
            get { return bookBank; }
            set { bookBank = value; OnPropertyChanged("BookBank"); }
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void btnTestAddRandomToComp_Click(object sender, RoutedEventArgs e)
        {
            BookCompany = LedgerGenerator.GetRandomBook(32, LedgerSides.FromCompany);
            Debug.Print("已经生成了公司侧账本");
            foreach (LedgerItem item in BookCompany)
            {
                Debug.Print(item.ToString());
            }
        }

        private void btnTestAddRandomToBank_Click(object sender, RoutedEventArgs e)
        {
            BookBank = LedgerGenerator.GetRandomBook(32, LedgerSides.FromBank);
        }
    }
}
