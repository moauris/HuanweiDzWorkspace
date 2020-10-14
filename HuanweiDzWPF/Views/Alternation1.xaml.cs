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
        public Alternation1()
        {
            InitializeComponent();
            ViewModel = Resources["ViewModel"] as MainWindowViewModels;
            //if (ViewModel == null)
            //{
            //    throw new NullReferenceException("ViewModel 不可以为 NULL");
            //}
        }

        private MainWindowViewModels ViewModel = null;
        private void btnTestAddRandomToComp_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnTestAddRandomToBank_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
