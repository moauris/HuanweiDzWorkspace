using HuanweiDzWPF.Models;
using HuanweiDzWPF.Tests;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace HuanweiDzWPF.ViewModels
{
    public class MainWindowViewModels : INotifyPropertyChanged
    {
        public MainWindowViewModels()
        {
            CompanyLedger = new ObservableCollection<LedgerItem>();
            BankLedger = new ObservableCollection<LedgerItem>();
        }
        private ObservableCollection<LedgerItem> bankLedger = null;

        public ObservableCollection<LedgerItem> BankLedger
        {
            get { return bankLedger; }
            set
            {
                bankLedger = value;
                OnPropertyChanged("BankLedger");
            }
        }
        private ObservableCollection<LedgerItem> companyLedger;

        public ObservableCollection<LedgerItem> CompanyLedger
        {
            get { return companyLedger; }
            set 
            {
                companyLedger = value;
                OnPropertyChanged("CompanyLedger");
            }
        }

        private LedgerItem selectedLedgerItem;

        public LedgerItem SelectedLedgerItem
        {
            get { return selectedLedgerItem; }
            set 
            { 
                selectedLedgerItem = value;
                OnPropertyChanged("SelectedLedgerItem");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string PropertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(PropertyName));
        }
    }
}
