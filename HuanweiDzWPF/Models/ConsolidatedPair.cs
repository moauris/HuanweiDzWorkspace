using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace HuanweiDzWPF.Models
{
    public class ConsolidatedPair : INotifyPropertyChanged
    {
        /*
        public ConsolidatedPair()
        {
            CompanySide = new ObservableCollection<LedgerItem>();
            BankSide = new ObservableCollection<LedgerItem>();
        }
        */
        public ConsolidatedPair(
            ObservableCollection<LedgerItem> companyCollection, 
            ObservableCollection<LedgerItem> bankCollection)
        {
            CompanyLedgerCollection = companyCollection;
            BankLedgerCollection = bankCollection;
        }


        private ObservableCollection<LedgerItem> companySideLedgerItems;

        public ObservableCollection<LedgerItem> CompanyLedgerCollection
        {
            get { return companySideLedgerItems; }
            set 
            { 
                companySideLedgerItems = value;
                OnPropertyChanged("CompanyLedgerCollection");
            }
        }

        private ObservableCollection<LedgerItem> bankSideLedgerItems;


        public ObservableCollection<LedgerItem> BankLedgerCollection
        {
            get { return bankSideLedgerItems; }
            set 
            { 
                bankSideLedgerItems = value;
                OnPropertyChanged("BankLedgerCollection");
            }
        }

        public double CompanyDebitRemain
        {
            get
            {
                double remain = 0;
                foreach (LedgerItem item in CompanyLedgerCollection)
                {
                    remain += item.Debit;
                    remain -= item.Credit;
                }
                return remain;
            }
        }
        public double BankDebitRemain
        {
            get
            {
                double remain = 0;
                foreach (LedgerItem item in BankLedgerCollection)
                {
                    remain += item.Debit;
                    remain -= item.Credit;
                }
                return remain;
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
