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
            MatchedCollection = new ObservableCollection<ConsolidatedPair>();
#if DEBUG
            //试图建立一个针对DEBUG模式的ViewModel，其中含有一些用于对比的pariedItem
            LedgerItem item1 = LedgerGenerator.GetRandomItem();
            LedgerItem item2 = LedgerGenerator.GetRandomItem();
            LedgerItem item3 = LedgerGenerator.GetRandomItem();
            LedgerItem item4 = LedgerGenerator.GetRandomItem();
            LedgerItem item5 = LedgerGenerator.GetRandomItem();

            var itemcollection1 = new ObservableCollection<LedgerItem>
            { item1, item2, item3 };
            var itemcollection2 = new ObservableCollection<LedgerItem>
            { item4, item5};

            MatchedCollection = new ObservableCollection<ConsolidatedPair>
            { new ConsolidatedPair(itemcollection1, itemcollection2) };



#endif
        }

        private ObservableCollection<ConsolidatedPair> matchedcollection;

        public ObservableCollection<ConsolidatedPair> MatchedCollection
        {
            get { return matchedcollection; }
            set 
            { 
                matchedcollection = value;
                OnPropertyChanged("MatchedCollection");
            }
        }

        private ObservableCollection<LedgerItem> bankLedger;

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
