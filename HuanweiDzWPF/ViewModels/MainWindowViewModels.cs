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
            LedgerItemCollection = new ObservableCollection<LedgerItem>
            {
                LedgerGenerator.GetRandomItem()
            };
        }

        private ObservableCollection<LedgerItem> ledgerItemCollection;

        public ObservableCollection<LedgerItem> LedgerItemCollection
        {
            get { return ledgerItemCollection; }
            set 
            { 
                ledgerItemCollection = value;
                OnPropertyChanged("LedgerItemCollection");
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
