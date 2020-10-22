using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;

namespace HuanweiDzWPF.Commands
{
    public class RoutedCommands : RoutedCommand
    {
        public static RoutedCommand AddRandomLedgerItemCommand = new RoutedCommand();
        public static RoutedCommand ReadExcelCommand = new RoutedCommand();
        public static RoutedCommand ConsolidateCommand = new RoutedCommand();
    }
}
