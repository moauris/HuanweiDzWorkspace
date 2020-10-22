using System;
using System.Globalization;
using System.Reflection;

namespace HuanweiDzWPF.Models
{
    public class LedgerItem : IEquatable<LedgerItem>
    {
        #region Constructor
        public LedgerItem(DateTime date
            , string info, double credit
            , double debit, string direction
            , double remain)
        {
            DateIncured = date;
            Info = info; Credit = credit; Debit = debit; Direction = direction;
            RemainingFund = remain;
        }
        public LedgerItem(object[] parameters)
        {
            //判定是否正好具有6个元素
            if (parameters.Length != 6) throw new TargetParameterCountException("试图生成的参数数量不正确。");
            if (parameters[0] is null)
            {
                DateIncured = null;
            }
            else
            {
                DateIncured = (DateTime)parameters[0];
            }
            Info = (string)parameters[1]; 
            Credit = (double)parameters[2]; 
            Debit = (double)parameters[3];
            Direction = (string)parameters[4]; ;
            RemainingFund = (double)parameters[5];
        }
        #endregion

        #region Properties
        public DateTime? DateIncured { get; set; }
        public string IncuredOn
        { 
            get
            {
                if (DateIncured is null) return "未知日期";
                return ((DateTime)DateIncured).ToString("yyyy年M月d日");
            }
        }
       
        public string Info { get; set; }
        public double Credit { get; set; }
        public string CreditAsString
        {
            get => Credit.ToString("C", new CultureInfo("zh-CN"));
        }
        public double Debit { get; set; }
        public string DebitAsString
        {
            
            get => Debit.ToString("C", new CultureInfo("zh-CN"));
        }
        public string Direction { get; set; }
        public double RemainingFund { get; set; }
        public string RemainingFundAsString
        {

            get => RemainingFund.ToString("C", new CultureInfo("zh-CN"));
        }
        #endregion

        #region Methods
        public static bool IsValid(object[] Items)
        {
            //TODO: Make new Validation method for LedgerItem 
            throw new NotImplementedException();
        }
        public static LedgerItem BuildIfValid(object[] Items)
        {
            if (IsValid(Items)) return new LedgerItem(Items);
            return null;
        }

        public override string ToString()
        {
            string format = "{0},\t{1}\t\t\t\t\t\t,￥{2,10},￥{3,10},{4},￥{5,10}";
            string OutString = string.Format(format, IncuredOn, Info, Credit, Debit, Direction, RemainingFund);
            return OutString;
        }

        public bool Equals(LedgerItem other)
        {
            if (this.IncuredOn != other.IncuredOn) return false;
            if (this.Info != other.Info) return false;
            if (this.CreditAsString != other.CreditAsString) return false;
            if (this.DebitAsString != other.DebitAsString) return false;
            if (this.Direction != other.Direction) return false;
            if (this.RemainingFundAsString != other.RemainingFundAsString) return false;
            return true;
        }
        #endregion

        #region Events

        #endregion
    }
}
