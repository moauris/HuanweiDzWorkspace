using System;
using System.Reflection;

namespace HuanweiDzModels
{
    public class LedgerItem
    {
        #region Constructor
        public LedgerItem(LedgerSides side
            , DateTime date, string info, double credit
            , double debit, string direction
            , double remain)
        {
            Side = side;
            DateIncured = date;
            Info = info; Credit = credit; Debit = debit; Direction = direction;
            RemainingFund = remain;
        }
        public LedgerItem(LedgerSides side, object[] parameters)
        {
            //判定是否正好具有6个元素
            if (parameters.Length != 6) throw new TargetParameterCountException("试图生成的参数数量不正确。");
            Side = side;
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
        public LedgerSides Side { get; set; }
        public DateTime? DateIncured { get; set; }
        public string IncuredOn
        { 
            get
            {
                if (DateIncured is null) return "未知";
                return ((DateTime)DateIncured).ToString("yyyy年MM月dd日");
            }
        }
       
        public string Info { get; set; }
        public double Credit { get; set; }
        public double Debit { get; set; }
        public string Direction { get; set; }
        public double RemainingFund { get; set; }
        #endregion

        #region Methods
        public static bool IsValid(object[] Items)
        {
            //TODO: Make new Validation method for LedgerItem 
            throw new NotImplementedException();
        }
        public static LedgerItem BuildIfValid(LedgerSides side, object[] Items)
        {
            if (IsValid(Items)) return new LedgerItem(side, Items);
            return null;
        }
        #endregion

        #region Events

        #endregion
    }

    public enum LedgerSides
    {
        FromCompany, FromBank
    }
}
