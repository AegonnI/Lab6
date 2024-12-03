using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CS_Lab6
{
    internal class Accrual
    {
        public int accountID;
        public int currencyID;
        public DateTime date;
        public double summ;

        public Accrual()
        {
            accountID = 0;
            currencyID = 0;
            date = DateTime.MinValue;
            summ = 0;
        }

        public Accrual(int accountID, int currencyID, DateTime date, double summ)
        {
            this.accountID = accountID;
            this.currencyID = currencyID;
            this.date = date;
            this.summ = summ;
        }

        public override string ToString()
        {
            return string.Join(" ", accountID.ToString(), currencyID.ToString(), date.ToString().Substring(0, 10), summ.ToString());
        }
    }
}
