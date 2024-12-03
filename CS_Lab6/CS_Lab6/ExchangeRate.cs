using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CS_Lab6
{
    internal class ExchangeRate
    {
        public string letterCode;
        public double exchangeRate;
        public string fullName;

        public ExchangeRate()
        {
            letterCode = "";
            exchangeRate = 0;
            fullName = "";
        }

        public ExchangeRate(string letterCode, double exchangeRate, string fullName)
        {
            this.letterCode = letterCode;
            this.exchangeRate = exchangeRate;
            this.fullName = fullName;
        }

        public override string ToString()
        {
            return string.Join(" ", letterCode, exchangeRate.ToString(), fullName);
        }
    }
}
