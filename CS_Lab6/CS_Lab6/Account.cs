using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CS_Lab6
{
    internal class Account
    {
        public string fullName;
        public DateTime date;

        public Account()
        {
            fullName = "";
            date = DateTime.MinValue;
        }
        public Account(string fullName, DateTime date)
        {
            this.fullName = fullName;
            this.date = date;
        }

        public override string ToString()
        {
            return string.Join(" ", fullName, date.ToString().Substring(0, 10));
        }
    }
}
