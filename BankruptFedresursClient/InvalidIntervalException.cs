using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankruptFedresursClient
{
    public class InvalidIntervalException : Exception
    {
        public InvalidIntervalException() : base()
        {

        }
        public InvalidIntervalException(string message) : base(message)
        {

        }
    }
}
