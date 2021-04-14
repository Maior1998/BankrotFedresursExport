using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankruptFedresursClient
{
    public class NegativeIntervalLengthException : InvalidIntervalException
    {
        public NegativeIntervalLengthException() : base()
        {

        }
        public NegativeIntervalLengthException(string message) : base(message)
        {

        }
    }
}
