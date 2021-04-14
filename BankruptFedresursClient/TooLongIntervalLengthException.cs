using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankruptFedresursClient
{
    public class TooLongIntervalLengthException : InvalidIntervalException
    {
        public TooLongIntervalLengthException() : base()
        {

        }
        public TooLongIntervalLengthException(string message) : base(message)
        {

        }
    }
}
