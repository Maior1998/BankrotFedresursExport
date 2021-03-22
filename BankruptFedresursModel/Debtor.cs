using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankruptFedresursModel
{
    /// <summary>
    /// Должник.
    /// </summary>
    public class Debtor
    {
        /// <summary>
        /// ФИО.
        /// </summary>
        public string FullName { get; set; }
        /// <summary>
        /// Дата рождения.
        /// </summary>
        public DateTime BirthDate { get; set; }

        public override string ToString()
        {
            return FullName;
        }
    }
}
