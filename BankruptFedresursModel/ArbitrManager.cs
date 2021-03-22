using System;
namespace BankruptFedresursModel
{
    /// <summary>
    /// Арбитражный управляющий.
    /// </summary>
    public class ArbitrManager
    {
        /// <summary>
        /// ФИО.
        /// </summary>
        public string FullName { get; set; }

        public override string ToString()
        {
            return FullName;
        }
    }
}
