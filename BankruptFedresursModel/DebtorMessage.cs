﻿using System;

namespace BankruptFedresursModel
{
    public enum CourtDecision
    {

    }

    public enum CourtDecisionType
    {

    }

    /// <summary>
    /// Сообщение по долэнику.
    /// </summary>
    public class DebtorMessage
    {
        /// <summary>
        /// Номер в формате GUID.
        /// </summary>
        public string Guid { get; set; }
        /// <summary>
        /// Номер в формате числа (используется в основном в их API).
        /// </summary>
        public string Id { get; set; }
        /// <summary>
        /// Должник, по которому было выложено сообщение.
        /// </summary>
        public Debtor Debtor { get; set; }
        /// <summary>
        /// Дата и время опубликования сообщения (МСК).
        /// </summary>
        public DateTime DatePublished { get; set; }
        /// <summary>
        /// Адрес опубликования сообщения (адрес суда?)
        /// </summary>
        public string Address { get; set; }
        /// <summary>
        /// Арбитражный управляющий, опубликовавший сообщение.
        /// </summary>
        public ArbitrManager Owner { get; set; }

        /// <summary>
        /// Тип сообщения.
        /// </summary>
        public DebtorMessageType Type { get; set; }

        public override string ToString()
        {
            return Guid;
        }
    }

    public class DebtorMessageType
    {
        /// <summary>
        /// Номер этого типа сообщения. Нужен при заполнении в браузере.
        /// </summary>
        public int Id { get; set; }

        public string Name { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}
