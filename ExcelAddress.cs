using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Norbit.Srm.RusAgro.GenerateExcelFromXml
{
    /// <summary>
    /// Класс для работы с адресами Excel в формате R1C1 и A1 ссылок
    /// </summary>
    public class ExcelAddress
    {
        private int _col;
        private int _row;
        private string _columnName;
        private string _address;
        private string _addressRC;
        private Regex _regexRC = new Regex(@"R(\d+)C(\d+)");
        private Regex _regex = new Regex(@"(\D+)(\d+)");
        private bool _fixedRow = false;
        private bool _fixedCol = false;

        /// <summary>
        /// Получение номера строки Row, номера столбца Col, ссылки в формате A1 Address и ссылки в формате R1C1 AddressRC
        /// </summary>
        public int Col { get { return _col; } }
        public int Row { get { return _row; } }
        public string Address { get { return _address; } }
        public string AddressRC { get { return _addressRC; } }

        public ExcelAddress(string Address, bool FixedRow = false, bool FixedCol = false)
        {
            Address = Address.ToUpper();
            Match mathes = _regexRC.Match(Address);

            // to-do код далее предполагает, что Address всегда приходит корректная ссылка в одном из двух форматов

            // получили ссылку в формате R1C1
            if (!mathes.Success)
            {
                Address = GetCellRCAddress(Address);
                mathes = _regexRC.Match(Address);
            }


            _row = Convert.ToInt32(mathes.Groups[1].Value);
            _col = Convert.ToInt32(mathes.Groups[2].Value);

            _addressRC = Address;
            _address = GetCellAddress(Address);

            _fixedRow = FixedRow;
            _fixedCol = FixedCol;
        }

        /// <summary>
        /// Получить адрес ячейки в стиле "А1". Необязательные параметры позволяют получить "фиксированное" значение ячейки для использования в формулах
        /// </summary>
        private string GetCellAddress(string Address)
        {
            int dividend = Convert.ToInt32(Address.Substring(Address.IndexOf('C') + 1));
            _columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                _columnName = Convert.ToChar(65 + modulo).ToString() + _columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return string.Format("{0}{1}", _columnName, Address.Substring(Address.IndexOf('R') + 1, Address.IndexOf('C') - Address.IndexOf('R') - 1));
        }

        private string GetCellRCAddress(string Address)
        {
            Address = Address.ToUpper();
            Match mathes = _regex.Match(Address);
            int Col = 0;
            int Count = mathes.Groups[1].Value.Length - 1;

            if (mathes.Success)
            {
                foreach (char Column in mathes.Groups[1].Value)
                {
                    Col += (int)Math.Pow(26, Count) * ((int)Column - 64);
                    Count--;
                }

                return String.Format("R{0}C{1}", mathes.Groups[2].Value, Col);
            }

            return "R1C1";
        }
        public override string ToString()
        {
            string FixedColString = _fixedCol ? "$" : "";
            string FixedRowString = _fixedRow ? "$" : "";

            return string.Format("{0}{1}{2}{3}", FixedColString, _columnName, FixedRowString, _row);
        }
    }
}
