using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Excel_Advanced_Search
{
    internal class Helper
    {

        private static readonly Regex LetterRegex = new Regex("[A-Za-z]+", RegexOptions.Compiled);
        private static readonly Regex NumberRegex = new Regex(@"\d+", RegexOptions.Compiled);

        public static string GetSheetName(string fullRef) =>
            fullRef.Split('!')[0];

        public static string GetRangeAddress(string fullRef) =>
            fullRef.Split('!')[1];


        public static (string startLetter, string endLetter) GetRangeLetters(Range range)
        {
            string[] parts = range.Address.Split(':');
            return (ExtractLetters(parts[0]), parts.Length > 1 ? ExtractLetters(parts[1]) : ExtractLetters(parts[0]));
        }

        public static string ExtractLetters(string address) =>
            LetterRegex.Match(address).Value;

        public static int GetRowNumber(string address, int index = 0)
        {
            string[] parts = address.Split(':');
            return (parts.Length > 1) ? Convert.ToInt32(NumberRegex.Match(parts[index]).Value)  : Convert.ToInt32(NumberRegex.Match(parts[0]).Value);
        }
    }
}
