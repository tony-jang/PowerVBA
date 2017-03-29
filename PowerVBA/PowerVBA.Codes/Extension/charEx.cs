using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Extension
{
    static class charEx
    {
        public static bool IsBracket(this char c)
        {
            if (c == '(' || c == ')') return true;

            return false;
        }
        public static bool IsOperator(this char c)
        {
            bool flag = false;

            string OperatorString = @"<>=-+*/&";

            if (OperatorString.Contains(c)) flag = true;

            /* 가능한 조합

            <
            >
            =
            =< (변경 필요)
            <=
            >=
            => (변경 필요)
            <>
            +
            -
            &
            *
            * 
            */
        
            return flag;
        }

        public static bool IsDivision(this char c)
        {
            string DivisionString = "().+-*=";


            if (DivisionString.Contains(c)) return true;

            return false;
        }

        public static bool IsLower(this char c)
        {
            return char.IsLower(c);
        }
        public static bool IsUpper(this char c)
        {
            return char.IsUpper(c);
        }
        public static bool IsLetter(this char c)
        {
            return char.IsLetter(c);
        }
        public static bool IsNumber(this char c)
        {
            return char.IsNumber(c);
        }

        public static bool IsLetterOrDigit(this char c)
        {
            return char.IsLetterOrDigit(c);
        }
        public static bool IsDigit(this char c)
        {
            return char.IsDigit(c);
        }
        public static bool IsWhiteSpace(this char c)
        {
            return char.IsWhiteSpace(c);
        }
        public static bool IsSymbol(this char c)
        {
            return char.IsSymbol(c);
        }
        public static char ToLower(this char c)
        {
            return char.ToLower(c);
        }
        public static char ToUpper(this char c)
        {
            return char.ToUpper(c);
        }
    }
}
