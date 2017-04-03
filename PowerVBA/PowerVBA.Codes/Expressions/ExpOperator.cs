using PowerVBA.Codes.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Expressions
{
    public enum ExpOperator
    {

        [Priority(0)]
        [Value("(")]
        OpenBracket,
        [Priority(0)]
        [Value(")")]
        CloseBracket,

        [Priority(1)]
        [Value("^")]
        Power,

        [Priority(2)]
        [Value("*")]
        Multiply,
        [Priority(2)]
        [Value("/")]
        Divide,

        [Priority(3)]
        [Value("\\")]
        IntDivide,

        [Priority(4)]
        [Value("Mod")]
        Mod,

        [Priority(5)]
        [Value("+")]
        Plus,
        [Priority(5)]
        [Value("-")]
        Minus,

        [Priority(6)]
        [Value("&")]
        Connect,

        [Priority(7)]
        [Value(">")]
        Left,
        [Priority(7)]
        [Value("<")]
        Right,
        [Priority(7)]
        [Value(">=")]
        LeftOrEqual,
        [Priority(7)]
        [Value("<=")]
        RightOrEqual,
        [Priority(7)]
        [Value("=")]
        Equal,
        [Priority(7)]
        [Value("<>")]
        NotEqual,

        [Priority(8)]
        [Value("Not")]
        Not,

        [Priority(9)]
        [Value("And")]
        And,
        [Priority(10)]
        [Value("Or")]
        Or,
        [Priority(10)]
        [Value("Xor")]
        Xor,
        
    }
}
