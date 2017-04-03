using PowerVBA.Codes.Attributes;
using PowerVBA.Codes.Extension;
using PowerVBA.Codes.TypeSystem;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Expressions
{
    public class Expression
    {
        public Expression(string inFixExpression, List<Error> Errors, int LineNum)
        {
            this.DisplayExpression = inFixExpression;
            this.Errors = Errors;
            this.LineNum = LineNum;
        }

        public static Expression Empty = new Expression("", null, -1);

        public List<Error> Errors;

        public int LineNum;

        public string DisplayExpression;
        public List<Expression> SubExpressions;
        public ExpressionType ExpressionType;

        public object Value;

        public string ErrorMessage;

        /// <summary>
        /// 식을 계산합니다.
        /// </summary>
        /// <param name="codeinfo"></param>
        public void Calculate(CodeInfo codeinfo)
        {
            bool IsError = false;
            List<IMiniExp> PostFixExpression = GetPostFix(codeinfo, out IsError);

            if (IsError) ExpressionType = ExpressionType.ErrorExpression;
            else
            {
                for (int i = 0; i < DisplayExpression.Length; i++)
                {

                }
            }
        }

        public List<IMiniExp> GetPostFix(CodeInfo codeInfo, out bool Error)
        {
            string inFixExp = DisplayExpression;

            Error = false;


            Stack<MiniOperatorExp> stack = new Stack<MiniOperatorExp>();

            List<IMiniExp> PostfixExpression = new List<IMiniExp>();

            char[] cArr = inFixExp.ToCharArray();
            
            void InsertionOp(ExpOperator Operator)
            {

                do
                {
                    if (stack.Count == 0)
                    {
                        stack.Push(new MiniOperatorExp(Operator));
                        break;
                    }

                    if (Operator == ExpOperator.CloseBracket)
                    {
                        while (stack.Peek().Operator != ExpOperator.OpenBracket)
                        {
                            PostfixExpression.Add(stack.Pop());
                        }
                        stack.Pop();
                    }

                    // 기존 값이 같거나 더 큰 경우 스택의 값을 Expression에 추가
                    else if ((stack.Peek().Operator.GetPriority() <= Operator.GetPriority()) 
                                && (Operator != ExpOperator.OpenBracket) 
                                && (Operator != ExpOperator.Not))
                    {
                        PostfixExpression.Add(stack.Pop());
                    }
                    else
                    {
                        stack.Push(new MiniOperatorExp(Operator));
                        break;
                    }
                    
                } while (true);
            }

            void InsertionBool(bool Boolean)
            {
                PostfixExpression.Add(new MiniBooleanExp(Boolean));
            }

            void InsertionString(string String)
            {
                PostfixExpression.Add(new MiniStrExp(String));
            }
            void InsertionVar(string VarLoc)
            {

            }
            void FinishStack()
            {
                while(stack.Count != 0)
                {
                    PostfixExpression.Add(stack.Pop());
                }
            }

            for (int i = 0; i< inFixExp.Length; i++)
            {
                char ch = cArr[i];
                char nextCh = i + 1 < cArr.Length ? cArr[i + 1] : '\0';

                // 키워드 먼저 분석
                bool Handled = false;

                string Word = "";
                List<char> WordArr = new List<char>();
                bool Break = false;
                // Operator이거나 빈칸 전 위치까지 가져오기.
                inFixExp.Substring(i)
                        .ToCharArray()
                        .ToList()
                        .ForEach((_ch)=> 
                             {
                                 if (Break) return;
                                 if (_ch == ' ' || _ch.IsOperator())
                                 {
                                     Break = true;
                                     return;
                                 }
                                 WordArr.Add(_ch);
                             });
                Word = new string(WordArr.ToArray());
                

                if (Word.Length != 0)
                {
                    i += Word.Length - 1;
                    Handled = true;

                    switch (Word)
                    {
                        case "not":   InsertionOp(ExpOperator.Not);  break;
                        case "and":   InsertionOp(ExpOperator.And);  break;
                        case "or":    InsertionOp(ExpOperator.Or);   break;
                        case "xor":   InsertionOp(ExpOperator.Xor);  break;
                        case "mod":   InsertionOp(ExpOperator.Mod);  break;
                        case "true":  InsertionBool(true);           break;
                        case "false": InsertionBool(false);          break;
                        default: Handled = false; break;
                    }
                    
                    
                }

                if (!Handled)
                    switch (ch)
                    {
                        case '(': InsertionOp(ExpOperator.OpenBracket);      break;
                        case ')': InsertionOp(ExpOperator.CloseBracket);     break;
                        case '^': InsertionOp(ExpOperator.Power);            break;
                        case '*': InsertionOp(ExpOperator.Multiply);         break;
                        case '/': InsertionOp(ExpOperator.Divide);           break;
                        case '+': InsertionOp(ExpOperator.Plus);             break;
                        case '-': InsertionOp(ExpOperator.Minus);            break;
                        case '&': InsertionOp(ExpOperator.Connect);          break;
                        case '<':
                            if (nextCh == '>')
                            { i++; InsertionOp(ExpOperator.NotEqual); }
                            if (nextCh == '<')
                            {    
                                ErrorMessage = "<<는 올바른 연산자가 아닙니다.";
                                Error = true; return null;
                            }
                            if (nextCh == '=')
                            { i++; InsertionOp(ExpOperator.RightOrEqual); }

                            break;
                        case '>':
                            if (nextCh == '>')
                            {
                                ErrorMessage = ">>는 올바른 연산자가 아닙니다.";
                                Error = true; return null;
                            }
                            if (nextCh == '<')
                            {
                                ErrorMessage = "><는 올바른 연산자가 아닙니다.";
                                Error = true; return null;
                            }
                            if (nextCh == '=')
                            { i++; InsertionOp(ExpOperator.LeftOrEqual); }
                            break;
                        case '=':
                            if (nextCh == '>')
                            { i++; InsertionOp(ExpOperator.LeftOrEqual); }
                            else if (nextCh == '<')
                            { i++; InsertionOp(ExpOperator.RightOrEqual); }
                            else if (nextCh == '=')
                            {
                                ErrorMessage = "==는 올바른 연산자가 아닙니다.";
                                Error = true; return null;
                            }
                            else
                            {
                                InsertionOp(ExpOperator.Equal);
                            }
                            break;
                        default:


                            break;

                    }
            }

            FinishStack();

            return PostfixExpression;
        }


        public void AddError(ErrorCode code, string[] parameters = null)
        {
            Errors.Add(new Error(ErrorType.Error, code, parameters, new DomRegion(LineNum, 0)));
        }

    }

    public static class ExpressionEx
    {
        public static string GetPostFixStr(this List<IMiniExp> exp)
        {
            List<string> strList = new List<string>();
            foreach (var itm in exp)
            {
                if (itm.GetType() == typeof(MiniStrExp))
                {
                    strList.Add(((MiniStrExp)itm).String);
                }
                else if(itm.GetType() == typeof(MiniOperatorExp))
                {
                    strList.Add(((MiniOperatorExp)itm).Operator.GetValue());
                }
                else if (itm.GetType() == typeof(MiniIntExp))
                {
                    strList.Add(((MiniIntExp)itm).Integer.ToString());
                }
                else if (itm.GetType() == typeof(MiniDoubleExp))
                {
                    strList.Add(((MiniDoubleExp)itm).Double.ToString());
                }
                else if (itm.GetType() == typeof(MiniBooleanExp))
                {
                    strList.Add(((MiniBooleanExp)itm).Boolean.ToString());
                }
            }

            return string.Join(" ", strList.ToArray());
        }
    }

    public enum ExpressionType
    {
        /// <summary>
        /// 문자열 값입니다.
        /// </summary>
        String,
        /// <summary>
        /// True/False 입니다.
        /// </summary>
        Boolean,
        /// <summary>
        /// 정수 값입니다.
        /// </summary>
        Integer,
        /// <summary>
        /// 계산되지 않았습니다.
        /// </summary>
        NonCaclulated,

        ErrorExpression,
    }
}
