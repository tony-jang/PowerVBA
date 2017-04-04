using PowerVBA.Codes.Attributes;
using PowerVBA.Codes.Extension;
using PowerVBA.Codes.TypeSystem;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

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

            bool IsError = false;


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
                        break;
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
                PostfixExpression.Add(new MiniVarExp(VarLoc));
            }
            void FinishStack()
            {
                while(stack.Count != 0)
                {
                    PostfixExpression.Add(stack.Pop());
                }
            }

            void RaiseError(string ErrorMessage)
            {
                IsError = true;
                this.ErrorMessage = ErrorMessage;
            }

            for (int i = 0; i< inFixExp.Length; i++)
            {
                char ch = cArr[i];
                char nextCh = i + 1 < cArr.Length ? cArr[i + 1] : '\0';

                // 키워드 먼저 분석
                bool Handled = false;

                string Word = "";
                List<char> WordArr = new List<char>();
                bool StringReco = false;

                // Operator이거나 빈칸 전 위치까지 가져오기.
                char[] SubStrArr = inFixExp.Substring(i).ToCharArray();

                for (int m = 0; m < SubStrArr.Length; m++)
                {
                    char mch = SubStrArr[m];
                    char nextmCh = m + 1 < SubStrArr.Length ? SubStrArr[m + 1] : '\0';

                    if (mch == '\"')
                    {
                        // 만약 "가 하나더 필요하면 아래 내용 주석 해제
                        // WordArr.Add('\"');

                        if (StringReco && nextmCh == '\"') m++;
                        else StringReco = !StringReco;
                    }

                    if ((mch == ' ' || mch.IsOperator() || mch.IsBracket()) && !StringReco) break;
                    WordArr.Add(mch);
                }

                Word = new string(WordArr.ToArray());
                
                if (Word.StartsWith("\"") && Word.EndsWith("\""))
                {
                    InsertionString(Word);
                    i += Word.Length;
                    Handled = true;
                }
                else if (Word.StartsWith("\""))
                {
                    MessageBox.Show("DEBUG 오류!\r\n시작점이 문자열 안내 지점이지만 끝점은 명확하지 않음.");
                }
                else if (Word.Length != 0)
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
                            if (nextCh == '>') { i++; InsertionOp(ExpOperator.NotEqual); }
                            if (nextCh == '<') { RaiseError("<<는 올바른 연산자가 아닙니다."); goto Return; }
                            if (nextCh == '=') { i++; InsertionOp(ExpOperator.RightOrEqual); }
                            break;
                        case '>':
                            if (nextCh == '>') { RaiseError(">>는 올바른 연산자가 아닙니다."); goto Return; }
                            if (nextCh == '<') { RaiseError("><는 올바른 연산자가 아닙니다."); goto Return; }
                            if (nextCh == '=') { i++; InsertionOp(ExpOperator.LeftOrEqual); }
                            break;
                        case '=':
                            if (nextCh == '>') { i++; InsertionOp(ExpOperator.LeftOrEqual); }
                            else if (nextCh == '<') { i++; InsertionOp(ExpOperator.RightOrEqual); }
                            else if (nextCh == '=') { RaiseError("==는 올바른 연산자가 아닙니다."); goto Return; }
                            else { InsertionOp(ExpOperator.Equal); }
                            break;
                        default:
                            // Sub 안의 로컬 변수가 우선,
                            // 이후 없으면 함수 검색
                            InsertionVar(Word);

                            break;

                    }
            }
            
            FinishStack();
            Return:
            Error = IsError;
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
                if (itm.GetType() == typeof(MiniStrExp))          strList.Add(((MiniStrExp)itm).String);
                else if(itm.GetType() == typeof(MiniOperatorExp)) strList.Add(((MiniOperatorExp)itm).Operator.GetValue());
                else if (itm.GetType() == typeof(MiniIntExp))     strList.Add(((MiniIntExp)itm).Integer.ToString());
                else if (itm.GetType() == typeof(MiniDoubleExp))  strList.Add(((MiniDoubleExp)itm).Double.ToString());
                else if (itm.GetType() == typeof(MiniBooleanExp)) strList.Add(((MiniBooleanExp)itm).Boolean.ToString());
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
