using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Codes.Enums;
using PowerVBA.Codes.Expressions;
using PowerVBA.Codes.Extension;
using PowerVBA.Codes.TypeSystem;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace PowerVBA.Codes
{
    class VBASeeker
    {
        public VBASeeker(CodeInfo codeInfo)
        {
            CodeInfo = codeInfo;
        }


        string NamePattern = @"^([_a-zA-z가-힣ㅏ-ㅣㄱ-ㅎ][_a-zA-Z가-힣ㅏ-ㅣㄱ-ㅎ1-9]*)$";
        public static string[] ExpressionReserveWords = Properties.Resources.예약어2.ToLower().Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
        public static string[] ReserveWords = Properties.Resources.예약어.ToLower().Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

        private CodeInfo CodeInfo;

        private LineInfo TempLineInfo;
        

        /// <summary>
        /// 줄을 인식해 오류를 가져옵니다. 단, 처리되지 않은 줄이 있을시 NonHandledLine에 정보를 넣어 반환합니다.
        /// </summary>
        /// <param name="FileName"></param>
        /// <param name="CodeLine"></param>
        /// <param name="Lines"></param>
        /// <param name="Nested"></param>
        /// <returns></returns>
        public NotHandledLine GetLine(string FileName, string CodeLine, (RangeInt,int) Lines, ref LineInfo LineInfo, bool Nested = false)
        {   
            NotHandledLine nHandledLine = NotHandledLine.Empty;
            //  for 카운터
            int i = 0,
                // 여는 괄호 갯수
                bracketCount = 0,
                // 단어 인식 처리 횟수
                WordRecognition = 0;
            (int, int) Offset = (0, 0);

            var data = new CodeData();
            data.IsFistNonWs = true;

            // 저장할 텍스트
            StringBuilder SavingText = new StringBuilder();

            char[] text = CodeLine.ToCharArray();


            TempLineInfo = LineInfo.Clone();

            for (i = 0;i < text.Length; i++)
            {
                char ch = text[i];
                bool Handled = true;
                bool MultiLineRead = false;
                bool IsLastChar = i >= text.Length - 1;
                char nextCh = i + 1 < text.Length ? text[i + 1] : '\0';
                
                // 기본적인 문자
                switch (ch)
                {
                    #region [  string / 주석 / 전처리기 지시문  ]
                    case '#': // 전처리기 지시문
                        if (data.IsInString || data.IsInVerbatimString || data.IsInComment) break;
                        // 첫번째 문자가 #일 경우 전처리기 지시문으로 인식함
                        if (data.IsFistNonWs) data.IsInPreprocessorDirective = true;
                        else AddError(ErrorCode.VB0042);
                        break;
                    case '\'': // 주석
                        // String 중이거나, Comment내부라면
                        if (data.IsInString || data.IsInVerbatimString || data.IsInComment) break;
                        data.IsInComment = true;

                        data.IsInPreprocessorDirective = false;

                        i = text.Length - 1;
                        IsLastChar = i >= text.Length - 1;
                        goto ExitIf;
                    case '"':
                        if (data.IsInPreprocessorDirective || data.IsInComment) break;

                        if ((data.AfterIf && WordRecognition == 1) ||
                            (data.AfterSelect && data.AfterCase && WordRecognition == 2) ||
                            (data.AfterCase && WordRecognition == 1))
                        {
                            i--;
                            if ((data.AfterIf && WordRecognition == 1))
                            {
                                ExpressionRange(i, RecognitionTypes.BeforeThen);
                                IsLastChar = i >= text.Length - 1;
                            }
                            else
                            {
                                ExpressionRange(i, RecognitionTypes.ToEndLine);
                                IsLastChar = i >= text.Length - 1;
                            }
                            break;
                        }

                        if (nextCh == '"')
                        {
                            i++;
                            break;
                        }

                        data.IsInString = !data.IsInString;
                        if (!data.IsInString) data.AfterString = true;
                        if (data.IsInString) Offset.Item1 = i;
                        else
                        {
                            Offset.Item2 = i;
                        }
                        break;


                    #endregion

                    #region [  새줄 인식/ 멀티 라인 인식  ]
                    case '\n':
                    case '\r': // 새 줄 인식될때 초기화
                        if (!data.UseMultiLine)
                        {
                            data.IsInComment = false;
                            data.IsInString = false;
                            data.IsInPreprocessorDirective = false;
                            data.IsFistNonWs = true;
                        }
                        if (!nextCh.ToString().ContainsWords(new string[] { "\n", "\r" })) data.UseMultiLine = false;
                        break;

                    case '_': // 멀티 줄 인식
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment || SavingText.ToString() != string.Empty)
                            break;
                        data.UseMultiLine = true;

                        break;


                    case ':':
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) goto ExitIf;

                        if (!Nested && WordRecognition == 0)
                        {
                            if (ReserveWords.Contains(SavingText.ToString().ToLower()))
                            {
                                // 예약어가 포함되어 있을때의 오류
                                AddError(ErrorCode.VB0047, new string[] { SavingText.ToString() });
                            }
                            else
                            {
                                // 정상적인 아이템

                            }

                            data.AfterLabel = true;
                        }

                        Handled = false;
                        MultiLineRead = true;

                        break;

                    #endregion

                    #region [  괄호 (여닫는 괄호)  ]
                    case '(': // 여는 괄호
                        
                        
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment)
                            break;
                        
                        if (data.AfterDeclarator) AddError(ErrorCode.VB0050);

                        if (data.AfterReDim && data.AfterIdentifier)
                        {
                            string expression = ExpressionRange(i, RecognitionTypes.BracketExpression);
                            if (string.IsNullOrEmpty(expression)) AddError(ErrorCode.VB0155);
                        }
                            
                        if (data.AfterIdentifier)
                        {
                            // Array
                            if (data.IsVarDeclaring)
                            {
                                data.AfterArray = true;
                            }
                            else if (data.AfterCallFunction)
                            {
                                ParameterUseExpression(i);
                            }
                            // 파라미터 (VBA에서 Public Sub A 나 Public Function A 이후에 괄호가 나오면 파라미터의 시작이다.
                            else
                            {
                                ParameterSeek(i);
                                IsLastChar = i >= text.Length - 1;
                            }
                        }
                        bracketCount++;

                        data.IsInBracket = true;
                        break;
                    case ')': // 닫는 괄호
                        
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment)
                            break;
                        
                        if (bracketCount <= 0)
                        {
                            AddError(ErrorCode.VB0051);
                            data.IsInBracket = false;
                        }
                        else bracketCount--;
                        
                        if (data.AfterArray)
                        {
                            // Array 선언이 끝난뒤에는 다시 식별자로 인식 될 수 있도록 함
                            data.AfterArray = false;
                            data.AfterIdentifier = true;
                        }

                        if (data.AfterArray && bracketCount <= 0) data.AfterArray = false;

                        break;
                    #endregion

                    #region [  연산자  ]

                    case '<': case '>':
                    case '+': case '-':
                    case '*': case '/':
                    case '!':
                    case '=':
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) break;

                        if ((data.AfterIf || data.AfterDo || data.AfterWhile) && !data.AfterExpression && !data.AfterEnd)
                        {
                            AddError(ErrorCode.VB0078);
                            break;
                        }

                        // 다음 글자도 Operator일시 중복 Operator로 인식해서 처리
                        if (nextCh.IsOperator())
                        {
                            string multiOperator = ch.ToString() + nextCh.ToString();

                            switch (multiOperator)
                            {
                                case "<>":
                                    
                                    break;
                                case "=<":
                                case "<=":
                                    if (multiOperator.StartsWith("="))
                                        AddWarning(ErrorCode.VB0060, new string[] { multiOperator, "<=" });

                                    
                                    break;
                                case "=>":
                                case ">=":
                                    if (multiOperator.StartsWith("="))
                                        AddWarning(ErrorCode.VB0060, new string[] { multiOperator, ">=" });
                                    

                                    break;
                                case "*=": case "/=":
                                case "-=": case "+=":
                                    AddError(ErrorCode.VB0009, new string[] { multiOperator, $"Object = Object { multiOperator.Substring(0,1) } Value" });
                                    break;
                                case "!=":
                                    AddError(ErrorCode.VB0009, new string[] { multiOperator, $"Left <> Right" });
                                    break;
                            }
                        }
                        else
                        {
                            if (data.AfterType)
                            {

                            }
                        }
                        if ((data.AfterIdentifier && data.IsVarDeclaring) && !data.AfterArray)
                        {
                            AddError(ErrorCode.VB0008);
                        }
                        break;

                    #endregion

                    #region [  특수 문자  ]

                    case ',':
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) break;

                        if (data.IsVarDeclaring && data.AfterIdentifier)
                        {
                            data.AfterIdentifier = false;
                        }
                        else
                        {
                            AddError(ErrorCode.VB0041, new string[] { ch.ToString() });
                        }
                        break;


                    case ';': case '@': case '$': case '^': case '%': case '~': // 사용되지 않는 특수문자
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) break;

                        AddError(ErrorCode.VB0041, new string[] { ch.ToString() });
                        break;

                    case '\\':
                    case '&':
                    case '|':
                        if (WordRecognition == 0) AddError(ErrorCode.VB0041, new string[] { ch.ToString() });
                        break;
                    #endregion

                    //이외 체크
                    default:
                        
                        Handled = false;
                        break;
                }

                // switch문에서 처리되지 않은 경우
                if (!Handled)
                {
                    (int, int) CurrentOffset = (i - SavingText.Length - 1, i);

                    // 빈칸이거나 마지막 단어 일경우 또는 읽을 필요가 있는 경우
                    // 또는 다음 글자가 '('나 ')' 같이 읽을 필요가 있는 경우
                    if (ch.IsWhiteSpace() || IsLastChar || nextCh.IsDivision() || MultiLineRead)
                    {
                        // string이거나 전처리기 지시문이거나 코멘트일 경우 넘김
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) goto ExitIf;

                        // 마지막 글자이면서 빈칸이 아니고 멀티 라인이 아니다라면 인식할 문자열에 추가
                        if ((IsLastChar || nextCh.IsDivision()) && !(ch.IsWhiteSpace() || ch.IsOperator() || ch.IsBracket())
                            && !MultiLineRead) SavingText.Append(ch);

                        if (SavingText.Length == 0) goto ExitIf;


                        // 기본적인 오류는 여기서 잡아냄
                        if (ErrorCheck(SavingText.ToString().ToLower(), CurrentOffset))
                        {
                            SavingText.Clear();
                            goto ExitIf;
                        }

                        switch (SavingText.ToString().ToLower())
                        {
                            /*
                            Option Compare {Binary|Text} 

                             */

                            #region [  Option  ]

                            case "option":
                                data.AfterOption = true;

                                break;
                            case "explicit":
                                if (!data.AfterOption) AddError(ErrorCode.VB0240);

                                break;
                            case "compare":

                                break;
                            case "text":

                                break;
                            case "binary":

                                break;
                            case "base":
                            case "module":
                                if (SavingText.ToString() == "module") AddError(ErrorCode.VB0002);

                                break;

                            #endregion
                                
                            #region [  Accessor  ]
                            case "public":
                                if (data.AfterAccessor) AddError(ErrorCode.VB0022);

                                data.AfterAccessor = true;
                                break;
                            case "private":
                                // 엑세서로 서의 private
                                if (data.AfterAccessor) AddError(ErrorCode.VB0022);

                                data.AfterAccessor = true;

                                // TODO : Option Private Module 구현
                                break;
                            case "dim":
                                if (data.AfterAccessor) AddError(ErrorCode.VB0022);
                                data.AfterAccessor = true;
                                data.IsVarDeclaring = true;

                                break;
                            case "const":
                                if (data.IsVarDeclaring && data.AfterAccessor) AddError(ErrorCode.VB0022);
                                if (LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0144);

                                data.AfterAccessor = true;
                                data.IsConstDeclaring = true;
                                break;

                            case "redim":

                                if (data.AfterAccessor) AddError(ErrorCode.VB0021);
                                if (WordRecognition != 0) AddError(ErrorCode.VB0153);

                                data.AfterAccessor = true;
                                data.AfterReDim = true;

                                break;

                            #endregion

                            // TODO List ::
                            // Const                                  (   미 확 인  )
                            // API 선언                               (구현되지 않음)
                            // ReDim                                  ( 구 현 완 료 )
                            // On Error Resume Next                   ( 구 현 완 료 )
                            // On Error goto Point                    ( 구 현 완 료 )
                            // 다중 변수 선언                         ( 구 현 완 료 )
                            // Set 이후 오브젝트로 인식               (구현되지 않음)
                            // Sub뒤 파라미터로 정상 인식하게 해주기  (구현되지 않음)

                            #region [  On Error [Goto Label / Resume Next]  ]
                            case "on":
                                if (WordRecognition != 0) AddError(ErrorCode.VB0180);

                                data.AfterOn = true;
                                break;
                            case "error":
                                if (!data.AfterOn) AddError(ErrorCode.VB0181);
                                data.AfterError = true;
                                break;
                            case "goto":
                                if (!(data.AfterOn && data.AfterError)) AddError(ErrorCode.VB0182);
                                data.AfterGoto = true;
                                break;
                            case "resume":
                                if (!(data.AfterOn && data.AfterError)) AddError(ErrorCode.VB0183);
                                data.AfterResume = true;
                                break;

                            #endregion

                            #region [  For or On Error Resume Next  ]

                            case "next": // For에서의 Next와 On Error Resume Next 잡아내기

                                if (!(data.AfterOn && data.AfterError && data.AfterResume))
                                { // For문의 Next로 인식
                                    if (LineInfo.IsInFor)
                                    {
                                        if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "Next" });
                                        TempLineInfo.IsInFor = false;
                                    }
                                    else
                                    {
                                        AddError(ErrorCode.VB0173);
                                        break;
                                    }
                                }
                                else
                                {

                                    if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "On Error Goto Next" });
                                }

                                // 아무런 오류가 없다면 On Error Resume 이후의 Next로 인식

                                data.AfterNext = true;
                                break;

                            #endregion

                            #region [  Declartor  ]

                            case "sub":
                                if (data.AfterEnd) // End Sub 인식
                                {
                                    if (WordRecognition == 1)
                                    {
                                        if (LineInfo.IsInSub)
                                        {
                                            TempLineInfo.IsInSub = false; // End Sub 정상 인식
                                            if (LineInfo.TempLine.Peek().Item2 == FoldingTypes.Sub)
                                            {
                                                LineInfo.Foldings.Add(new RangeInt() { StartInt = LineInfo.TempLine.Pop().Item1, EndInt = Lines.Item1.StartInt });
                                            }
                                        }
                                        else AddError(ErrorCode.VB0139);                    // Sub가 없는 End Sub
                                    }
                                    else AddError(ErrorCode.VB0057);

                                }
                                else if (data.AfterDeclare && !data.AfterExpression && !data.AfterLib)
                                {
                                    // 정상적 인식
                                }
                                else if (data.AfterExit)
                                {
                                    if (!LineInfo.IsInSub) AddError(ErrorCode.VB0127, new string[] { "Sub" });
                                }
                                else
                                {
                                    if (LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0133); // Nest 오류
                                    else if (LineInfo.IsInEnum) AddError(ErrorCode.VB0136);
                                    else
                                    {
                                        if (LineInfo.IsGlobalVarDeclaring)
                                        {
                                            TempLineInfo.IsGlobalVarDeclaring = false;
                                            TempLineInfo.LastGlobalVarInt = Lines.Item1.StartInt - 1;
                                        }
                                        TempLineInfo.IsInSub = true;
                                        LineInfo.TempLine.Push((Lines.Item1.StartInt, FoldingTypes.Sub));
                                    }
                                }

                                data.AfterSub = true;
                                data.AfterDeclarator = true;

                                break;
                            case "function":
                                if (data.AfterEnd)
                                {
                                    if (WordRecognition == 1)
                                    {
                                        if (LineInfo.IsInFunction)
                                        {
                                            TempLineInfo.IsInFunction = false;
                                            if (LineInfo.TempLine.Peek().Item2 == FoldingTypes.Function)
                                            {
                                                LineInfo.Foldings.Add(new RangeInt(LineInfo.TempLine.Pop().Item1) { EndInt = Lines.Item1.StartInt });
                                            }
                                        }
                                        else AddError(ErrorCode.VB0138);
                                    }
                                    else AddError(ErrorCode.VB0057);
                                }
                                else if (data.AfterDeclare && !data.AfterExpression && !data.AfterLib)
                                {
                                    // 정상적 인식
                                }
                                else if (data.AfterExit)
                                {
                                    if (!LineInfo.IsInFunction) AddError(ErrorCode.VB0127, new string[] { "Function" });
                                }
                                else
                                {
                                    // Nest 오류
                                    if (LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0132);
                                    else if (LineInfo.IsInEnum) AddError(ErrorCode.VB0135);
                                    else
                                    {
                                        if (LineInfo.IsGlobalVarDeclaring)
                                        {
                                            TempLineInfo.IsGlobalVarDeclaring = false;
                                            TempLineInfo.LastGlobalVarInt = Lines.Item1.StartInt - 1;
                                        }
                                        TempLineInfo.IsInFunction = true;
                                        LineInfo.TempLine.Push((Lines.Item1.StartInt, FoldingTypes.Function));
                                    }
                                }
                                data.AfterFunction = true;
                                data.AfterDeclarator = true;

                                break;
                            case "property":

                                if (data.AfterDeclare) { AddError(ErrorCode.VB0201); break; }

                                if (data.AfterEnd)
                                {
                                    if (LineInfo.IsInProperty)
                                    {
                                        TempLineInfo.IsInProperty = false;
                                        LineInfo.Foldings.Add(new RangeInt(LineInfo.TempLine.Pop().Item1) { EndInt = Lines.Item1.StartInt });
                                    }
                                    else AddError(ErrorCode.VB0140);
                                }
                                else if (data.AfterExit)
                                {
                                    if (!LineInfo.IsInFunction) AddError(ErrorCode.VB0127, new string[] { "Property" });
                                }
                                else
                                {
                                    if (LineInfo.IsNestInProcedure) // Nest Error
                                    { AddError(ErrorCode.VB0134); }
                                    if (LineInfo.IsInEnum)
                                    { AddError(ErrorCode.VB0137); }
                                    else
                                    {
                                        if (LineInfo.IsGlobalVarDeclaring)
                                        {
                                            // Lines.Item1.EndInt
                                            TempLineInfo.IsGlobalVarDeclaring = false;
                                            TempLineInfo.LastGlobalVarInt = Lines.Item1.StartInt - 1;
                                        }
                                        TempLineInfo.IsInProperty = true;
                                        LineInfo.TempLine.Push((Lines.Item1.StartInt, FoldingTypes.Property));
                                    }
                                }

                                data.AfterProperty = true;
                                data.AfterDeclarator = true;
                                break;
                            case "declare":
                                if (data.AfterDeclarator) AddError(ErrorCode.VB0200);
                                data.AfterDeclare = true;
                                data.AfterDeclarator = true;
                                break;
                            case "enum":
                                if (data.AfterEnd)
                                {
                                    if (LineInfo.IsInEnum)
                                    {
                                        TempLineInfo.IsInEnum = false;
                                    }
                                }
                                else
                                {

                                }

                                data.AfterEnum = true;
                                data.AfterDeclarator = true;
                                break;
                            #endregion

                            #region [  API  ]

                            case "lib":
                                // Public/Private Declare Sub SubName Lib "LibName" Alias "AliasName" (argument list)
                                // Public / Private Declare Function FunctionName Lib "Libname" alias "aliasname"(argument list) As Type

                                if (data.AfterDeclare && (data.AfterSub || data.AfterFunction) && data.AfterIdentifier)
                                {
                                    if (!data.AfterLib) data.AfterLib = true;
                                    else AddError(ErrorCode.VB0208, new string[] { "Lib" });
                                }
                                else
                                {
                                    AddError(ErrorCode.VB0206);
                                }
                                break;
                            case "alias":
                                if (data.AfterDeclare && (data.AfterSub || data.AfterFunction) && data.AfterIdentifier && data.AfterLib)
                                {
                                    if (!data.AfterAlias)
                                    {
                                        data.AfterAlias = true;
                                        data.AfterString = false;
                                    }
                                    else AddError(ErrorCode.VB0208, new string[] { "Alias" });
                                }
                                else
                                {
                                    AddError(ErrorCode.VB0207);
                                }

                                break;

                            #endregion

                            #region [  Conditional  ]
                            case "if":
                                if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "If" });
                                if (data.AfterEnd)
                                {
                                    if (LineInfo.IsInIf) TempLineInfo.IsInIf = false;
                                    else AddError(ErrorCode.VB0079);
                                }
                                else TempLineInfo.IsInIf = true;

                                data.AfterIf = true;

                                break;
                            case "elseif":
                                if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "ElseIf" });
                                data.AfterElseIf = true;

                                break;
                            case "else":
                                if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "Else" });
                                data.AfterElse = true;
                                break;
                            case "select":
                                if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "Select" });
                                data.AfterSelect = true;
                                break;
                            case "case":
                                if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "Case" });
                                data.AfterCase = true;
                                break;
                            case "then":
                                if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "Then" });
                                data.AfterThen = true;
                                break;
                            #endregion

                            #region [  Iterative  ]

                            #region [  While / Do Until / Do While  ]

                            case "while":
                                if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "While" });
                                // Do이후 절 이거나 Loop 이후절일때
                                if (((data.AfterDo || data.AfterLoop) && WordRecognition == 1) || WordRecognition == 0)
                                {
                                    data.AfterWhile = true;
                                }
                                if (data.AfterEnd)
                                {
                                    AddError(ErrorCode.VB0005);
                                }
                                break;
                            case "until":
                                if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "Until" });

                                // Do Until or Loop Until
                                if ((data.AfterDo || data.AfterLoop) && WordRecognition == 1)
                                { data.AfterUntil = true; }
                                else
                                { AddError(ErrorCode.VB0081); }

                                break;
                            case "do":
                                if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "Do" });
                                if (WordRecognition == 0)
                                {
                                    TempLineInfo.IsInDo = true;
                                    data.AfterDo = true;
                                }
                                else if (data.AfterExit)
                                {
                                    data.AfterDo = true;
                                }

                                break;
                            case "loop":
                                if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "Loop" });
                                if (!LineInfo.IsInDo) AddError(ErrorCode.VB0087);

                                // 이후 Loop Until 또는 Loop While로 식을 사용 가능한 것을 인식 추가
                                if (WordRecognition != 0) AddError(ErrorCode.VB0084);
                                else
                                {
                                    TempLineInfo.IsInDo = false;
                                    data.AfterLoop = true;
                                }
                                break;
                            // 특이한 End While의 형태
                            case "wend":
                                if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "Wend" });
                                if (!LineInfo.IsInWhile) AddError(ErrorCode.VB0086);

                                if (WordRecognition == 0)
                                {
                                    data.AfterWend = true;

                                }
                                else AddError(ErrorCode.VB0082);
                                break;
                            #endregion

                            #region [  For / For Each  ]

                            case "for":
                                if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "For" });
                                if (WordRecognition == 0)
                                {
                                    data.AfterFor = true;
                                    TempLineInfo.IsInFor = true;
                                }
                                break;
                            case "each":
                                if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "Each" });
                                if (data.AfterFor)
                                {
                                    data.AfterFor = false;
                                    data.AfterForEach = true;
                                    TempLineInfo.IsInForEach = true;
                                }
                                else
                                {
                                    // For 이후가 아니라면 Each는 현재 위치에서 사용 불가라는 걸 알려주기
                                    AddError(ErrorCode.VB0171);
                                }
                                break;
                            case "in":
                                if (data.AfterForEach && data.AfterIdentifier)
                                {
                                    data.AfterIn = true;
                                    data.AfterIdentifier = false;
                                }
                                else
                                {
                                    AddError(ErrorCode.VB0174);
                                }
                                break;

                            #endregion

                            #endregion


                            #region [  With  ]
                            case "with":
                                if (data.AfterEnd)
                                {
                                    if (LineInfo.IsInWith) LineInfo.IsInWith = false;
                                    else AddError(ErrorCode.VB0230);
                                }

                                data.AfterWith = true;
                                break;
                            #endregion


                            #region [  Get/Set/Let  ]
                            case "get":
                                // 오류는 이미 확인
                                if (data.AfterProperty && !data.AfterPropAccessor)
                                {
                                    data.AfterPropAccessor = true;
                                    data.AfterGet = true;
                                }
                                else AddError(ErrorCode.VB0130, new string[] { "Get" });

                                data.AfterSet = true;
                                break;
                            case "set":
                                // 오류는 이미 확인
                                if (data.AfterProperty && !data.AfterPropAccessor) data.AfterPropAccessor = true;
                                else
                                {
                                    if (WordRecognition != 0) AddError(ErrorCode.VB0130, new string[] { "Set" });
                                    else data.AfterSet = true;
                                }

                                break;
                            case "let":
                                if (data.AfterProperty && !data.AfterPropAccessor)
                                {
                                    data.AfterPropAccessor = true;
                                    data.AfterLet = true;
                                }
                                else AddError(ErrorCode.VB0130, new string[] { "Let" });

                                break;
                            #endregion

                            #region [  Parameter  ]

                            case "byval":
                            case "byref":
                            case "paramarray":
                            case "optional":

                                AddError(ErrorCode.VB0102);
                                break;
                            #endregion

                            #region [  End/As/Exit/Return/Call  ]

                            case "end":

                                if (WordRecognition != 0) AddError(ErrorCode.VB0000);
                                else data.AfterEnd = true;
                                break;

                            case "as":
                                data.AfterAs = true;
                                break;

                            case "exit":
                                data.AfterExit = true;
                                break;
                            case "return":
                                data.AfterReturn = true;
                                break;
                            case "call":
                                data.AfterCallFunction = true;
                                break;

                            #endregion

                            #region [  VB.NET 미호환 문법  ]

                            case "readonly":
                                AddError(ErrorCode.VB0006);
                                break;
                            case "addhandler":
                                AddError(ErrorCode.VB0007);
                                break;
                            #endregion

                            default:

                                // 선언문 이후이거나 선언문이 아니고 엑세서임을 만족하며 | 식별자이후가 아니라면 식별자로 인식
                                if ((data.AfterDeclarator || (!data.AfterDeclarator && data.AfterAccessor)) && !data.AfterIdentifier)
                                {
                                    // 프로퍼티 라면
                                    if (data.AfterProperty)
                                    {
                                        // 프로퍼티 엑세서 이후라면
                                        if (data.AfterPropAccessor)
                                        {
                                            // 정상 인식
                                            data.AfterIdentifier = true;
                                        }
                                        // 아니라면 나가기
                                        else goto ExitIf;
                                    }
                                    else if (data.AfterReDim && WordRecognition == 1)
                                    {
                                        data.AfterIdentifier = true;
                                    }
                                    else
                                    {
                                        // 식별자 인식

                                        // Accessor 이후이지만 변수 선언이 확실하지 않은 (Public, Private) 키워드의 경우
                                        // 프로시져 안에 포함되어 있으면 오류로 간주 (프로시져 내부에서는 Dim을 이용한 변수 선언 밖에 안됨)
                                        if (data.AfterAccessor && !data.IsVarDeclaring && LineInfo.IsNestInProcedure)
                                        {
                                            AddError(ErrorCode.VB0143);
                                        }

                                        data.AfterDeclarator = false;
                                        data.AfterIdentifier = true;
                                    }

                                    // Public/Private/Dim 뒤에 식별자가 나왔을 경우 변수 선언으로 인식 및 처리
                                    if (data.AfterAccessor && !data.IsConstDeclaring)
                                    {
                                        // 첫 부분이 아닌 곳에 변수 선언 하고 있을때
                                        if (!LineInfo.IsGlobalVarDeclaring && !LineInfo.IsNestInProcedure
                                            && !(data.AfterFunction || data.AfterSub || data.AfterProperty))
                                        {
                                            AddError(ErrorCode.VB0145);
                                        }
                                        if (!(data.AfterSub || data.AfterFunction || data.AfterProperty)) data.IsVarDeclaring = true;
                                    }
                                }
                                // For Each 또는 For문 이후라면
                                else if (data.AfterFor || data.AfterForEach)
                                {
                                    // For Each Identifier In 이후
                                    if (data.AfterForEach && data.AfterIdentifier && data.AfterIn)
                                    {
                                        data.AfterIdentifier = true;
                                    }
                                }
                                // 식별자 뒤라면
                                else if (data.AfterIdentifier && !data.AfterAs && !data.AfterArray)
                                {
                                    // For이나 For Each이후의 식별자는 제외
                                    if (data.AfterFor || data.AfterForEach) break;
                                    // As 이어야 함.
                                    if (SavingText.ToString().ToLower() != "as" && !data.AfterAs) AddError(ErrorCode.VB0027, new string[] { "As" });
                                    // 식별자 해제
                                    data.AfterIdentifier = false;
                                }
                                // On Error Goto 이후 절이라면 Label 인식
                                else if (data.AfterOn && data.AfterError && data.AfterGoto)
                                {
                                    if (!LineInfo.IsNestInProcedure) AddError(ErrorCode.VB0141, new string[] { "On Error Goto " + SavingText });
                                    data.AfterLabel = true;
                                }
                                // 식별자 이후이면서, As 뒤이고, (Function이거나 변수 선언중) 이라는 걸 동시에 만족 시키면
                                else if (data.AfterIdentifier && data.AfterAs && !data.AfterType)
                                {
                                    // 타입으로 인식

                                    data.AfterType = true;
                                }
                                else if (data.AfterDeclare && (data.AfterSub || data.AfterFunction) && data.AfterIdentifier && data.AfterLib && !data.AfterAlias)
                                {
                                    // TODO : API Expression (String으로 인식해서 하는거) 잘 하기
                                    continue;
                                }
                                // 배열 선언 이후라면
                                else if (data.AfterArray)
                                {
                                    if (SavingText.ToString().IsReservedKeyWords() && !(data.AfterSub || data.AfterFunction || data.AfterProperty))
                                    {
                                        AddError(ErrorCode.VB0046);
                                    }
                                }

                                else
                                {
                                    string exp = string.Empty;
                                    // If를 처음 사용했거나 ElseIf 이후 절인 경우
                                    if ((data.AfterIf || data.AfterElseIf) && WordRecognition == 1)
                                    {
                                        i -= SavingText.Length;
                                        exp = ExpressionRange(i, RecognitionTypes.BeforeThen);
                                        IsLastChar = i >= text.Length - 1;
                                        SavingText.Clear();
                                        
                                    }
                                    
                                    else if ((data.AfterCase && (!data.AfterEnd || !data.AfterElse)) ||
                                             ((data.AfterDo) && data.AfterWhile || data.AfterUntil) ||
                                             data.AfterWhile ||
                                             data.AfterSet)
                                    {
                                        i -= SavingText.Length;
                                        exp = ExpressionRange(i, RecognitionTypes.ToEndLine);
                                        IsLastChar = i >= text.Length - 1;
                                        SavingText.Clear();
                                    }
                                    else if (WordRecognition == 0)
                                    {
                                        i = 0;
                                        exp = ExpressionRange(i, RecognitionTypes.ToEndLine);
                                        IsLastChar = i >= text.Length - 1;
                                        SavingText.Clear();
                                    }

                                    if (!string.IsNullOrEmpty(exp))
                                    {
                                        var Expression = ToExpression(exp);
                                        goto ExitIf;
                                    }
                                }
                                break;
                        }
                        WordRecognition++;

                        SavingText.Clear();
                    }
                    else if (ch.IsLetterOrDigit())
                    {
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) goto ExitIf;

                        SavingText.Append(ch);
                    }
                    else if (ch.IsOperator())
                    {
                        // DEBUG : 위쪽에서 처리 되었을 가능성 존재 확인후 삭제
                        
                    }
                }
                ExitIf:
                // 마지막 문자일시
                if (IsLastChar || MultiLineRead)
                {
                    if (bracketCount != 0) AddError(ErrorCode.VB0052, new string[] { bracketCount.ToString() });

                    // 미완성된 구문 체크

                    #region [  On Error Goto ~~ / On Error Resume Next  ]

                    if (data.AfterOn && WordRecognition == 1)
                        AddError(ErrorCode.VB0186);

                    if (data.AfterOn && data.AfterError && WordRecognition == 2)
                        AddError(ErrorCode.VB0187);

                    if (data.AfterOn && data.AfterError && data.AfterGoto && WordRecognition == 3)
                        AddError(ErrorCode.VB0188);

                    if (data.AfterOn && data.AfterError && data.AfterResume && WordRecognition == 3)
                        AddError(ErrorCode.VB0189);

                    #endregion

                    #region [  엑세서 이후 식별자  ]

                    if (data.AfterAccessor && !data.AfterDeclarator && !data.AfterIdentifier )
                    {
                        AddError(ErrorCode.VB0123);
                    }

                    if (data.AfterDeclarator && !data.AfterIdentifier && !data.IsVarDeclaring && (!data.AfterEnd && !data.AfterExit))
                    {
                        if (data.AfterDeclare && !(data.AfterSub || data.AfterFunction))
                        {
                            AddError(ErrorCode.VB0209);
                        }
                        else AddError(ErrorCode.VB0121);
                    }

                    #endregion

                    #region [  ReDim  ]

                    if (data.AfterReDim && !data.AfterIdentifier && WordRecognition == 0)
                        AddError(ErrorCode.VB0154);

                    if (data.AfterReDim && data.AfterIdentifier && !data.IsInBracket)
                        AddError(ErrorCode.VB0155);

                    #endregion

                    #region [  API 미완성  ]

                    if (data.AfterDeclare && (data.AfterSub || data.AfterFunction) && data.AfterIdentifier && !data.AfterLib)
                        AddError(ErrorCode.VB0202);

                    if (data.AfterDeclare && (data.AfterSub || data.AfterFunction) && data.AfterLib && !data.AfterString && !data.AfterAlias)
                        AddError(ErrorCode.VB0203);

                    if (data.AfterDeclare && (data.AfterSub || data.AfterFunction) && data.AfterLib && data.AfterString && !data.AfterAlias)
                        AddError(ErrorCode.VB0204);

                    if (data.AfterDeclare && (data.AfterSub || data.AfterFunction) && data.AfterLib && data.AfterAlias && !data.AfterString)
                        AddError(ErrorCode.VB0205);

                    #endregion

                    // 식별자, As까지 나왔는데 Type이 안 나온 경우
                    if (data.AfterIdentifier && data.AfterAs && !data.AfterType)
                        AddError(ErrorCode.VB0124);

                        // While 이후
                    if (data.AfterWhile || 
                        // If 이후
                        (data.AfterIf && !data.AfterEnd) || 
                        // ElseIf 이후
                        data.AfterElseIf || 
                        // Select이후가 아니지만 Case이후일 경우
                        (!data.AfterSelect && data.AfterCase) || 
                        // Select Case 이후 일경우
                        (data.AfterSelect && data.AfterCase) ||
                        (data.AfterDo && (data.AfterWhile || data.AfterUntil)))
                        // Operator나 Object가 하나도 안 나오면
                        if (!data.AfterExpression) AddError(ErrorCode.VB0122);



                    if (data.AfterSelect && (!data.AfterCase && !data.AfterEnd))
                        AddError(ErrorCode.VB0076);
                    if (data.AfterDo && (!data.AfterUntil && !data.AfterWhile) && WordRecognition != 1)
                        AddError(ErrorCode.VB0083);
                    if (data.AfterIf && data.AfterExpression && !data.AfterThen)
                        AddError(ErrorCode.VB0077);

                }

                // 만약 첫번째줄이 유지되고 있다면 빈칸, 탭, 새로 띄우기에서 False가 되진 않음
                data.IsFistNonWs &= ch == ' ' || ch == '\t' || ch == '\n' || ch == '\r';

                if (MultiLineRead)
                {
                    nHandledLine = new NotHandledLine(FileName, CodeLine.Substring(i + 1), Lines);
                    // 마지막으로 인식 시켜주기
                    i = CodeLine.Length - 1;
                }
            }

            LineInfo = TempLineInfo.Clone();

            return nHandledLine;
            //======================================================================================================================



            // true : Handled   false : Not Handled
            bool ErrorCheck(string Keyword, (int,int) ErrOffset)
            {
                bool Handled = false;
                if (data.IsInString || data.IsInComment || data.IsInPreprocessorDirective) return false;

                if (data.AfterOnErrorResumeNext)
                {
                    AddError(ErrorCode.VB0184);
                    Handled = true;
                }
                if (data.AfterOnErrorGotoLabel)
                {
                    AddError(ErrorCode.VB0185);
                    Handled = true;
                }
                // Return뒤에 식을 작성했을 경우 오류 안내
                if (data.AfterReturn)
                {
                    AddError(ErrorCode.VB0012);
                    Handled = true;
                }

                // class, module 오류
                if (Keyword.ContainsWords(new string[] { "class"}))
                {
                    if (Keyword == "class") AddError(ErrorCode.VB0001);
                    
                    Handled = true;
                }
                // public, private, dim 오류
                if (Keyword.ContainsWords(new string[] { "public", "private", "dim" }))
                {
                    if (data.AfterAccessor)
                    {
                        AddError(ErrorCode.VB0022);
                        Handled = true;
                    }
                    // 현재 키워드가 Dim이면서 이미 선언자 이후라면 중복
                    if (Keyword == "dim" && data.AfterDeclarator)
                    {
                        AddError(ErrorCode.VB0021);
                        Handled = true;
                    }
                    // 현재 키워드가 Dim이면서 단어 인식 횟수가 처음이 아니라면 오류
                    if (Keyword == "dim" && WordRecognition != 0)
                    {
                        AddError(ErrorCode.VB0040, new string[] { "Dim" });
                        Handled = true;
                    }
                }
                // Enum, Property, Function, Sub 오류
                if (Keyword.ContainsWords(new string[] { "enum", "property", "function", "sub" }))
                {
                    // API 선언 부분이 아닌데 중복되었을 경우
                    if (data.AfterDeclarator && !data.AfterDeclare)
                    {
                        AddError(ErrorCode.VB0021);
                        Handled = true;
                    }
                    else if (!(data.AfterAccessor || WordRecognition == 0) && !(data.AfterEnd && WordRecognition == 1))
                    {
                        // Enum이 아닐때 Exit sub, Exit Property등등을 체크
                        if (!(Keyword != "enum" && data.AfterExit && WordRecognition == 1))
                        {
                            AddError(ErrorCode.VB0120);
                            Handled = true;
                        }
                        
                    }
                    else
                    {
                        if (Keyword == "function") data.AfterFunction = true;
                        if (Keyword == "sub") data.AfterSub = true;
                        if (Keyword == "property") data.AfterProperty = true;
                        if (Keyword == "enum") data.AfterEnum = true;

                        data.AfterDeclarator = true;
                    }
                    
                }
                if (Keyword.ContainsWords(new string[] { "get", "set", "let" }))
                {
                    if (!(data.AfterProperty && !data.AfterIdentifier) && Keyword != "set")
                    {
                        AddError(ErrorCode.VB0130, new string[] { Keyword });
                        Handled = true;
                    }

                    if (WordRecognition == 0 && Keyword != "set") {
                        AddError(ErrorCode.VB0003);
                        Handled = true;
                    }
                }

                if (Keyword == "as")
                {
                    if (!data.AfterIdentifier)
                    {
                        AddError(ErrorCode.VB0151);
                        Handled = true;
                    }
                }

                if (!Keyword.ContainsWords(new string[] { "do", "for", "sub", "function", "property" }) && data.AfterExit)
                {
                    AddError(ErrorCode.VB0011, new string[] { Keyword });
                    Handled = true;
                }

                // If, ElseIf, Else에 대한 오류
                if (Keyword.ContainsWords(new string[] {"if","elseif","else" }))
                {
                    if (Keyword == "else" && (data.AfterElse || data.AfterElseIf))
                    {
                        AddError(ErrorCode.VB0071);
                        Handled = true;
                    }
                    if (Keyword == "elseif" && data.AfterIf)
                    {
                        AddError(ErrorCode.VB0072);
                        Handled = true;
                    }
                    if (Keyword == "if" && data.AfterIf)
                    {
                        AddError(ErrorCode.VB0073);
                        Handled = true;
                    }
                    if (Keyword == "if" && WordRecognition != 0 && !data.AfterEnd)
                    {
                        AddError(ErrorCode.VB0074);
                        Handled = true;
                    }
                    if (Keyword == "if" && data.AfterElse)
                    {
                        AddError(ErrorCode.VB0075);
                        Handled = true;
                    }
                }
                // ReadOnly, AddHandler 키워드 오류
                if (Keyword.ContainsWords(new string[] { "readonly", "addhandler" }))
                {
                    if (Keyword == "readonly") AddError(ErrorCode.VB0006);
                    if (Keyword == "addhandler") AddError(ErrorCode.VB0007);
                    Handled = true;
                }

                if (ReserveWords.Contains(Keyword) && data.AfterOn && data.AfterError && data.AfterGoto && WordRecognition == 3) 
                {
                    AddError(ErrorCode.VB0190);
                    Handled = true;
                }

                // 기본 처리
                if (!Handled)
                {
                    if ((data.AfterEnd && !Keyword.ContainsWords(new string[] { "if", "select", "sub", "function", "property", "type", "with", "enum" })))
                    {
                        AddError(ErrorCode.VB0055, new string[] { Keyword });
                        Handled = true;
                    }
                    if (data.AfterProperty && Keyword != "property" && !data.AfterPropAccessor && !Keyword.ContainsWords(new string[] { "get", "set", "let" }))
                    {
                        AddError(ErrorCode.VB0131);
                        Handled = true;
                    }
                    
                    if (data.IsVarDeclaring && WordRecognition == 1 &&
                        Properties.Resources.예약어.ToLower().Split(new string[] { "\r\n" }, StringSplitOptions.None).Contains(Keyword.ToLower()))
                    {
                        AddError(ErrorCode.VB0045, new string[] { Keyword });
                        Handled = true;
                    }
                }

                if (Handled) WordRecognition++;
                
                return Handled;
            }
            
            void ParameterSeek(int TextOffset)
            { 
                int BracketInt = 0;
                
                for (int j = TextOffset; j< text.Length; j++)
                {
                    if (text[j] == '(') BracketInt++;
                    if (text[j] == ')') BracketInt--;


                    if (BracketInt == 0 || j == text.Length - 1)
                    {
                        string ParamString = "";
                        try
                        {
                            if (BracketInt == 0) ParamString = CodeLine.Substring(TextOffset + 1, j - TextOffset - 1);
                            else ParamString = CodeLine.Substring(TextOffset + 1, j - TextOffset);
                        }
                        catch (Exception)
                        {
                            return;
                        }
                        char[] cArr = ParamString.ToCharArray();

                        i += ParamString.Length;

                        string savText = string.Empty;


                        (int, int) ParamOffset = (0, 0), ValueOffset = (0, 0);

                        bool AfterBracket = false;
                        bool AfterIdentifier = false;
                        bool AfterAs = false, AfterArray = false;



                        bool AfterParamArray = false, AfterOptional = false;
                        bool AfterByVal = false, AfterByRef = false;
                        bool NeedRest = false, NeedExpression = false;

                        bool AfterString = false, AfterInt = false;

                        bool IsInString = false;

                        string name = "";
                        string type = "";

                        Expression initValue = Expression.Empty;
                        bool FirstSubstitiution = true;
                        bool UseParamArray = false, UseOptional = false, UseRest = false;
                        bool LastChar = false;
                        for(int k = 0; k< cArr.Length; k++)
                        {
                            char i_ch = cArr[k];
                            char i_nextCh = k + 1 < cArr.Length ? cArr[k + 1] : '\0';
                            LastChar = k == cArr.Length - 1 ;
                            switch (i_ch)
                            {
                                case '(':
                                    if (IsInString) continue;
                                    if (AfterBracket) { AddError(ErrorCode.VB0101); break; }
                                    if (AfterIdentifier && AfterOptional) { AddError(ErrorCode.VB0100); break; }
                                    if (AfterIdentifier && !AfterOptional) { AfterArray = true; break; }
                                    AfterBracket = true;
                                    break;
                                case '\"':
                                    
                                    if (IsInString && i_nextCh == '\"') { i_ch++; break; }

                                    if (!NeedExpression && !AfterString) { AddError(ErrorCode.VB0104); break; }

                                    IsInString = !IsInString;

                                    if (!IsInString)
                                    {
                                        NeedExpression = false;
                                        AfterString = true;
                                        AfterOptional = false;
                                    }
                                    break;
                                case ')':
                                    if (IsInString) continue;
                                    if (!AfterBracket || !AfterArray)
                                    {
                                        AddError(ErrorCode.VB0052, new string[] { "1" });
                                        break;
                                    }
                                    if (AfterArray) AfterArray = false;
                                    else AfterBracket = false; 

                                    break;
                                case '&':
                                case '+':
                                    if (NeedExpression) AddError(ErrorCode.VB0122);
                                    else if (!AfterString && !AfterInt) AddError(ErrorCode.VB0044);
                                    else NeedExpression = true;
                                    break;
                                case ',':
                                    if (IsInString) continue;
                                    if (!NeedRest)
                                    {
                                        AddError(ErrorCode.VB0098);
                                        break;
                                    }
                                    NeedRest = false;
                                    UseRest = true;
                                    break;
                                case '=':
                                    if (IsInString) continue;
                                    
                                    if (!AfterOptional)
                                    {
                                        AddError(ErrorCode.VB0103);
                                        break;
                                    }
                                    if (FirstSubstitiution)
                                    {
                                        FirstSubstitiution = false;
                                        ValueOffset.Item1 = k;
                                    }
                                    break;
                                default:
                                    if (IsInString) continue;
                                    if (i_ch == ' ') continue;
                                    savText += i_ch;
                                    break;
                            }
                            if (i_nextCh.IsDivision() || i_nextCh == ' ' || LastChar || k == cArr.Length - 1)
                            {
                                if (IsInString) continue;
                                if (savText.Replace(' ','\0') == "") continue;

                                switch (savText.ToLower())  
                                {
                                    #region [  파라미터 인식  ]
                                    case "byval":
                                        if (NeedRest)
                                        {
                                            AddError(ErrorCode.VB0099);
                                            break;
                                        }
                                        if (AfterByRef || AfterByVal || AfterOptional || AfterParamArray)
                                        {
                                            AddError(ErrorCode.VB0097);
                                            break;
                                        }
                                        if (UseParamArray || UseOptional)
                                        {
                                            AddError(ErrorCode.VB0093);
                                            break;
                                        }
                                        ParamOffset.Item1 = k;
                                        AfterByVal = true;
                                        break;
                                    case "byref":
                                        if (NeedRest)
                                        {
                                            AddError(ErrorCode.VB0099);
                                            break;
                                        }
                                        if (AfterByRef || AfterByVal || AfterOptional || AfterParamArray)
                                        {
                                            AddError(ErrorCode.VB0097);
                                            break;
                                        }
                                        if (UseParamArray || UseOptional)
                                        {
                                            AddError(ErrorCode.VB0093);
                                            break;
                                        }
                                        ParamOffset.Item1 = k;
                                        AfterByRef = true;
                                        break;
                                    case "optional":
                                        if (NeedRest)
                                        {
                                            AddError(ErrorCode.VB0099);
                                            break;
                                        }
                                        if (AfterByRef || AfterByVal || AfterOptional || AfterParamArray)
                                        {
                                            AddError(ErrorCode.VB0097);
                                            break;
                                        }
                                        if (UseParamArray)
                                        {
                                            AddError(ErrorCode.VB0091);
                                        }
                                        ParamOffset.Item1 = k;
                                        AfterOptional = true;
                                        UseOptional = true;
                                        break;
                                    case "paramarray":

                                        ParamOffset.Item1 = k;

                                        if (NeedRest) { AddError(ErrorCode.VB0099); break; }
                                        if (AfterByRef || AfterByVal || AfterOptional || AfterParamArray) { AddError(ErrorCode.VB0097); break; }
                                        if (UseOptional) { AddError(ErrorCode.VB0090); break; }
                                        if (UseParamArray) { AddError(ErrorCode.VB0092); break; }
                                        
                                        AfterParamArray = true;
                                        UseParamArray = true;
                                        break;
                                    #endregion

                                    case "as":
                                        if (NeedRest)
                                        {
                                            AddError(ErrorCode.VB0099);
                                            break;
                                        }
                                        if (AfterIdentifier) AfterAs = true;
                                        else AddError(ErrorCode.VB0151);
                                        break;
                                    default:

                                        if (savText.IsDigit() || NeedExpression)
                                        {   
                                            NeedExpression = false;
                                            AfterOptional = false;
                                            break;
                                        }

                                        if (NeedRest)
                                        {
                                            AddError(ErrorCode.VB0099);
                                            break;
                                        }
                                        
                                        if (AfterIdentifier && AfterAs)
                                        {
                                            type = savText;



                                            if (AfterByVal)
                                            {
                                                
                                                NeedRest = true;
                                            }
                                            else if (AfterByRef)
                                            {
                                                
                                                NeedRest = true;
                                            }
                                            else if (AfterOptional)
                                            {
                                                

                                                UseOptional = true;

                                                AfterParamArray = false; AfterIdentifier = false; AfterAs = false; AfterByRef = false; AfterByVal = false;

                                                AfterOptional = true;
                                                NeedExpression = true;
                                                break;
                                            }
                                            else if (AfterParamArray)
                                            {
                                                
                                                NeedRest = true;   
                                            }
                                            else
                                            {
                                                
                                                NeedRest = true;
                                            }

                                            AfterParamArray = false; AfterIdentifier = false; AfterAs = false; AfterByRef = false; AfterByVal = false;

                                            AfterOptional = false;

                                        }
                                        else if (AfterIdentifier && !AfterAs)
                                        {
                                            AddError(ErrorCode.VB0152);
                                        }
                                        else
                                        {
                                            // Identifier로 인식
                                            
                                            if (Regex.IsMatch(savText, NamePattern))
                                            {
                                                AfterIdentifier = true;
                                                name = savText;
                                            }
                                            else
                                            {
                                                AddError(ErrorCode.VB0043);
                                                AfterIdentifier = true;
                                                name = "";
                                            }
                                        }
                                        break;
                                }
                                
                                savText = string.Empty;
                            }
                        }

                        if (!AfterIdentifier && UseRest && !NeedRest)
                        {
                            if (AfterOptional && NeedExpression) AddError(ErrorCode.VB0095);
                            else if (AfterOptional) AddError(ErrorCode.VB0097);
                        }

                        if (IsInString) AddError(ErrorCode.VB0160);


                        break;
                    }
                }
                if (BracketInt != 0) AddError(ErrorCode.VB0052, new string[] { BracketInt.ToString() });

            }
            
            // 식을 인식하는 메소드
            string ExpressionRange(int TextOffset, RecognitionTypes RecognitionType)
            {
                string str = CodeLine.Substring(TextOffset);
                char[] cArr = str.ToCharArray();

                string ReturnStr = string.Empty;

                string saveStr = string.Empty;

                bool IsInString = false;

                int bracketctr = 0;
                int j;
                bool LastChar = false;
                for (j = 0; j < cArr.Length; j++)
                {
                    if (j == cArr.Length - 1) LastChar = true;
                    else LastChar = false;
                    char e_ch = cArr[j];
                    char e_nextCh = j + 1 < cArr.Length ? cArr[j + 1] : '\0';
                    if (e_ch == '\'')
                    {
                        j--;
                        goto returnPoint;
                    }
                    if (e_ch == '\"') IsInString = !IsInString;
                    if (e_ch == ':' && !IsInString)
                    {
                        goto returnPoint;
                    }
                    if (e_ch == '_' && !IsInString)
                        continue;
                    if (IsInString)
                    {
                        // String 내부면 반환 텍스트에 추가
                        ReturnStr += e_ch;
                        continue;
                    }

                    switch (e_ch)
                    {
                        case '(':
                            bracketctr++;

                            if ((RecognitionType == RecognitionTypes.BracketExpression || RecognitionType == RecognitionTypes.Parameter) &&
                                bracketctr == 1) continue;

                            saveStr += e_ch;
                            break;
                        case ')':
                            bracketctr--;
                            saveStr += e_ch;
                            if (bracketctr == 0 && (RecognitionType == RecognitionTypes.BracketExpression || RecognitionType == RecognitionTypes.Parameter))
                            {
                                
                                return ReturnStr.Trim();
                            }
                            break;
                        case ',':
                            if (bracketctr == 0 && RecognitionType == RecognitionTypes.Parameter) return ReturnStr;

                            break;
                        case '.':
                        case ' ':
                            // 빈칸 중복 방지
                            if (e_nextCh != ' ') ReturnStr += e_ch;
                            break;
                        case '\r':
                        case '\n':
                            break;
                        default:
                            saveStr += e_ch;
                            break;
                            
                    }

                    if (bracketctr == 0 && RecognitionType == RecognitionTypes.BracketExpression)  goto returnPoint;


                    if (e_nextCh.IsDivision() || e_nextCh == ' ' || e_nextCh == ':' || LastChar)
                    {
                        bool Handled = true;
                        switch (saveStr.ToLower())
                        {
                            case "then":
                                if (RecognitionType == RecognitionTypes.BeforeThen)
                                {
                                    j -= 4;
                                    goto returnPoint;
                                }
                                else AddError(ErrorCode.VB0048, new string[] { saveStr });
                                break;
                            default:
                                Handled = false;
                                break;
                        }

                        if (!Handled)
                        {
                            if (ExpressionReserveWords.Contains(saveStr.ToLower()))
                            {
                                AddError(ErrorCode.VB0048,new string[] { saveStr });
                            }
                            ReturnStr += saveStr;
                        }

                        saveStr = "";
                    }

                }
                returnPoint:

                i += j - 1;
                if (!string.IsNullOrEmpty(ReturnStr.Trim())) data.AfterExpression = true;
                return ReturnStr.Trim();
            }

            // 파라미터 사용 인자 인식
            // ()가 하나 있을때의 ,를 기준으로 파라미터 인식 (예를 들면. 식 안의 파라미터는 인자로 인식 안함)
            // 단, 첫번째 문자가 '('가 아닐경우 정상적으로 인식 되지 않음 (바로 나가버림)
            string[] ParameterUseExpression(int TextOffset)
            {
                List<string> Parameters = new List<string>();
                string str = CodeLine.Substring(TextOffset);
                char[] cArr = str.ToCharArray();

                int innerBracketCtr = 0;

                int startPos = 1;
                int j;
                for (j = 0; j < str.Length; j++)
                {
                    char e_ch = cArr[j];
                    char e_nextCh = j + 1 < cArr.Length ? cArr[j + 1] : '\0';


                    if (e_ch == '(') innerBracketCtr++;
                    if (e_ch == ')') innerBracketCtr--;
                    
                    switch (e_ch)
                    {
                        case ',':
                            // 메인 괄호 안에서만 작동
                            if (innerBracketCtr == 1)
                            {
                                string Msg = str.Substring(startPos, j - startPos);
                                Parameters.Add(Msg);

                                //MessageBox.Show(Msg);

                                startPos = j + 1;
                            }
                            break;
                        case ')':
                            if (innerBracketCtr == 0)
                            {
                                string Msg = str.Substring(startPos, j - startPos);
                                Parameters.Add(Msg);

                                //MessageBox.Show(Msg);

                                startPos = j + 1;
                            }
                            break;
                    }

                    if (innerBracketCtr == 0)
                    {
                        break;
                    }
                }
                i += j;

                

                return Parameters.ToArray();
            }

            // 식 계산
            Expression ToExpression(string Expression)
            {
                return new Expression(Expression, CodeInfo.ErrorList, i, FileName);
            }
            
            void AddError(ErrorCode Code, string[] parameters = null, int Line = -1)
            {
                if (Line == -1) CodeInfo.ErrorList.Add(new Error(ErrorType.Error, Code, parameters, FileName, new DomRegion(Lines.Item1.StartInt, 0)));
                else CodeInfo.ErrorList.Add(new Error(ErrorType.Error, Code, parameters, FileName, new DomRegion(Line, 0)));
            }

            void AddWarning(ErrorCode Code, string[] parameters = null, int Line = -1)
            {
                if (Line == -1) CodeInfo.ErrorList.Add(new Error(ErrorType.Warning, Code, parameters, FileName, new DomRegion(Lines.Item1.StartInt, 0)));
                else CodeInfo.ErrorList.Add(new Error(ErrorType.Warning, Code, parameters, FileName, new DomRegion(Line, 0)));
            }
        }
        
    }

}
