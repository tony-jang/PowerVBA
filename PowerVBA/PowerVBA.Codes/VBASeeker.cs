using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Codes.CodeItems;
using PowerVBA.Codes.Enums;
using PowerVBA.Codes.Extension;
using PowerVBA.Codes.TypeSystem;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerVBA.Codes
{
    class VBASeeker
    {
        public VBASeeker(CodeInfo codeInfo)
        {
            CodeInfo = codeInfo;
            CurrentLineCodeList = new LineCodeItem((0, 0));
        }

        private CodeInfo CodeInfo;
        // TODO : LineInfo 추가한 것으로 처리

        private LineCodeItem CurrentLineCodeList { get; set; }

        // 가장 마지막에 추가된 아이템
        private CodeItemBase LastCodeItm { get; set; }

        public void GetLine(string CodeLine, (RangeInt,int) Lines)
        {
            //CurrentLineCodeList = new LineCodeItem(CodeLine);
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
            string SavingText = "";

            char[] text = CodeLine.ToCharArray();
            for (i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                bool IsLastChar = i == text.Length - 1;
                bool Handled = true;
                char nextCh = i + 1 < text.Length ? text[i + 1] : '\0';
                switch (ch)
                {
                    #region [  string / 주석 / 전처리기 지시문  ]
                    case '#': // 전처리기 지시문
                        // 첫번째 문자가 #일 경우 전처리기 지시문으로 인식함
                        if (data.IsFistNonWs)
                        {

                            data.IsInPreprocessorDirective = true;
                        }

                        break;
                    case '\'': // 주석
                        // String 중이거나, Comment내부라면
                        if (data.IsInString || data.IsInVerbatimString || data.IsInComment)
                            break;
                        data.IsInComment = true;

                        data.IsInPreprocessorDirective = false;
                        break;
                    case '"': // string 문자열 시작
                        if (data.IsInComment) break;

                        if (nextCh == '"')
                        {
                            i++;
                            break;
                        }

                        data.IsInString = !data.IsInString;

                        if (data.IsInString) Offset.Item1 = i;
                        else
                        {
                            Offset.Item2 = i;
                            AddItem(new TextItem(CodeLine.Substring(Offset.Item1, i - Offset.Item1), "", Offset));
                        }
                        break;
                    #endregion

                    case '\n':
                    case '\r': // 새 줄 인식될때 초기화
                        data.IsInComment = false;
                        data.IsInString = false;
                        data.IsInPreprocessorDirective = false;
                        data.IsFistNonWs = true;
                        break;

                    #region [  괄호 (여닫는 괄호)  ]
                    case '(': // 여는 괄호
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment)
                            break;

                        if (data.AfterDeclarator) AddError(ErrorCode.VB0050);

                        if (data.AfterIdentifier)
                        {
                            // Array
                            if (data.IsVarDeclaring)
                            {
                                data.AfterArray = true;
                                AddError(ErrorCode.VB0000, new string[] { "Array가 인식되었습니다." });
                            }
                            // 파라미터 (VBA에서 Public Sub A 나 Public Function A 이후에 괄호가 나오면 파라미터의 시작이다.
                            else
                            {
                                data.IsInParameters = true;
                            }
                            data.AfterIdentifier = false;
                        }

                        bracketCount++;

                        data.IsInBracket = true;
                        break;
                    case ')': // 닫는 괄호
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) break;

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

                    case '<':
                    case '>':
                    case '+':
                    case '-':
                    case '*':
                    case '!':
                    case '=':

                        // 다음 글자도 Operator일시 중복 Operator로 인식해서 처리
                        if (nextCh.IsOperator())
                        {
                            string multiOperator = ch.ToString() + nextCh.ToString();

                            switch (multiOperator)
                            {
                                case "<>":
                                    AddItem(new OperatorItem(multiOperator, (i, 2)));
                                    break;
                                case "=<":
                                case "<=":
                                    if (multiOperator.StartsWith("="))
                                        AddWarning(ErrorCode.VB0065, new string[] { multiOperator, "<=" });

                                    AddItem(new OperatorItem(multiOperator, (i, 2)));
                                    break;
                                case "=>":
                                case ">=":
                                    if (multiOperator.StartsWith("="))
                                        AddWarning(ErrorCode.VB0065, new string[] { multiOperator, ">=" });
                                    AddItem(new OperatorItem(multiOperator, (i, 2)));

                                    break;
                                case "*=":
                                case "-=":
                                case "+=":
                                    AddError(ErrorCode.VB0009, new string[] { multiOperator, $"Object = Object { multiOperator.Substring(0,1) } Value" });
                                    break;
                                case "!=":
                                    AddError(ErrorCode.VB0009, new string[] { multiOperator, $"Left <> Right" });
                                    break;
                            }
                        }
                        else
                        {

                        }
                        if (data.AfterDeclarator || data.AfterIdentifier || data.IsVarDeclaring)
                        {
                            AddError(ErrorCode.VB0008);
                            
                        }
                        break;

                    #endregion

                    #region [  특수 문자  ]
                    case ';':
                    case '@':
                    case '$':
                    case '^':
                    case '%':
                    case '\\':
                    case '~':
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) break;

                        AddError(ErrorCode.VB0041, new string[] { ch.ToString() });
                        break;

                    #endregion
                    case ':': // 멀티 라인 인식
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) break;
                        break;
                    //이외 체크
                    default:
                        Handled = false;
                        break;
                }

                // switch문에서 처리되지 않은 경우
                if (!Handled)
                {
                    if (ch.ToLower().ToString() == "a")
                        Console.WriteLine("!");
                    (int, int) CurrentOffset = (i - SavingText.Length - 1, i);

                    // 빈칸이거나 마지막 단어 일경우 또는 읽을 필요가 있는 경우
                    // 또는 다음 글자가 '('나 ')' 같이 읽을 필요가 있는 경우
                    if (ch.IsWhiteSpace() || IsLastChar || nextCh.IsDivision())
                    {
                        // string이거나 전처리기 지시문이거나 코멘트일 경우 넘김
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) continue;

                        // 마지막 글자이면서 빈칸이 아니라면 인식할 문자열에 추가
                        if ((IsLastChar || nextCh.IsDivision()) && !(ch.IsWhiteSpace() || ch.IsOperator() || ch.IsBracket())) SavingText += ch;

                        if (string.IsNullOrEmpty(SavingText)) continue;



                        switch (SavingText.ToLower())
                        {
                            #region [  Accessor  ]
                            case "public":
                                if (data.AfterAccessor) AddError(ErrorCode.VB0022);

                                AddItem(new AccessorItem(Accessor.Public, "", CurrentOffset));
                                data.AfterAccessor = true;
                                break;
                            case "private":
                                if (data.AfterAccessor) AddError(ErrorCode.VB0022);

                                AddItem(new AccessorItem(Accessor.Private, "", CurrentOffset));
                                data.AfterAccessor = true;
                                break;
                            case "dim":
                                if (data.AfterAccessor || data.AfterDeclarator) AddError(ErrorCode.VB0040, new string[] { "Dim" });
                                if (data.IsInParameters) AddError(ErrorCode.VB0094);

                                data.AfterAccessor = false;
                                data.AfterDeclarator = true;
                                data.IsVarDeclaring = true;

                                AddItem(new AccessorItem(Accessor.Dim, "", CurrentOffset));

                                break;
                            #endregion

                            #region [  Declartor  ]

                            case "sub":
                                if (data.AfterDeclarator) AddError(ErrorCode.VB0021);

                                if (data.AfterEnd) AddItem(new EndItem(ClosingItem.Sub, CurrentOffset));
                                else if (!(data.AfterAccessor || WordRecognition == 0 || data.AfterEnd)) AddError(ErrorCode.VB0100);
                                else AddItem(new DeclaratorItem(DeclaratorType.Sub, CurrentOffset));

                                data.AfterDeclarator = true;
                                break;
                            case "function":
                                if (data.AfterDeclarator) AddError(ErrorCode.VB0021);

                                if (data.AfterEnd) AddItem(new EndItem(ClosingItem.Function, CurrentOffset));
                                else if (!(data.AfterAccessor || WordRecognition == 0 || data.AfterEnd)) AddError(ErrorCode.VB0100);
                                else AddItem(new DeclaratorItem(DeclaratorType.Function, CurrentOffset));

                                data.AfterDeclarator = true;
                                break;
                            case "property":
                                if (data.AfterDeclarator) AddError(ErrorCode.VB0021);

                                if (data.AfterEnd) AddItem(new EndItem(ClosingItem.Property, CurrentOffset));
                                else if (!(data.AfterAccessor || WordRecognition == 0)) AddError(ErrorCode.VB0100);
                                else AddItem(new DeclaratorItem(DeclaratorType.Property, CurrentOffset));

                                data.AfterProperty = true;
                                break;
                            case "enum":
                                if (data.AfterDeclarator) AddError(ErrorCode.VB0021);

                                if (data.AfterEnd) AddItem(new EndItem(ClosingItem.Enum, CurrentOffset));
                                else if (!(data.AfterAccessor || WordRecognition == 0)) AddError(ErrorCode.VB0100);
                                else AddItem(new DeclaratorItem(DeclaratorType.Enum, CurrentOffset));

                                data.AfterDeclarator = true;
                                break;
                            case "class":
                                AddItem(new UnknownItem(CurrentOffset));
                                AddError(ErrorCode.VB0001);

                                break;
                            case "module":
                                AddItem(new UnknownItem(CurrentOffset));
                                AddError(ErrorCode.VB0002);

                                break;
                            #endregion

                            #region [  Conditional  ]

                            case "if":
                                if (data.AfterEnd)
                                {
                                    data.AfterIf = true;
                                    AddItem(new EndItem(ClosingItem.If, CurrentOffset));
                                }
                                else if (data.AfterIf) AddError(ErrorCode.VB0073);
                                else if (data.AfterElse) AddError(ErrorCode.VB0075);
                                else if (WordRecognition != 0) AddError(ErrorCode.VB0074);
                                else data.AfterIf = true;

                                break;
                            case "elseif":
                                if (data.AfterIf) AddError(ErrorCode.VB0072);
                                else
                                {
                                    data.AfterElseIf = true;
                                }

                                break;
                            case "else":
                                if (data.AfterElse || data.AfterElseIf) AddError(ErrorCode.VB0071);

                                data.AfterElse = true;

                                break;
                            case "select":

                                break;

                            case "case":

                                break;
                            #endregion

                            #region [  Iterative  ]

                            case "while":
                                if (data.AfterDo || data.AfterLoop)
                                {
                                    data.AfterWhile = true;
                                }
                                if (data.AfterEnd)
                                {
                                    AddError(ErrorCode.VB0005);
                                }
                                break;
                            case "until":
                                // Do Until or Loop Until
                                if (data.AfterDo || data.AfterLoop)
                                {
                                    data.AfterUntil = true;
                                }
                                break;
                            case "do":
                                if (WordRecognition == 0) data.AfterDo = true;
                                else if (data.AfterExit) data.AfterDo = true;
                                AddItem(new ExitItem(CanExitItems.Do, CurrentOffset));
                                break;

                            // 특이한 End While의 형태
                            case "wend":

                                break;

                            #endregion

                            #region  [  Get/Set  ]
                            case "get":
                                if (WordRecognition == 0) AddError(ErrorCode.VB0003);
                                break;
                            case "set":
                                if (WordRecognition == 0) AddError(ErrorCode.VB0004);
                                break;

                            #endregion

                            #region [  Parameter  ]

                            case "byval":

                                break;
                            case "byref":

                                break;
                            case "paramarray":

                                break;
                            case "optional":

                                break;

                            #endregion

                            case "end":
                                data.AfterEnd = true;
                                break;

                            #region [  VB.NET 미호환 문법  ]


                            case "readonly":
                                AddError(ErrorCode.VB0006);
                                break;
                            case "addhandler":
                                AddError(ErrorCode.VB0007);
                                break;
                            #endregion

                            default:
                                if (WordRecognition == 0) AddItem(new UnknownItem(CurrentOffset));
                                else
                                {
                                    if (data.AfterEnd)
                                    {
                                        AddError(ErrorCode.VB0060, new string[] { SavingText });
                                    }
                                    if (data.AfterDeclarator)
                                    {
                                        if (SavingText.ToLower() == "as") AddError(ErrorCode.VB0045, new string[] { "As" });

                                        data.AfterDeclarator = false;
                                        data.AfterIdentifier = true;

                                        // Public/Private/Dim 뒤에 식별자가 나왔을 경우 변수 선언으로 인식 및 처리
                                        if (LastCodeItm.GetType() == typeof(AccessorItem)) data.IsVarDeclaring = true;

                                        AddItem(new IdentifierItem(SavingText, CurrentOffset));
                                    }
                                    else if (data.AfterIdentifier)
                                    {
                                        if (SavingText.ToLower() != "as") AddError(ErrorCode.VB0027, new string[] { "As" });
                                        data.AfterIdentifier = false;
                                    }
                                    else if (data.AfterArray)
                                    {
                                        if (SavingText.IsReservedKeyWords()) AddError(ErrorCode.VB0046);
                                    }
                                }

                                break;
                        }
                        WordRecognition++;

                        SavingText = string.Empty;

                        //MessageBox.Show("빈칸");
                    }
                    else if (ch.IsLetterOrDigit())
                    {
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) continue;

                        SavingText += ch;

                    }
                    else if (ch.IsOperator())
                    {
                        // DEBUG : 위쪽에서 처리 되었을 가능성 존재 확인후 삭제
                        AddItem(new OperatorItem(SavingText, CurrentOffset));
                    }
                }
                // 마지막 문자일시
                if (IsLastChar)
                {
                    if (bracketCount >= 1) AddError(ErrorCode.VB0052, new string[] { bracketCount.ToString() });

                    // 전처리기 지시문, 코멘트 등 아이템으로 추가시키기

                    if (data.IsInComment)
                    {
                        AddItem(new CommentItem("", Offset));
                    }
                    else if (data.IsInPreprocessorDirective)
                    {

                    }

                }

                // 만약 첫번째줄이 유지되고 있다면 빈칸, 탭, 새로 띄우기에서 False가 되진 않음
                data.IsFistNonWs &= ch == ' ' || ch == '\t' || ch == '\n' || ch == '\r';
            }

            CodeInfo.Childrens.Add(CurrentLineCodeList);

            void AddError(ErrorCode Code, string[] parameters = null, int Line = -1)
            {
                if (Line == -1) CodeInfo.ErrorList.Add(new Error(ErrorType.Error, Code, parameters, new DomRegion(Lines.Item1.StartInt, 0)));
                else CodeInfo.ErrorList.Add(new Error(ErrorType.Error, Code, parameters, new DomRegion(Line, 0)));
            }

            void AddWarning(ErrorCode Code, string[] parameters = null, int Line = -1)
            {
                if (Line == -1) CodeInfo.ErrorList.Add(new Error(ErrorType.Warning, Code, parameters, new DomRegion(Lines.Item1.StartInt, 0)));
                else CodeInfo.ErrorList.Add(new Error(ErrorType.Warning, Code, parameters, new DomRegion(Line, 0)));
            }

            void AddItem(CodeItemBase codeItm)
            {
                CurrentLineCodeList.Childrens.Add(codeItm);
                LastCodeItm = codeItm;
            }
        }
    }

}
