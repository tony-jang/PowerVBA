using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
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
        enum RecognitionCases
        {
            AfterPreprocessor,
            AfterDim,
            AfterPublic,
            AfterPrivate,
            AfterSub,
            AfterFunction,
        }

        public VBASeeker(string codeLine, RangeInt line, List<Error> Errors, LineInfo info)
        {
            CodeLine = codeLine;
            Lines = line;
            this.Errors = Errors;
            Info = info;
        }

        private string CodeLine;
        private RangeInt Lines;
        private List<Error> Errors;
        private LineInfo Info;

        // TODO : LineInfo 추가한 것으로 처리

        public List<CodeData> GetLine()
        {
            
            void AddErrMsg(string Message, int Line = -1)
            {
                if (Line == -1) Errors.Add(new Error(ErrorType.Error, Message, new DomRegion(Lines.StartInt, 0)));
                else Errors.Add(new Error(ErrorType.Error, Message, new DomRegion(Line, 0)));
            }
            void AddErrCode(ErrorCode Code, string[] parameters = null, int Line = -1)
            {
                if (Line == -1) Errors.Add(new Error(ErrorType.Error, Code, parameters, new DomRegion(Lines.StartInt, 0)));
                else Errors.Add(new Error(ErrorType.Error, Code, parameters, new DomRegion(Line, 0)));
            }

            List<CodeData> codeData = new List<CodeData>();
            //  for 카운터    여는 괄호 갯수    단어 인식 처리 횟수
            int i = 0, bracketCount = 0, WordRecognition = 0;



            var data = new CodeData();
            data.IsFistNonWs = true;
            data.CodeSegment = new TextSegment() { StartOffset = 0, EndOffset = 0, Length = 0 };

            data.PropertyChanged += delegate (object sender, PropertyChangedEventArgs e)
            {
                var d = data.Clone();
                ((TextSegment)d.CodeSegment).EndOffset = i;

                codeData.Add(d);

                ((TextSegment)data.CodeSegment).StartOffset = i + 1;
            };

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
                    case '#': // 전처리기 지시문
                        // 첫번째 문자가 #일 경우 전처리기 지시문으로 인식함
                        if (data.IsFistNonWs)
                            data.IsInPreprocessorDirective = true;
                        break;
                    case '\'': // 주석
                        // String 중이거나, Comment내부라면
                        if (data.IsInString || data.IsInVerbatimString || data.IsInComment)
                            break;

                        data.IsInComment = true;
                        data.IsInPreprocessorDirective = false;
                        break;
                    case '\n':
                    case '\r': // 새 줄 인식될때 초기화
                        data.IsInComment = false;
                        data.IsInString = false;
                        data.IsInPreprocessorDirective = false;
                        data.IsFistNonWs = true;
                        break;
                    case '"': // string 문자열 시작
                        if (data.IsInComment) break;
                        if (data.IsInVerbatimString)
                        {
                            if (nextCh == '"')
                            {
                                i++;
                                break;
                            }
                            data.IsInVerbatimString = false;
                            break;
                        }
                        data.IsInString = !data.IsInString;

                        if (data.IsInString) data.Type = CodeType.String;

                        break;
                    case '(': // 여는 괄호
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment)
                            break;

                        if (data.AfterDeclarator) AddErrCode(ErrorCode.VB0050);

                        if (data.AfterIdentifier)
                        {
                            // Array
                            if (data.IsVar)
                            {
                                data.AfterArray = true;
                                AddErrCode(ErrorCode.VB0000, new string[] { "Array가 인식되었습니다." });
                            }
                            // 파라미터
                            else
                            {
                                
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
                            AddErrCode(ErrorCode.VB0051);
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
                    case '=':
                        if (data.AfterDeclarator || data.AfterIdentifier || data.IsVar)
                        {
                            AddErrCode(ErrorCode.VB0008);
                        }
                        break;
                    case ';': case '!': case '@': case '$': case '^': case '%': case '\\': case '~':
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) break;

                        AddErrCode(ErrorCode.VB0041, new string[] { ch.ToString() });
                        break;

                    //이외 체크
                    default:
                        Handled = false;
                        break;
                }

                // switch문에서 처리되지 않았거나 처리 되었지만 읽을 필요가 있는 경우
                if (!Handled)
                {
                    // 빈칸이거나 마지막 단어 일경우 또는 읽을 필요가 있는 경우
                    // 또는 다음 글자가 '('나 ')' 같이 읽을 필요가 있는 경우
                    if (ch.IsWhiteSpace() || IsLastChar || nextCh.IsDivision())
                    {
                        // string이거나 전처리기 지시문이거나 코멘트일 경우 넘김
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) continue;

                        // 마지막 글자이면서 빈칸이 아니라면 인식할 문자열에 추가
                        if (IsLastChar && !(ch.IsWhiteSpace() || ch.IsOperator() || ch.IsBracket())) SavingText += ch;
                        
                        if (string.IsNullOrEmpty(SavingText)) continue;
                        switch (SavingText.ToLower())
                        {
                            #region [  Accessor  ]
                            case "public":
                                if (data.AfterAccessor) AddErrCode(ErrorCode.VB0022);

                                data.Type = CodeType.PublicAccessor;
                                data.AfterAccessor = true;
                                break;
                            case "private":
                                if (data.AfterAccessor) AddErrCode(ErrorCode.VB0022);

                                data.Type = CodeType.PrivateAccessor;
                                data.AfterAccessor = true;
                                break;
                            #endregion

                            #region [  Declartor  ]
                            case "dim":
                                if (data.AfterAccessor || data.AfterDeclarator) AddErrCode(ErrorCode.VB0040, new string[] { "Dim" });
                                data.AfterAccessor = false;
                                data.AfterDeclarator = true;
                                data.IsVar = true;
                                data.Type = CodeType.Dim;
                                break;
                            case "sub":

                                if (data.AfterDeclarator) AddErrCode(ErrorCode.VB0021);
                                if (!(data.AfterAccessor || WordRecognition == 0 || data.AfterEnd)) AddErrCode(ErrorCode.VB0100);

                                data.AfterDeclarator = true;

                                if (data.AfterEnd) data.Type = CodeType.EndSub;
                                else data.Type = CodeType.DeclareSub;

                                break;
                            case "function":
                                if (data.AfterDeclarator) AddErrCode(ErrorCode.VB0021);
                                if (!(data.AfterAccessor || WordRecognition == 0 || data.AfterEnd)) AddErrCode(ErrorCode.VB0100);

                                if (data.AfterEnd)
                                    data.Type = CodeType.EndFunction;
                                else
                                    data.Type = CodeType.DeclareFunction;
                                data.AfterDeclarator = true;
                                break;
                            case "property":
                                if (data.AfterDeclarator) AddErrCode(ErrorCode.VB0021);
                                if (data.AfterEnd)
                                {
                                    data.Type = CodeType.EndProperty;
                                }
                                else if (!(data.AfterAccessor || WordRecognition == 0)) AddErrCode(ErrorCode.VB0100);

                                data.AfterProperty = true;
                                break;
                            case "enum":
                                if (data.AfterDeclarator) AddErrCode(ErrorCode.VB0021);

                                if (data.AfterEnd)
                                {
                                    data.Type = CodeType.EndEnum;
                                }
                                else
                                {
                                    data.Type = CodeType.DeclareEnum;
                                }
                                data.AfterDeclarator = true;
                                break;
                            case "class":
                                data.Type = CodeType.Class;
                                AddErrCode(ErrorCode.VB0001);
                                break;
                            case "module":
                                data.Type = CodeType.Module;
                                AddErrCode(ErrorCode.VB0002);

                                break;
                            #endregion

                            #region [  Conditional  ]

                            case "if":
                                if (data.AfterEnd)
                                {
                                    data.AfterIf = true;
                                    data.Type = CodeType.EndIf;
                                }
                                else if (data.AfterIf)
                                {
                                    AddErrCode(ErrorCode.VB0068);
                                }
                                else if (WordRecognition != 0)
                                {
                                    AddErrCode(ErrorCode.VB0069);
                                }
                                else
                                {
                                    data.AfterIf = true;
                                }
                                break;
                            case "elseif":
                                if (data.AfterIf) AddErrCode(ErrorCode.VB0067);
                                else
                                {
                                    data.AfterElseIf = true;
                                }

                                break;
                            case "else":
                                if (data.AfterElse || data.AfterElseIf) AddErrCode(ErrorCode.VB0066);

                                break;

                            case "select":
                                
                                break;

                            #endregion

                            #region [  Iterative  ]
                                
                            case "while":

                                if (data.AfterEnd)
                                {
                                    AddErrCode(ErrorCode.VB0005);
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
                                data.Type = CodeType.ExitDo;
                                break;
                            #endregion
                                
                            #region  [  Get/Set  ]
                            case "get":
                                if (WordRecognition == 0) AddErrCode(ErrorCode.VB0003);
                                break;
                            case "set":
                                if (WordRecognition == 0) AddErrCode(ErrorCode.VB0004);
                                break;

                            #endregion

                            case "end":
                                data.AfterEnd = true;
                                break;

                            #region [  VB.NET 미호환 문법  ]


                            case "readonly":
                                AddErrCode(ErrorCode.VB0006);
                                break;
                            case "addhandler":
                                AddErrCode(ErrorCode.VB0007);
                                break;
                            #endregion

                            default:

                                if (WordRecognition == 0) data.Type = CodeType.Unknown;

                                else
                                {
                                    if (data.AfterEnd)
                                    {
                                        AddErrCode(ErrorCode.VB0060, new string[] { SavingText });
                                    }
                                    if (data.AfterDeclarator)
                                    {
                                        if (SavingText.ToLower() == "as") AddErrCode(ErrorCode.VB0045, new string[] { "As" });

                                        data.AfterDeclarator = false;
                                        data.AfterIdentifier = true;
                                        
                                        // Public/Private/Dim 뒤에 식별자가 나왔을 경우 변수 선언으로 인식 및 처리
                                        if (data.Type == CodeType.PublicAccessor ||
                                            data.Type == CodeType.PrivateAccessor ||
                                            data.Type == CodeType.Dim) data.IsVar = true;
                                        

                                        data.Type = CodeType.Identifier;
                                    }
                                    else if (data.AfterIdentifier)
                                    {
                                        if (SavingText.ToLower() != "as") AddErrCode(ErrorCode.VB0027, new string[] { "As" });
                                        data.AfterIdentifier = false;
                                    }
                                    else if (data.AfterArray)
                                    {
                                        if (SavingText.IsReservedKeyWords()) AddErrCode(ErrorCode.VB0046);
                                    }
                                }

                                break;
                        }
                        WordRecognition++;
                        SavingText = "";

                        //MessageBox.Show("빈칸");
                    }
                    else if (ch.IsLetterOrDigit())
                    {
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) continue;

                        SavingText += ch;

                    }
                    else if (ch.IsOperator())
                    {
                        data.Type = CodeType.Operator;
                    }
                }
                // 마지막 문자일시
                if (IsLastChar)
                {
                    if (bracketCount >= 1) AddErrCode(ErrorCode.VB0052, new string[] { bracketCount.ToString() });
                }

                // 만약 첫번째줄이 유지되고 있다면 빈칸, 탭, 새로 띄우기에서 False가 되진 않음
                data.IsFistNonWs &= ch == ' ' || ch == '\t' || ch == '\n' || ch == '\r';
            }

            return codeData;
        }
    }

}
