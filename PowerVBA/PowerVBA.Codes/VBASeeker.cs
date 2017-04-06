using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Codes.CodeItems;
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
            CurrentLineCodeList = new LineCodeItem((0, 0));
        }


        string NamePattern = @"^([_a-zA-z가-힣ㅏ-ㅣㄱ-ㅎ][_a-zA-Z가-힣ㅏ-ㅣㄱ-ㅎ1-9]*)$";


        private CodeInfo CodeInfo;
        // TODO : LineInfo 추가한 것으로 처리

        private LineCodeItem CurrentLineCodeList { get; set; }

        // 가장 마지막에 추가된 아이템
        private CodeItemBase LastCodeItm { get; set; }

        public void GetLine(string FileName, string CodeLine, (RangeInt,int) Lines, bool Nested = false)
        {
            CurrentLineCodeList = new LineCodeItem((Lines.Item1.StartInt, Lines.Item1.EndInt));
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
                bool MultiLineRead = false;
                char nextCh = i + 1 < text.Length ? text[i + 1] : '\0';
                // 기본적인 문자
                switch (ch)
                {
                    #region [  string / 주석 / 전처리기 지시문  ]
                    case '#': // 전처리기 지시문
                        if (data.IsInString || data.IsInVerbatimString || data.IsInComment)
                            break;
                        // 첫번째 문자가 #일 경우 전처리기 지시문으로 인식함
                        if (data.IsFistNonWs) data.IsInPreprocessorDirective = true;
                        else AddError(ErrorCode.VB0042);
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
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment || SavingText != string.Empty)
                            break;
                        data.UseMultiLine = true;

                        break;

                    #region [  괄호 (여닫는 괄호)  ]
                    case '(': // 여는 괄호
                        
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment)
                            break;
                        
                        if (data.AfterDeclarator) AddError(ErrorCode.VB0050);

                        if (data.AfterIdentifier)
                        {
                            AddItem(new TokenItem("(", (i, 1)));
                            // Array
                            if (data.IsVarDeclaring)
                            {
                                data.AfterArray = true;
                            }
                            // 파라미터 (VBA에서 Public Sub A 나 Public Function A 이후에 괄호가 나오면 파라미터의 시작이다.
                            else
                            {
                                ParameterSeek(i);
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

                        // 다음 글자도 Operator일시 중복 Operator로 인식해서 처리
                        if (nextCh.IsOperator())
                        {
                            string multiOperator = ch.ToString() + nextCh.ToString();

                            switch (multiOperator)
                            {
                                case "<>":
                                    AddItem(new TokenItem(multiOperator, (i, 2)));
                                    break;
                                case "=<":
                                case "<=":
                                    if (multiOperator.StartsWith("="))
                                        AddWarning(ErrorCode.VB0060, new string[] { multiOperator, "<=" });

                                    AddItem(new TokenItem(multiOperator, (i, 2)));
                                    break;
                                case "=>":
                                case ">=":
                                    if (multiOperator.StartsWith("="))
                                        AddWarning(ErrorCode.VB0060, new string[] { multiOperator, ">=" });
                                    AddItem(new TokenItem(multiOperator, (i, 2)));

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
                            AddItem(new TokenItem(ch.ToString(), (i, 1)));
                        }
                        if (data.AfterDeclarator || data.AfterIdentifier || data.IsVarDeclaring)
                        {
                            AddError(ErrorCode.VB0008);
                        }
                        break;

                    #endregion

                    #region [  특수 문자  ]
                    case ';': case '@': case '$': case '^': case '%': case '\\': case '~':
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) break;

                        AddError(ErrorCode.VB0041, new string[] { ch.ToString() });
                        break;

                    #endregion

                    case ':': // 멀티 라인 인식
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) break;

                        if (!Nested && WordRecognition == 0)
                        {
                            AddItem(new LabelItem(FileName, SavingText, Offset));
                            data.AfterLabel = true;
                        }

                        Handled = false;
                        MultiLineRead = true;
                        IsLastChar = true;

                        break;
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
                    if (ch.IsWhiteSpace() || IsLastChar || nextCh.IsDivision())
                    {
                        // string이거나 전처리기 지시문이거나 코멘트일 경우 넘김
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) goto ExitIf;

                        // 마지막 글자이면서 빈칸이 아니라면 인식할 문자열에 추가
                        if ((IsLastChar || nextCh.IsDivision()) && !(ch.IsWhiteSpace() || ch.IsOperator() || ch.IsBracket())
                            && !MultiLineRead) SavingText += ch;

                        if (string.IsNullOrEmpty(SavingText)) goto ExitIf;


                        // 기본적인 오류는 여기서 잡아냄
                        if (ErrorCheck(SavingText.ToLower(), CurrentOffset))
                        {
                            SavingText = string.Empty;
                            goto ExitIf;
                        }
                        
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
                                if (data.AfterAccessor) AddError(ErrorCode.VB0022);
                                data.AfterAccessor = true;
                                data.IsVarDeclaring = true;

                                AddItem(new AccessorItem(Accessor.Dim, "", CurrentOffset));
                                break;
                            #endregion

                            #region [  Declartor  ]

                            case "sub":
                                if (data.AfterEnd) AddItem(new KeywordItem(Keywords.Sub, FileName, CurrentOffset));
                                else AddItem(new DeclaratorItem(DeclaratorType.Sub, CurrentOffset));

                                data.AfterSub = true;
                                data.AfterDeclarator = true;
                                break;
                            case "function":
                                if (data.AfterEnd) AddItem(new KeywordItem(Keywords.Function, FileName, CurrentOffset));
                                else AddItem(new DeclaratorItem(DeclaratorType.Function, CurrentOffset));

                                data.AfterFunction = true;
                                data.AfterDeclarator = true;
                                break;
                            case "property":
                                if (data.AfterEnd) AddItem(new KeywordItem(Keywords.Function, FileName, CurrentOffset));
                                else AddItem(new DeclaratorItem(DeclaratorType.Property, CurrentOffset));

                                data.AfterProperty = true;
                                data.AfterDeclarator = true;
                                break;
                            case "enum":
                                if (data.AfterEnd) AddItem(new KeywordItem(Keywords.Enum, FileName, CurrentOffset));
                                else AddItem(new DeclaratorItem(DeclaratorType.Enum, CurrentOffset));

                                data.AfterEnum = true;
                                data.AfterDeclarator = true;
                                break;
                            #endregion

                            #region [  Conditional  ]
                            case "if":
                                if (data.AfterEnd) AddItem(new KeywordItem(Keywords.If, FileName, CurrentOffset));
                                data.AfterIf = true;

                                break;
                            case "elseif":
                                data.AfterElseIf = true;

                                break;
                            case "else":
                                data.AfterElse = true;
                                break;
                            case "select":
                                data.AfterSelect = true;
                                break;
                            case "case":
                                data.AfterCase = true;
                                break;
                            case "then":
                                data.AfterThen = true;
                                break;
                            #endregion

                            #region [  Iterative  ]

                            case "while":
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
                            case "loop":

                                break;
                            // 특이한 End While의 형태
                            case "wend":
                                if (WordRecognition == 0)
                                {
                                    data.AfterWend = true;
                                    AddItem(new KeywordItem(Keywords.Wend, FileName, CurrentOffset));
                                }
                                else AddError(ErrorCode.VB0082);
                                break;

                            case "for":
                                if (WordRecognition == 0)
                                {
                                    data.AfterFor = true;
                                    AddItem(new KeywordItem(Keywords.For, FileName, CurrentOffset));
                                }
                                break;
                            case "each":
                                if (data.AfterFor)
                                {
                                    data.AfterFor = false;
                                    data.AfterForEach = true;
                                    
                                }
                                else
                                {
                                    AddError(ErrorCode.VB0171);
                                }
                                break;
                            #endregion

                            #region  [  Get/Set/Let  ]
                            case "get":
                                // 오류는 이미 확인
                                if (data.AfterProperty && !data.AfterPropAccessor) data.AfterPropAccessor = true;
                                else AddError(ErrorCode.VB0130, new string[] { "Get" });
                                break;
                            case "set":
                                // 오류는 이미 확인
                                if (data.AfterProperty && !data.AfterPropAccessor) data.AfterPropAccessor = true;
                                else AddError(ErrorCode.VB0130, new string[] { "Set" });

                                break;

                            case "let":
                                if (data.AfterProperty && !data.AfterPropAccessor) data.AfterPropAccessor = true;
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

                            case "end":
                                if (WordRecognition != 0) AddError(ErrorCode.VB0000);
                                else
                                {
                                    AddItem(new KeywordItem(Keywords.End, FileName, (i, 1)));
                                    data.AfterEnd = true;
                                }
                                break;

                            case "as":
                                AddItem(new KeywordItem(Keywords.As, FileName, (i, 1)));
                                data.AfterAs = true;
                                break;

                            case "exit":
                                data.AfterExit = true;

                                break;
                            case "return":
                                data.AfterReturn = true;
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
                                if ((data.AfterDeclarator || (!data.AfterDeclarator && data.AfterAccessor)) && !data.AfterIdentifier)
                                {
                                    if (SavingText.ToLower() == "as") AddError(ErrorCode.VB0045, new string[] { "As" });
                                    
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
                                    else
                                    {
                                        // 식별자 인식

                                        data.AfterDeclarator = false;
                                        data.AfterIdentifier = true;
                                    }

                                    // Public/Private/Dim 뒤에 식별자가 나왔을 경우 변수 선언으로 인식 및 처리
                                    if (LastCodeItm.GetType() == typeof(AccessorItem)) data.IsVarDeclaring = true;

                                    AddItem(new IdentifierItem(SavingText, CurrentOffset));
                                }
                                // 식별자 뒤라면
                                else if (data.AfterIdentifier && !data.AfterAs)
                                {
                                    // As 이어야 함.
                                    if (SavingText.ToLower() != "as" && !data.AfterAs) AddError(ErrorCode.VB0027, new string[] { "As" });
                                    // 식별자 해제
                                    data.AfterIdentifier = false;
                                }
                                // 식별자 이후이면서, As 뒤이고, (Function이거나 변수 선언중) 이라는 걸 동시에 만족 시키면
                                else if (data.AfterIdentifier && data.AfterAs && !data.AfterType)
                                {
                                    // 타입으로 인식
                                    AddItem(new TypeItem(SavingText, CurrentOffset));
                                    data.AfterType = true;
                                }
                                else if (data.AfterArray)
                                {
                                    if (SavingText.IsReservedKeyWords()) AddError(ErrorCode.VB0046);
                                }
                                else
                                {
                                    // int 인식
                                    if (int.TryParse(SavingText, out int DataInt))
                                    {
                                        if (nextCh == '.')
                                        {
                                            bool CanCheck = i + 1 < text.Length ? true : false;
                                            int length = 0;
                                            // TODO : 완성 (식 계산)
                                            if (CanCheck)
                                            {
                                                length++;
                                                // 다음 텍스트 체크
                                                char prevnextCh = i + length + 1 < text.Length ? text[i + length + 1] : '\0';
                                                // 다음 텍스트 체크
                                                if (prevnextCh.IsDigit())
                                                {
                                                    // TODO : 실수 계산식
                                                    // 단, Expression으로 계산하는 방식을 추가시켜야 함.
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        AddItem(new UnknownItem(CurrentOffset));
                                        AddError(ErrorCode.VB0010, new string[] { SavingText });
                                        
                                        char c = nextCh;
                                        int Expctr = 0;
                                        do
                                        {
                                            if (nextCh == '\0') break;
                                            Expctr++;
                                            int txtIndex = i + 1 + Expctr;
                                            if (txtIndex == text.Length) break;
                                            c = text[txtIndex];

                                            // .이 발견되면 Member를 읽으라고 명시
                                            if (c == '.')
                                            {
                                                data.ReadMember = true;
                                            }
                                        } while (c != ' ') ;
                                    }
                                }
                                break;
                        }
                        WordRecognition++;

                        SavingText = string.Empty;
                    }
                    else if (ch.IsLetterOrDigit())
                    {
                        if (data.IsInString || data.IsInPreprocessorDirective || data.IsInComment) goto ExitIf;

                        SavingText += ch;
                        
                    }
                    else if (ch.IsOperator())
                    {
                        // DEBUG : 위쪽에서 처리 되었을 가능성 존재 확인후 삭제
                        AddItem(new TokenItem(SavingText, CurrentOffset));
                    }
                }
                ExitIf:
                // 마지막 문자일시
                if (IsLastChar)
                {
                    if (bracketCount != 0) AddError(ErrorCode.VB0052, new string[] { bracketCount.ToString() });

                    // 전처리기 지시문, 코멘트 등 아이템으로 추가시키기

                    if (data.IsInComment)
                    {
                        AddItem(new CommentItem("", Offset));
                    }
                    else if (data.IsInPreprocessorDirective)
                    {
                        AddItem(new PreprocessorDirectiveItem("", Offset));
                    }
                    // 미완성된 구문 체크

                    if (data.AfterAccessor && !data.AfterDeclarator && !data.AfterIdentifier)
                        AddError(ErrorCode.VB0123);

                    if (data.AfterDeclarator && !data.AfterIdentifier && !data.AfterEnd)
                        AddError(ErrorCode.VB0121);

                    // 식별자, As까지 나왔는데 Type이 안 나온 경우
                    if (data.AfterIdentifier && data.AfterAs && !data.AfterType)
                        AddError(ErrorCode.VB0124);

                    if (data.AfterWhile || data.AfterIf || data.AfterElseIf || (!data.AfterSelect && data.AfterCase) || (data.AfterSelect && data.AfterCase))
                        if (!data.AfterObject && !data.AfterOperator) AddError(ErrorCode.VB0122);

                    if (data.AfterSelect && (!data.AfterCase && !data.AfterEnd))
                        AddError(ErrorCode.VB0076);
                    
                }

                // 만약 첫번째줄이 유지되고 있다면 빈칸, 탭, 새로 띄우기에서 False가 되진 않음
                data.IsFistNonWs &= ch == ' ' || ch == '\t' || ch == '\n' || ch == '\r';

                if (MultiLineRead)
                {

                    new VBASeeker(CodeInfo).GetLine(FileName, CodeLine.Substring(i + 1), Lines, true);
                    i = CodeLine.Length - 1;
                }
            }

            RecognitionList(CurrentLineCodeList);
            CodeInfo.Childrens.Add(CurrentLineCodeList);

            //======================================================================================================================



            // true : Handled   false : Not Handled
            bool ErrorCheck(string Keyword, (int,int) ErrOffset)
            {
                bool Handled = false;
                if (data.IsInString || data.IsInComment || data.IsInPreprocessorDirective) return false;

                if (data.AfterReturn)
                {
                    AddError(ErrorCode.VB0012);
                    Handled = true;
                }

                // class, module 오류
                if (Keyword.ContainsWords(new string[] { "class", "module" }))
                {
                    if (Keyword == "class") AddError(ErrorCode.VB0001);
                    if (Keyword == "module") AddError(ErrorCode.VB0002);

                    AddItem(new UnknownItem(ErrOffset));
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
                    if (Keyword == "dim" && data.AfterDeclarator)
                    {
                        AddError(ErrorCode.VB0021);
                        Handled = true;
                    }
                    if (Keyword == "dim" && WordRecognition != 0)
                    {
                        AddError(ErrorCode.VB0040, new string[] { "Dim" });
                        Handled = true;
                    }
                }
                // Enum, Property, Function, Sub 오류
                if (Keyword.ContainsWords(new string[] { "enum", "property", "function", "sub" }))
                {
                    if (data.AfterDeclarator)
                    {
                        AddError(ErrorCode.VB0021);
                        Handled = true;
                    }
                    else if (!(data.AfterAccessor || WordRecognition == 0) && !(data.AfterEnd && WordRecognition == 1))
                    {
                        AddError(ErrorCode.VB0120);
                        Handled = true;
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
                    if (!(data.AfterProperty && !data.AfterIdentifier))
                    {
                        AddError(ErrorCode.VB0130, new string[] { Keyword });
                        Handled = true;
                    }

                    if (WordRecognition == 0 && Keyword != "let") {
                        if (Keyword == "get") AddError(ErrorCode.VB0003);
                        if (Keyword == "set") AddError(ErrorCode.VB0004);
                        Handled = true;
                    }
                }

                if (Keyword == "as")
                {
                    if (!data.AfterIdentifier)
                    {
                        AddError(ErrorCode.VB0140);
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


                    if (BracketInt == 0)
                    {
                        string ParamString = CodeLine.Substring(TextOffset + 1, j - TextOffset - 1);
                        
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
                                    else
                                    {
                                        NeedExpression = true;
                                    }
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
                                    NeedExpression = true;
                                    break;
                                default:
                                    if (IsInString) continue;
                                    if (i_ch == ' ') continue;
                                    savText += i_ch;
                                    break;
                            }

                            if (i_nextCh.IsDivision() || i_nextCh == ' ' || LastChar)
                            {
                                if (IsInString) continue;
                                if (savText.Replace(' ','\0') == "") continue;

                                switch (savText.ToLower())  
                                {
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
                                    case "as":
                                        if (NeedRest)
                                        {
                                            AddError(ErrorCode.VB0099);
                                            break;
                                        }
                                        if (AfterIdentifier)
                                        {
                                            AfterAs = true;
                                        }
                                        else
                                        {
                                            AddError(ErrorCode.VB0140);
                                        }
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
                                        if ((AfterByRef || AfterByVal || AfterOptional || AfterParamArray) && AfterIdentifier && AfterAs)
                                        {
                                            type = savText;



                                            if (AfterByVal)
                                            {
                                                AddItem(new ParameterItem(ParamAccessor.ByVal, name, type, "","", ParamOffset));
                                                NeedRest = true;
                                            }
                                            else if (AfterByRef)
                                            {
                                                AddItem(new ParameterItem(ParamAccessor.ByRef, name, type, "", "", ParamOffset));
                                                NeedRest = true;
                                            }
                                            else if (AfterOptional)
                                            {
                                                AddItem(new ParameterItem(ParamAccessor.Optional, name, type, "", "", ParamOffset));

                                                UseOptional = true;

                                                AfterParamArray = false; AfterIdentifier = false; AfterAs = false; AfterByRef = false; AfterByVal = false;

                                                AfterOptional = true;
                                                NeedExpression = true;
                                                break;
                                            }
                                            else if (AfterParamArray)
                                            {
                                                AddItem(new ParameterItem(ParamAccessor.ParamArray, name, type, "", "", ParamOffset));
                                                NeedRest = true;   
                                            }
                                            else
                                            {
                                                AddItem(new ParameterItem(ParamAccessor.ByRef, name, type, "", "", ParamOffset));
                                                NeedRest = true;
                                            }

                                            AfterParamArray = false; AfterIdentifier = false; AfterAs = false; AfterByRef = false; AfterByVal = false;

                                            AfterOptional = false;

                                        }
                                        else if (AfterIdentifier && !AfterAs)
                                        {
                                            AddError(ErrorCode.VB0141);
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
            
            string ExpressionRange(int TextOffset, bool RecoBaseOnBracket)
            {
                string str = CodeLine.Substring(TextOffset);
                char[] cArr = str.ToCharArray();

                string tmpstr = "";

                int bracketctr = 0;

                for (int j = TextOffset; j < text.Length; j++)
                {
                    char e_ch = cArr[j];
                    char e_nextCh = j + 1 < cArr.Length ? cArr[j + 1] : '\0';

                    if (e_ch == '(') bracketctr++;
                    if (e_ch == ')') bracketctr--;

                    tmpstr += e_ch;

                    // asign = basign


                    if (bracketctr == 0)
                    {
                        string ExpressionStr = CodeLine.Substring(1, j);

                        return ExpressionStr;
                    }
                }
                return "";
            }

            // 식 계산
            Expression ToExpression(string Expression)
            {
                return new Expression(Expression, CodeInfo.ErrorList, i);
            }
            
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


        void RecognitionList(LineCodeItem codeItm)
        {

            int counter = 0;
            foreach (var itm in codeItm.Childrens)
            {
                //MessageBox.Show(itm.ToString());
            }
        }


    }

}
