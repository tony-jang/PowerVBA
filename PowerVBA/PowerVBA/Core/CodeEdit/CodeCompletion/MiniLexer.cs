using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.CodeEdit.CodeCompletion
{
    class MiniLexer
    {
        readonly string text;

        /// <summary>
        /// 
        /// </summary>
        public bool IsFistNonWs = true;
        /// <summary>
        /// 한줄 주석인지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsInSingleComment = false;
        /// <summary>
        /// String인지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsInString = false;
        /// <summary>
        /// 그대로의 String인지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsInVerbatimString = false;
        /// <summary>
        /// Char인지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsInChar = false;
        /// <summary>
        /// 전처리기 지시문인지에 대한 여부를 가져옵니다.
        /// </summary>
        public bool IsInPreprocessorDirective = false;

        public MiniLexer(string text)
        {
            this.text = text;
        }

        /// <summary>
        /// 거의 모든 char에서 모든 텍스트를 파싱하고 대리자를 호출합니다.
        /// 주석의 시작을 건너뛰고 축어적 문자열의 시작과 이스케이프 문자를 건너뜁니다.
        /// </summary>
        /// <param name="act">true를 반환하는 것으로 파싱을 중단합니다. 정수 인수는 텍스트의 오프셋입니다.</param>
        /// <returns>중단 되었을때 True를 반환합니다.</returns>
        public bool Parse(Func<char, int, bool> act = null)
        {
            return Parse(0, text.Length, act);
        }


        /// <summary>
        /// 텍스트를 처음부터 시작+길이로 파싱하고 거의 모든 문자에 대해서 대리자를 호출합니다.
        /// 주석의 시작을 건너뛰고 축어적 문자열의 시작과 이스케이프 문자를 건너뜁니다.
        /// </summary>
        /// <param name="start">시작 오프셋입니다.</param>
        /// <param name="length">파싱할 길이입니다.</param>
        /// <param name="act">true를 반환하는 것으로 파싱을 중단합니다. 정수 인수는 텍스트의 오프셋입니다.</param>
        /// <returns>중단 되었을때 True를 반환합니다.</returns>
        public bool Parse(int start, int length, Func<char, int, bool> act = null)
        {
            
            for (int i = start; i < length; i++)
            {
                char ch = text[i];
                char nextCh = i + 1 < text.Length ? text[i + 1] : '\0';
                switch (ch)
                {
                    case '#': // 전처리기 지시문

                        // 첫번째 문자가 #일 경우 전처리기 지시문으로 인식함
                        if (IsFistNonWs)
                            IsInPreprocessorDirective = true;
                        break;
                    case '\'': // 주석

                        // String 중이거나, Char 내부이거나, Comment내부라면
                        if (IsInString || IsInChar || IsInVerbatimString || IsInSingleComment)
                            break;
                        
                        IsInSingleComment = true;
                        IsInPreprocessorDirective = false;
                        break;
                    case '\n': // 새 줄 인식될때 초기화
                    case '\r': // 새 줄 인식될때 초기화
                        IsInSingleComment = false;
                        IsInString = false;
                        IsInChar = false;
                        // 첫줄이라고 인식 시켜주기
                        IsFistNonWs = true;
                        IsInPreprocessorDirective = false;
                        break;
                    case '"': // string 문자열 시작
                        if (IsInSingleComment || IsInChar)
                            break;
                        if (IsInVerbatimString)
                        {
                            if (nextCh == '"')
                            {
                                i++;
                                break;
                            }
                            IsInVerbatimString = false;
                            break;
                        }
                        IsInString = !IsInString;
                        break;
                }
                if (act != null)
                    if (act(ch, i))
                        return true;
                IsFistNonWs &= ch == ' ' || ch == '\t' || ch == '\n' || ch == '\r';
            }
            return false;
        }
    }
}
