using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Codes.Extension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{

    /// <summary>
    /// Enum that describes the type of an error.
    /// </summary>
    public enum ErrorType
    {
        Unknown,
        Error,
        Warning
    }

    /// <summary>
    /// 파싱중에 오류가 발생한 것을 설명합니다.
    /// </summary>
    [Serializable]
    public class Error
    {
        readonly ErrorType errorType;
        readonly ErrorCode errorCode;
        readonly DomRegion region;
        readonly string[] parameters;

        public ErrorCode ErrorCode { get => errorCode; }


        /// <summary>
        /// The type of the error.
        /// </summary>
        public ErrorType ErrorType { get => errorType; }

        /// <summary>
        /// The error description.
        /// </summary>
        public string Message
        {
            get
            {
                string s = errorCode.GetDescription("ko-KR");
                for (int i=1; i<=errorCode.GetReplaceCount(); i++)
                {
                    // 2개 인데 1개 밖에 없을때
                    if (parameters.Count() < i) break;
                    s = s.Replace($"%{i}", parameters[i-1]);
                }

                return s;
            }
        }


        /// <summary>
        /// The region of the error.
        /// </summary>
        public DomRegion Region { get => region; }

        public int Line { get => Region.BeginLine; }

        public Error(ErrorType errorType, ErrorCode errorCode, string[] parameters, DomRegion region)
        {
            this.errorType = errorType;
            this.errorCode = errorCode;
            this.parameters = parameters;
            this.region = region;
        }
        

        /// <summary>
        /// Initializes a new instance of the <see cref="ICSharpCode.NRefactory.TypeSystem.Error"/> class.
        /// </summary>
        /// <param name='errorType'>
        /// The error type.
        /// </param>
        /// <param name='message'>
        /// The description of the error.
        /// </param>
        /// <param name='region'>
        /// The region of the error.
        /// </param>
        public Error(ErrorType errorType, string message, DomRegion region)
        {
            this.errorType = errorType;
            this.parameters = new string[] { message };
            this.region = region;
            this.errorCode = ErrorCode.VB0000;
        }
        

        /// <summary>
        /// Initializes a new instance of the <see cref="ICSharpCode.NRefactory.TypeSystem.Error"/> class.
        /// </summary>
        /// <param name='errorType'>
        /// The error type.
        /// </param>
        /// <param name='message'>
        /// The description of the error.
        /// </param>
        /// <param name='location'>
        /// The location of the error.
        /// </param>
        public Error(ErrorType errorType, string message, TextLocation location)
        {
            this.errorType = errorType;
            this.parameters = new string[] { message };
            this.region = new DomRegion(location, location);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ICSharpCode.NRefactory.TypeSystem.Error"/> class.
        /// </summary>
        /// <param name='errorType'>
        /// The error type.
        /// </param>
        /// <param name='message'>
        /// The description of the error.
        /// </param>
        /// <param name='line'>
        /// The line of the error.
        /// </param>
        /// <param name='col'>
        /// The column of the error.
        /// </param>
        public Error(ErrorType errorType, string message, int line, int col) : this(errorType, message, new TextLocation(line, col))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ICSharpCode.NRefactory.TypeSystem.Error"/> class.
        /// </summary>
        /// <param name='errorType'>
        /// The error type.
        /// </param>
        /// <param name='message'>
        /// The description of the error.
        /// </param>
        public Error(ErrorType errorType, string message)
        {
            this.errorType = errorType;
            this.parameters = new string[] { message };
            this.region = DomRegion.Empty;
        }
    }
}
