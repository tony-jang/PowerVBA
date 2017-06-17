﻿using ICSharpCode.AvalonEdit.Document;
using PowerVBA.Codes.Extension;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Parsing
{
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
        public ErrorType ErrorType { get; set; }
        public ErrorCode ErrorCode { get; set; }
        public string[] Parameters { get; set; }
        public string FileName { get; set; }
        
        public string Message
        {
            get
            {
                string s = ErrorCode.GetDescription("ko-KR");
                for (int i = 1; i <= ErrorCode.GetReplaceCount(); i++)
                {
                    if (Parameters == null) return "";
                    // 2개 인데 1개 밖에 없을때
                    if (Parameters.Count() < i) break;
                    s = s.Replace($"%{i}", Parameters[i - 1]);
                }

                return s;
            }
        }

        public int Line { get; set; }
        
        public Error(ErrorType errorType, ErrorCode errorCode, string[] parameters, string fileName, int line)
        {
            this.ErrorType = errorType;
            this.ErrorCode = errorCode;
            this.Parameters = parameters;
            this.FileName = fileName;
            this.Line = line;
        }
        
        public Error(ErrorType errorType, string message, int line)
        {
            this.ErrorType = errorType;
            this.Parameters = new string[] { message };
            this.ErrorCode = ErrorCode.VB0000;
            this.Line = line;
        }
        
        public Error(ErrorType errorType, string message, TextLocation location)
        {
            this.ErrorType = errorType;
            this.Parameters = new string[] { message };
            this.Line = location.Line;
        }

        public Error(ErrorType errorType, string message, int line, int col) : this(errorType, message, new TextLocation(line, col))
        {
        }
    }
}