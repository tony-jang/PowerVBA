using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerVBA.Codes.Parsing
{
    public struct LinePoint
    {
        public LinePoint(int line, int offset, int length)
        {
            this.Line = line;
            this.Offset = offset;
            this.Length = length;
        }
        

        /// <summary>
        /// 현재 위치한 줄입니다.
        /// </summary>
        int Line { get; set; }
        /// <summary>
        /// 오프셋입니다. 오프셋은 Line 기준입니다.
        /// </summary>
        int Offset { get; set; }
        /// <summary>
        /// 오프셋으로 부터의 길이입니다.
        /// </summary>
        int Length { get; set; }

        #region [  IsOverlapped Method  ]

        /// <summary>
        /// 라인 포인트가 겹쳐있는 부분이 있는지에 대해서 확인합니다.
        /// (Length가 0이여도 Offset 위치를 기준으로 함)
        /// </summary>
        public bool IsOverlapped(LinePoint linePoint)
        {
            if (linePoint == this)
                return true;

            if (linePoint.Line == this.Line)
            {
                if (linePoint.Offset <= this.Offset + this.Length - 1 &&
                    linePoint.Offset >= this.Offset + this.Length - 1 ||
                    this.Offset <= linePoint.Offset + linePoint.Length - 1 &&
                    this.Offset >= linePoint.Offset + linePoint.Length - 1)
                {
                    return true;
                }
            }
            return false;
        }

        #endregion
        
        #region [  IsBackward Method  ]

        /// <summary>
        /// 비교 라인보다 해당 라인이 뒤에 있는지를 체크합니다. 단, 겹치면 False를 반환합니다.
        /// </summary>
        /// <param name="linePoint">비교할 라인입니다.</param>
        public bool IsBackward(LinePoint linePoint)
        {
            if (this.Line < linePoint.Line)
                return false;
            if (this.Line > linePoint.Line)
                return true;

            if (this.Line == linePoint.Line)
            {
                if (this.Offset + this.Length - 1 < linePoint.Offset)
                {
                    return true;
                }
            }

            return false;
        }

        #endregion
        
        #region [  IsForward Method  ]
        /// <summary>
        /// 비교 라인보다 해당 라인이 뒤에 있는지를 체크합니다. 단, 겹치면 False를 반환합니다.
        /// </summary>
        /// <param name="linePoint">비교할 라인입니다.</param>
        public bool IsForward(LinePoint linePoint)
        {
            if (this.Line > linePoint.Line)
                return false;
            if (this.Line < linePoint.Line)
                return true;

            if (this.Line == linePoint.Line)
            {
                if (linePoint.Offset + linePoint.Length - 1 < this.Offset)
                {
                    return true;
                }
            }

            return false;
        }


        #endregion

        #region [  Override & Operator  ]

        public static bool operator ==(LinePoint left, LinePoint right)
        {
            return (left.Length == right.Length &&
                    left.Line == right.Line &&
                    left.Offset == right.Offset);
        }

        public static bool operator !=(LinePoint left, LinePoint right)
        {
            return (left.Length != right.Length &&
                    left.Line != right.Line &&
                    left.Offset != right.Offset);
        }

        /// <summary>
        /// int를 줄 정보만 넣은 LinePoint로 변환합니다.
        /// </summary>
        /// <param name="i"></param>
        public static implicit operator LinePoint(int i)
        {
            return new LinePoint(i, 0, 0);
        }


        public override bool Equals(object obj)
        {
            if (obj is LinePoint lp)
            {
                return (Length == lp.Length &&
                    Line == lp.Line &&
                    Offset == lp.Offset);
            }

            return false;
        }

        public override string ToString()
        {
            return $"Line:{Line} Offset:{Offset} Length:{Length}";
        }

        public override int GetHashCode()
        {
            return Offset.GetHashCode();
        }

        #endregion
    }
}
