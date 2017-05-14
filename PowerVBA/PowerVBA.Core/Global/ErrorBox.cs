using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerVBA.Core.Global
{
    public static class ErrorBox
    {
        /*
        
        */

        public static void Show(string message, ErrorType type)
        {
            StringBuilder sb = new StringBuilder();

            switch (type)
            {
                case ErrorType.Unexpected:
                    sb.AppendLine("프로그램에서 예기치 못한 오류가 발생했습니다.");
                    sb.AppendLine();

                    sb.AppendLine("다음의 내용은 오류에 대한 자세한 정보입니다.");
                    break;
                case ErrorType.Resource:
                    sb.AppendLine("리소스 오류가 발생했습니다.");

                    sb.Append("보통 이런 오류는 프로그램 내부에 포함되어 있는 리소스의 내용이 잘못되어 발생하는 경우가 대부분이거나, ");
                    sb.AppendLine("코드에서 오류가 발생했을 수도 있습니다.");

                    sb.AppendLine();
                    sb.AppendLine("다음의 내용은 리소스 오류에 대한 자세한 정보입니다.");
                    break;
                case ErrorType.User:
                    sb.AppendLine("사용자 오류가 발생했습니다.");
                    sb.AppendLine();
                    sb.AppendLine("다음의 내용은 사용자 오류에 대한 자세한 정보입니다.");
                    break;
            }

            sb.AppendLine();
            sb.AppendLine(message);

            MessageBox.Show(sb.ToString());
        }
        public static void Show(string message, string name)
        {

        }
    }
    public enum ErrorType
    {
        /// <summary>
        /// 예기치 못한 오류 (단순히 Try .. Catch로만 잡은 오류)
        /// </summary>
        Unexpected = 0,
        /// <summary>
        /// 리소스 오류 (Code의 Attribute의 경우)
        /// </summary>
        Resource = 1,
        /// <summary>
        /// 사용자가 사용하다가 가변 값에 의한 오류
        /// </summary>
        User = 2,
    }
}
