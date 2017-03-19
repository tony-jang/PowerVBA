using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Controls.Tools
{
    //////////////////////////////////////////////////////////////////
    //                                                              //
    //                PreDeclare Functions Ver 1.0                  //
    //                                                              //
    //       미리 선언된 Function들에 대해서 정리해놓습니다.        //
    //                                                              //
    //////////////////////////////////////////////////////////////////

    public static class PreDeclareFunctions
    {
        public static List<(string, string)> Functions = new List<(string, string)>();

        static PreDeclareFunctions()
        {
            // 테스트 : 곱하기 함수
            Functions.Add(("Multiply", "Public Function Multiply(a As Integer, b As Integer) As Integer" +
                "   Multiply = a * b" +
                "End Function"));
        }
    }
}
