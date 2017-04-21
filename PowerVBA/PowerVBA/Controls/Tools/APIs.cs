using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Controls.Tools
{
    static class APIs
    {
        public static List<APIInfo> APIList { get; set; }
        static APIs()
        {
            APIList = new List<APIInfo>();
            APIList.Add(new APIInfo("Sleep", string.Empty, "입력된 기간 동안 실행을 일시 중단합니다. ",
                                    "Sleep 함수는 지정된 기간 동안 실행을 일시 중단합니다. " +
                                    "함수에 전달 된 시간 (밀리초 :: 1000ms가 1초) 동안 실행 코드를 비활성 상태로 만듭니다.",
                                    @"Private Declare Sub Sleep Lib ""kernel32"" (ByVal dwMilliseconds As Long)"));

            APIList.Add(new APIInfo("GetCursorPos", "Long", "현재 커서의 위치를 반환합니다. ",
                                    "GetCursorPos를 사용하면 " +
                                    "함수에 전달 된 시간 (밀리초 :: 1000ms가 1초) 동안 실행 코드를 비활성 상태로 만듭니다. " +
                                    "X (int), Y (int)를 갖는 Type을 직접 선언해야 합니다. (추후 업데이트로 지원 예정)",
                                    @"Private Declare Function GetCursorPos Lib ""user32"" (lpPoint As POINTAPI) As Long"));

            APIList.Add(new APIInfo("BringWindowToTop", "Long", "지정된 창을 맨위에 표시해줍니다.",
                                    "이 API 함수는 지정된 창을 맨 위에 표시합니다." +
                                    "윈도우가 자식 윈도우 인 경우에 이 함수는 연결된 최상위 부모 윈도우를 활성화합니다." +
                                    "적절한 창 핸들을 전달하세요. " +
                                    "함수가 실패하면 0을 반환합니다." +
                                    "성공하면 0이 아닌 값을 반환합니다.",
                                    @"Private Declare Function BringWindowToTop Lib ""user32"" (ByVal lngHWnd As Long) As Long"));

            APIList.Sort();

        }
    }

    public class APIInfo : IComparable
    {
        public APIInfo(string Name, string ReturnData, string Descrption, string FullDescription, string APIStr)
        {
            this.Name = Name;
            this.ReturnData = ReturnData;
            this.Description = Descrption;
            this.APIStr = APIStr;
            this.FullDescription = FullDescription;
        }
        public string Name { get; set; }
        public string ReturnData { get; set; }
        public string Description { get; set; }
        public string FullDescription { get; set; }
        public string APIStr { get; set; }

        public int CompareTo(object obj)
        {
            return Name.CompareTo(((APIInfo)obj).Name);
        }
    }
}
