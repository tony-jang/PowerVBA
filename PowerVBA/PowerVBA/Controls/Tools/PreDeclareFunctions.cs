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
    
    public class PreDeclareFunction
    {
        public PreDeclareFunction(string Identifier, string Description, string File, string Code, string ReturnData = "", List<DependencyData> Dependencies = null)
        {
            this.Identifier = Identifier;
            this.Description = Description;
            this.File = File;
            this.Code = Code;
            this.ReturnData = ReturnData;
            this.Dependencies = Dependencies;
            this.IsUse = false;
        }
        public string Code { get; set; }
        public string Description { get; set; }
        public string File { get; set; }
        public string Identifier { get; set; }
        public string ReturnData { get; set; }
        public bool IsUse { get; set; }

        /// <summary>
        /// 종속되어 있는 리스트를 나타냅니다.
        /// </summary>
        public List<DependencyData> Dependencies { get; set; }
    }
    public struct DependencyData
    {
        public DependencyData(string DependencyName, DependencyType Type)
        {
            this.DependencyName = DependencyName;
            this.Type = Type;
        }
        /// <summary>
        /// 현재 종속되어 있는 파일 이름을 나타냅니다.
        /// </summary>
        public string DependencyName;
        /// <summary>
        /// 종속되어 있는 오브젝트의 타입을 나타냅니다.
        /// </summary>
        public DependencyType Type { get; set; }


        public enum DependencyType
        {
            CustomFunction,
            DLLFile,
        }
    }
    
    public static class PreDeclareFunctions
    {
        public static List<PreDeclareFunction> Functions { get; set; }

        static PreDeclareFunctions()
        {
            Functions = new List<PreDeclareFunction>();

            #region File : MathEx

            // Min (값 적은 값 체크)
            Functions.Add(new PreDeclareFunction("Min", "두 값을 비교해 더 작은 값을 반환합니다.", "Math", 
                @"Public Function Min(val1 As Integer, val2 As Integer) As Integer
    If val1 < val2 Then 
        Min = val1
    Else
        Min = val2
    End If
End Function
", "Integer"));

            // Max (값 큰 값 체크)
            Functions.Add(new PreDeclareFunction("Max", "두 값을 비교해 더 큰 값을 반환합니다.", "Math", 
                @"Public Function Max(val1 As Integer, val2 As Integer) As Integer
    If val1 > val2 Then 
        Max = val1
    Else
        Max = val2
    End If
End Function", "Integer"));

            #endregion

            #region File : File

            // Create (파일 생성 메소드)
            Functions.Add(new PreDeclareFunction("Create", "파일을 생성합니다.\r\noverride가 False로 설정되면 파일 존재시 파일을 새로 만들지 않습니다.", "File",
                @"Public Function Create(path As String, Optional override As Boolean = False) As Boolean
    On Error GoTo Exception
    If Dir(path) <> """" And override Then
        Dim fso As Object
        Set fso = CreateObject(""Scripting.FileSystemObject"")
        Dim oFile As Object
        Set oFile = fso.CreateTextFile(path)
    End If
    Create = True
Exception:
    If Err.Description <> """" Then
        Create = False
    End If
End Function", "Boolean"));

            // Delete (파일 삭제 메소드)
            Functions.Add(new PreDeclareFunction("Delete", "파일을 삭제합니다. 단, 파일이 존재하지 않거나 오류가 발생하면 False를 반환합니다.", "File",
    @"Public Function Delete(path As String) As Boolean
    If Dir(path) <> """" Then
        Call SetAttr(path, vbNormal)
        Call Kill(path)
        Delete = True
    Else
        Delete = False
    End If
End Function", "Boolean"));

            // Exists (파일 존재 확인 메소드)
            Functions.Add(new PreDeclareFunction("Exists", "파일이 존재하는지에 대한 여부입니다.", "File",
    @"Public Function Exists(path As String) As Boolean
    Exists = (Dir(path) = "")
End Function", "Boolean"));

            // Copy (파일 복사 메소드)
            Functions.Add(new PreDeclareFunction("Copy",
                "파일을 복사합니다. 파일이 존재하지 않거나 오류가 발생하면 False를 반환합니다." +
                "\r\noverride가 False라면 목적지에 파일이 있을때도 False를 반환합니다.", "File",
    @"Public Function Copy(source As String, destination As String, Optional override As Boolean = False) As Boolean
    If Dir(source) <> """" And (Dir(destination) = """" Or override) Then
        Call FileCopy(source, destination)
        Copy = True
    Else
        Copy = False
    End If
End Function", "Boolean"));

            // Move (파일 이동 메소드)
            Functions.Add(new PreDeclareFunction("Move",
                "파일을 이동합니다. 파일이 존재하지 않거나 오류가 발생하면 False를 반환합니다." +
                "\r\noverride가 False라면 목적지에 파일이 있을때도 False를 반환합니다.", "File",
    @"Public Function Move(source As String, destination As String, Optional override As Boolean = False) As Boolean
    If Dir(source) <> """" And (Dir(destination) = """" Or override) Then
        Call FileCopy(source, destination)
        Call Kill(source)
        Copy = True
    Else
        Copy = False
    End If
End Function", "Boolean"));

            #endregion

            #region File : Web

            Functions.Add(new PreDeclareFunction("GetWebSource", "웹 소스에 있는 페이지 소스를 읽어옵니다.", "Web",
                @"Public Function GetWebSource(url As String) As String
    Dim WinHttp As New WinHttpRequest
    WinHttp.Open ""Get"", url
    WinHttp.Send

    GetURL = WinHttp.ResponseText
End Function
", "String",new List<DependencyData>() { new DependencyData("Microsoft WinHTTP Services, version 5.0과 같거나 유사한 DLL 파일이 추가되어 있어야 합니다.", DependencyData.DependencyType.DLLFile)} ));

            #endregion

            #region File : PPTHelper

            Functions.Add(new PreDeclareFunction("Shape", "현재 선택되어 있는 프레젠테이션의 도형을 가져옵니다.", "PPTHelper",
                @"Public Function Shape(SlideNumber As Integer, Name As String) As Shape
    Dim shpe As Shape
    Set shpe = ActivePresentation.Slides(SlideNumber).Shapes(Name)
    Set Shape = shpe
End Function", "Shape"
));

            Functions.Add(new PreDeclareFunction("Slide", "현재 선택되어 있는 프레젠테이션의 슬라이드를 가져옵니다.", "PPTHelper",
                @"Public Function Slide(SlideNumber As Integer) As Slide
    Set Slide = ActivePresentation.Slides(SlideNumber)
End Function", "Slide"
));

            #endregion
        }
    }
}
