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
        public static List<(string, string, string)> Functions = new List<(string, string, string)>();

        static PreDeclareFunctions()
        {
            #region File : MathEx

            Functions.Add(("MathEx", "Min",
@"'두 값을 비교해 더 작은 값을 반환합니다.
Public Function Min(val1 As Integer, val2 As Integer) As Integer
    If val1 < val2 Then 
        Min = val1
    Else
        Min = val2
    End If
End Function"));

            Functions.Add(("MathEx", "Max",
@"'두 값을 비교해 더 큰 값을 반환합니다.
Public Function Max(val1 As Integer, val2 As Integer) As Integer
    If val1 > val2 Then 
        Max = val1
    Else
        Max = val2
    End If
End Function"));

            #endregion

            #region File : File

            // Create (파일 생성 메소드)
            Functions.Add(("File", "Create",
                @"
' 파일을 생성합니다. override가 False로 설정되면 파일 존재시 파일을 새로 만들지 않습니다.
Public Function Create(path As String, Optional override As Boolean = False) As Boolean
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
End Function
"));

            // Delete (파일 삭제 메소드)
            Functions.Add(("File", "Delete",
                @"
' 파일을 삭제합니다. 단, 파일이 존재하지 않거나 오류가 발생하면 False를 반환합니다.
Public Function Delete(path As String) As Boolean
    If Dir(path) <> """" Then
        Call SetAttr(path, vbNormal)
        Call Kill(path)
        Delete = True
    Else
        Delete = False
    End If
End Function
"));

            // Exists (파일 존재 확인 메소드)
            Functions.Add(("File", "Exists",
                @"
' 파일이 존재하는지에 대한 여부입니다.
Public Function Exists(path As String) As Boolean
    Exists = (Dir(path) = "")
End Function"));

            // Copy (파일 복사 메소드)
            Functions.Add(("File","Copy",
                @"
' 파일을 복사합니다. 파일이 존재하지 않거나 오류가 발생하면 False를 반환합니다. override가 False라면 목적지에 파일이 있을때도 False를 반환합니다.
Public Function Copy(source As String, destination As String, Optional override As Boolean = False) As Boolean
    If Dir(source) <> """" And (Dir(destination) = """" Or override) Then
        Call FileCopy(source, destination)
        Copy = True
    Else
        Copy = False
    End If
    
End Function
"));
            
            #endregion
        }
    }
}
