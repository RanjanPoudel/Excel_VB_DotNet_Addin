Imports System
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Namespace ExcelUdf.Automation
    Public MustInherit Class UdfBase
        <ComRegisterFunction>
        Public Shared Sub ComRegisterFunction(ByVal type As Type)
            Registry.ClassesRoot.CreateSubKey(GetClsIdSubKeyName(type, "Programmable"))
            Dim key = Registry.ClassesRoot.OpenSubKey(GetClsIdSubKeyName(type, "InprocServer32"), True)

            If key Is Nothing Then
                Return
            End If

            key.SetValue("", String.Format("{0}\mscoree.dll", Environment.SystemDirectory), RegistryValueKind.String)
        End Sub

        <ComUnregisterFunction>
        Public Shared Sub ComUnregisterFunction(ByVal type As Type)
            Registry.ClassesRoot.DeleteSubKey(GetClsIdSubKeyName(type, "Programmable"))
        End Sub

        Private Shared Function GetClsIdSubKeyName(ByVal type As Type, ByVal subKeyName As String) As String
            Return String.Format("CLSID\&#123;&#123;&#123;0&#124;&#124;&#124;\{1}", type.GUID.ToString().ToUpper(), subKeyName)
        End Function

        <ComVisible(False)>
        Public Overrides Function ToString() As String
            Return MyBase.ToString()
        End Function

        <ComVisible(False)>
        Public Overrides Function Equals(ByVal obj As Object) As Boolean
            Return MyBase.Equals(obj)
        End Function

        <ComVisible(False)>
        Public Overrides Function GetHashCode() As Integer
            Return MyBase.GetHashCode()
        End Function
    End Class
End Namespace