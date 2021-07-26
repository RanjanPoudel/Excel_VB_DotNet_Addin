Imports System.Runtime.InteropServices

Namespace ExcelUdf.Automation


    <ComClass(ComClass_Jarden.ClassId, ComClass_Jarden.InterfaceId, ComClass_Jarden.EventsId)>
    Public Class ComClass_Jarden

        Public Const ClassId As String = "F0056CA0-9096-414B-A193-990F3D2974C2"
        Public Const InterfaceId As String = "571BC8A0-93AE-49ED-AC08-EC8FBDAF9B27"
        Public Const EventsId As String = "8761401F-097C-4C89-B606-8847F0EDFFA3"

        Public Sub New()
            MyBase.New()
        End Sub

        Public Function Jarden_AddNumber(FirstNumber As Double, SecondNumber As Double) As Double
            Return FirstNumber + SecondNumber
        End Function


    End Class


End Namespace
