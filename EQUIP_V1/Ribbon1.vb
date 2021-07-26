Imports Microsoft.Office.Tools.Ribbon


Public Class Ribbon1

    Public myUserControl1 As MyUserControl
    Public WithEvents myCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs)
        MsgBox("Hello from Ribbon!")
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        ' this is creation of task pane.

        myUserControl1 = New MyUserControl()
        myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(myUserControl1, "EQUIP Add-in")
        myCustomTaskPane.Visible = True

        With myCustomTaskPane
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
            .Height = 500
            .Width = 500
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
            .Width = 300
            .Visible = True
        End With


    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click

    End Sub
End Class
