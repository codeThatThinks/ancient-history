Public Class frmHowTo
    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Me.Close()
        frmMain.pictCanvas.Focus()
    End Sub
End Class