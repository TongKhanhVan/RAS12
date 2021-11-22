Public Class TimedMsgBox

    Private Sub TimedMsgBox_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
        End
    End Sub

    Private Sub TimedMsgBox_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TxtMsg.Text = "Maintenance In Progress. Please Come Back in About 5 Mins. Tks"
    End Sub

    Private Sub LblGotIt_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblGotIt.LinkClicked
        Me.Dispose()
        End
    End Sub
End Class