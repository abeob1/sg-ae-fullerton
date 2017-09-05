Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Dim Imp As New oImport
            Imp.Import()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
      
    End Sub
End Class
