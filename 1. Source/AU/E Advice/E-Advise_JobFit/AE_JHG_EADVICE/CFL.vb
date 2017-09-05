Public Class CFL

   
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If Me.TextBox1.Text = 1 Then
            sCFL = Me.CFL_Customer.Item(0, Me.CFL_Customer.CurrentRow.Index).Value
            sCFL1 = Me.CFL_Customer.Item(1, Me.CFL_Customer.CurrentRow.Index).Value
            ''SOA_F = New frmSOA
            ''SOA_F.txtGroup.Text = sCFL
            ' frmSOA.txtGroup.Text = sCFL
            ''frmSOA.TextBox1.Text = sCFL
            frmSOA.Refresh()
        ElseIf Me.TextBox1.Text = 3 Then
            sCFL = Me.CFL_Customer.Item(0, Me.CFL_Customer.CurrentRow.Index).Value

        ElseIf Me.TextBox1.Text = 4 Then
            sCFL = Me.CFL_Customer.Item(0, Me.CFL_Customer.CurrentRow.Index).Value
            sCFL1 = Me.CFL_Customer.Item(1, Me.CFL_Customer.CurrentRow.Index).Value
        End If
        Me.Close()
        'MsgBox(CFL_Customer.Item("CardCode", CFL_Customer.CurrentRow.Index).Value)
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Txt_Find_TextChanged(sender As Object, e As System.EventArgs) Handles Txt_Find.TextChanged
        Try

            For Each row As DataGridViewRow In CFL_Customer.Rows
                If row.Cells(1).Value.ToString.Length > 0 Then
                    If row.Cells(1).Value.ToString.Length >= Txt_Find.TextLength Then
                        If LCase(row.Cells(1).Value.ToString.Substring(0, Txt_Find.Text.Length)) = LCase(Txt_Find.Text) Then
                            Dim _index As Integer = row.Index
                            row.Selected = True
                            CFL_Customer.CurrentCell = CFL_Customer.Rows(row.Index).Cells(0)
                            ' Me.CFL_Customer.Rows(row.Index).Selected = True
                            Exit For
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    Public Sub New(ByVal sSQLP As String, ByVal sCond As String)

        ' This call is required by the designer.
        InitializeComponent()
        Dim sErrDesc As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oDS As DataSet = Nothing
        ''  sSQL = "SELECT T0.""GroupCode"",T0.""GroupName"" FROM " & """" & sCompanyName & """" & " .OCRG T0 WHERE T0.""GroupType"" = 'C'"
        oDS = ExecuteSQLQuery(sSQLP)
        Me.CFL_Customer.DataSource = oDS.Tables(0)
        Me.CFL_Customer.Columns(1).Width = 300
        Me.TextBox1.Text = sCond
        Me.CFL_Customer.Rows(0).Selected = False
        Txt_Find.Select()


        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub CFL_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
       
    End Sub

    Private Sub CFL_Customer_Click(sender As Object, e As System.EventArgs) Handles CFL_Customer.Click
        Call Button1_Click(Me, New System.EventArgs)
    End Sub

    Protected Overloads Overrides Function ProcessDialogKey(ByVal keyData As Keys) As Boolean
        If keyData = Keys.Enter Then
            'MyBase.ProcessTabKey(Keys.Tab)
            Call Button1_Click(Me, New System.EventArgs)
            Return True
        End If
        Return MyBase.ProcessDialogKey(keyData)
    End Function


    
End Class