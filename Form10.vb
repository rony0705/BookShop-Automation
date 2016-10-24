Public Class Form10
    Dim cmn As New OleDb.OleDbConnection
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Hide()
        TextBox2.Hide()
        cmn = New OleDb.OleDbConnection
        cmn.ConnectionString = "provider=microsoft.jet.oledb.4.0; Data source=C:\Users\Rony\Documents\BSA.mdb"
        Dim dtReader As OleDb.OleDbDataReader
        Dim dtCommand As OleDb.OleDbCommand
        Dim sqlstr As String
        sqlstr = "select Book_id from Stock"
        dtCommand = New OleDb.OleDbCommand(sqlstr, cmn)
        cmn.Open()
        dtReader = dtCommand.ExecuteReader
        cmn.Close()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cmd As New OleDb.OleDbCommand
        If Not cmn.State = ConnectionState.Open Then
            cmn.Open()
        End If
        cmd.Connection = cmn
        cmd.CommandText = "update Stock set Copies=" & Me.TextBox2.Text & " Where Book_id=" & Me.TextBox1.Text & ""
        cmd.ExecuteNonQuery()
        cmn.Close()
        MsgBox("Data successfully updated.......")
        Me.Hide()
        Form1.Show()
        Form3.ListView1.Items.Clear()
        Form3.ListView2.Items.Clear()
        Form3.ListView3.Items.Clear()
    End Sub
End Class