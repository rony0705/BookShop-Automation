Public Class Form7
    Dim cmn As OleDb.OleDbConnection
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cmn = New OleDb.OleDbConnection()
        cmn.ConnectionString = "provider=Microsoft.jet.oledb.4.0; Data source=C:\Users\Rony\Documents\BSA.mdb"
        TextBox1.Focus()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cmd As New OleDb.OleDbCommand
        If Not cmn.State = ConnectionState.Open Then
            cmn.Open()
        End If
        cmd.Connection = cmn
        cmd.CommandText = "INSERT INTO Stock(Book_id,Book_Name,Author,Subject,Publisher,Volume,Copies,Price)Values(" & Me.TextBox1.Text & ",'" & Me.TextBox2.Text & "','" & Me.TextBox3.Text & "','" & Me.TextBox6.Text & "','" & Me.TextBox4.Text & "'," & Me.TextBox7.Text & "," & Me.TextBox5.Text & "," & Me.TextBox8.Text & ")"
        cmd.ExecuteNonQuery()
        Refresh()
        cmn.Close()
        MsgBox("Data Sucessfully Inserted.....")
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Form1.Show()
    End Sub
    Private Sub textbox1_lostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        Dim DtReader, dtr As OleDb.OleDbDataReader
        Dim DtCommand, dtc As OleDb.OleDbCommand
        Dim SqlStr, s As String
        SqlStr = "select Book_id from Stock"
        s = "SELECT SL_NO From Stock"
        DtCommand = New OleDb.OleDbCommand(SqlStr, cmn)
        dtc = New OleDb.OleDbCommand(s, cmn)
        cmn.Open()
        DtReader = DtCommand.ExecuteReader
        dtr = dtc.ExecuteReader
        While DtReader.Read = True
            If DtReader(0) = Val(TextBox1.Text) Then
                MsgBox("BOOK ID Already Exist.......")
                TextBox1.Text = ""
                TextBox1.Focus()
            End If
        End While
        cmn.Close()
    End Sub
End Class