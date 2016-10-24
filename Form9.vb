Public Class Form9
    Dim cmn As New OleDb.OleDbConnection
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cmn.ConnectionString = "provider=Microsoft.jet.oledb.4.0; Data source=C:\Users\Rony\Documents\BSA.mdb"
        Dim dtReader As OleDb.OleDbDataReader
        Dim dtCommand As OleDb.OleDbCommand
        Dim sqlstr As String
        sqlstr = "select Book_id from Stock"
        dtCommand = New OleDb.OleDbCommand(sqlstr, cmn)
        cmn.Open()
        dtReader = dtCommand.ExecuteReader
        While dtReader.Read = True
            ComboBox1.Items.Add(dtReader(0))
        End While
        cmn.Close()
        ComboBox1.Text = ""
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cmd As New OleDb.OleDbCommand
        If Not cmn.State = ConnectionState.Open Then
            cmn.Open()
        End If
        cmd.Connection = cmn
        cmd.CommandText = "delete from Stock where Book_id = " & Me.ComboBox1.SelectedItem & ""
        Refresh()
        cmd.ExecuteNonQuery()
        cmn.Close()
        MsgBox("Data Succesfully Deleted............")
        ComboBox1.Items.Remove(ComboBox1.SelectedItem)
        ComboBox1.Text = ""
        TextBox1.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        ComboBox1.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Form1.Show()
    End Sub
    Private Sub ComboBox1_SeletedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim dtReader As OleDb.OleDbDataReader
        Dim dtCommand As OleDb.OleDbCommand
        Dim sqlstr As String
        sqlstr = "select Book_Name,Author,Subject,Publisher,Volume,Copies,Price from Stock where Book_id = " & Me.ComboBox1.SelectedItem & ""
        dtCommand = New OleDb.OleDbCommand(sqlstr, cmn)
        If Not cmn.State = ConnectionState.Open Then
            cmn.Open()
        End If
        dtReader = dtCommand.ExecuteReader
        If dtReader.Read = True Then
            TextBox1.Text = dtReader(0)
            TextBox3.Text = dtReader(1)
            TextBox4.Text = dtReader(2)
            TextBox5.Text = dtReader(3)
            TextBox6.Text = dtReader(4)
            TextBox7.Text = dtReader(5)
            TextBox8.Text = dtReader(6)
        End If
        cmn.Close()
    End Sub
End Class