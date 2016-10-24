Public Class Form3
    Dim cnn As New OleDb.OleDbConnection
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListView4.Hide()
        cnn = New OleDb.OleDbConnection
        cnn.ConnectionString = "provider = Microsoft.Jet.Oledb.4.0; Data Source =C:\Users\Rony\Documents\BSA.mdb"
        TextBox5.Text = 1
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
        Form1.Show()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        End
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dtr, dtr1, dtr2, dtr3 As OleDb.OleDbDataReader
        Dim dtc, dtc1, dtc2, dtc3 As OleDb.OleDbCommand
        Dim i, tot, b3, b4, b5, x, y As Double
        Dim s, s1, s2, s3, a, b, b1, b2 As String
        b5 = 0
        s = "SELECT Book_Id From Stock"
        s1 = "SELECT Book_Name From Stock "
        s2 = "SELECT Price From Stock "
        s3 = "SELECT Copies From Stock"

        dtc = New OleDb.OleDbCommand(s, cnn)
        dtc1 = New OleDb.OleDbCommand(s1, cnn)
        dtc2 = New OleDb.OleDbCommand(s2, cnn)
        dtc3 = New OleDb.OleDbCommand(s3, cnn)
        cnn.Open()
        dtr = dtc.ExecuteReader
        dtr1 = dtc1.ExecuteReader
        dtr2 = dtc2.ExecuteReader
        dtr3 = dtc3.ExecuteReader
        Form6.TextBox1.Text = TextBox1.Text
        x = TextBox5.Text
        For i = 0 To 50
            If dtr.Read = True And dtr1.Read = True And dtr2.Read = True And dtr3.Read = True Then
                b = dtr(0).ToString
                a = TextBox2.Text
                If String.Compare(b, a, True) = 0 Then
                    b5 = 1
                    Form10.TextBox1.Text = TextBox2.Text
                    b1 = dtr1(0).ToString
                    b2 = dtr2(0).ToString
                    b4 = dtr3(0)
                    Form10.TextBox2.Text = (b4 - x)
                    y = b2 * x
                    tot = y + Val(TextBox3.Text)
                    ListView1.Items.Add(b1)
                    ListView2.Items.Add(x)
                    ListView3.Items.Add(b2)
                    Form6.ListView2.Items.Add(b1)
                    Form6.ListView3.Items.Add(x)
                    Form6.ListView4.Items.Add(b2)
                    TextBox3.Text = Val(tot)
                    For b3 = 1 To 50
                        Form6.ListView1.Items.Add(b3)
                    Next
                    TextBox2.Text = ""
                    TextBox5.Text = 1
               
                End If

            End If
        Next
        If b5 = 0 Then
            MsgBox("Invalid Id")
            TextBox2.Text = ""
        End If
        Form6.TextBox4.Text = TextBox3.Text
        cnn.Close()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Val(TextBox3.Text) > 0 Then
            Me.Hide()
            Form4.Show()
        End If
    End Sub
End Class