Public Class Form1

    Dim cnn As New OleDb.OleDbConnection

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cnn = New OleDb.OleDbConnection
        cnn.ConnectionString = "provider = Microsoft.Jet.Oledb.4.0; Data Source =C:\Users\Rony\Documents\BSA.mdb"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.Text = "Book Name" Then
            Dim dtr, dtr1, dtr2, dtr3, dtr4, dtr5 As OleDb.OleDbDataReader
            Dim dtc, dtc1, dtc2, dtc3, dtc4, dtc5 As OleDb.OleDbCommand
            Dim i, j, q As Double
            Dim s, s1, s2, s3, s4, s5, a, b, c, b1, b2, b3, b4, b5 As String

            q = 0
            s = "SELECT Book_Name From Stock"
            s1 = "SELECT Book_Id From Stock "
            s2 = "SELECT Author From Stock "
            s3 = "SELECT Publisher From Stock "
            s4 = "SELECT Copies From Stock "
            s5 = "SELECT Price From Stock "
            dtc = New OleDb.OleDbCommand(s, cnn)
            dtc1 = New OleDb.OleDbCommand(s1, cnn)
            dtc2 = New OleDb.OleDbCommand(s2, cnn)
            dtc3 = New OleDb.OleDbCommand(s3, cnn)
            dtc4 = New OleDb.OleDbCommand(s4, cnn)
            dtc5 = New OleDb.OleDbCommand(s5, cnn)
            cnn.Open()
            dtr = dtc.ExecuteReader
            dtr1 = dtc1.ExecuteReader
            dtr2 = dtc2.ExecuteReader
            dtr3 = dtc3.ExecuteReader
            dtr4 = dtc4.ExecuteReader
            dtr5 = dtc5.ExecuteReader
            For i = 0 To 50
                If dtr.Read = True And dtr1.Read = True And dtr2.Read = True And dtr3.Read = True And dtr4.Read = True And dtr5.Read = True Then
                    a = TextBox2.Text
                    b = dtr(0).ToString
                    c = dtr(0).ToString
                    Trim(b)
                    b = Microsoft.VisualBasic.LCase(b)
                    j = InStr(1, b, a)
                    If (j <> 0) Then
                        b1 = dtr1(0).ToString
                        b2 = dtr2(0).ToString
                        b3 = dtr3(0).ToString
                        b4 = dtr4(0).ToString
                        b5 = dtr5(0).ToString

                        Form2.ListView1.Items.Add(c)
                        Form2.ListView2.Items.Add(b1)
                        Form2.ListView3.Items.Add(b2)
                        Form2.ListView4.Items.Add(b3)
                        Form2.ListView5.Items.Add(b4)
                        Form2.ListView6.Items.Add(b5)
                        q = 1
                        Form2.Show()
                    End If
                End If
            Next

            If q = 0 Then
                MsgBox("Book Not Available")
            End If
            TextBox2.Text = ""
            cnn.Close()
        ElseIf ComboBox1.Text = "Author" Then
            Dim dtr, dtr1, dtr2, dtr3, dtr4, dtr5 As OleDb.OleDbDataReader
            Dim dtc, dtc1, dtc2, dtc3, dtc4, dtc5 As OleDb.OleDbCommand
            Dim i, j, q As Double
            Dim s, s1, s2, s3, s4, s5, a, b, c, b1, b2, b3, b4, b5 As String

            q = 0
            s = "SELECT Author From Stock"
            s1 = "SELECT Book_Id From Stock "
            s2 = "SELECT Book_Name From Stock "
            s3 = "SELECT Publisher From Stock "
            s4 = "SELECT Copies From Stock "
            s5 = "SELECT Price From Stock "
            dtc = New OleDb.OleDbCommand(s, cnn)
            dtc1 = New OleDb.OleDbCommand(s1, cnn)
            dtc2 = New OleDb.OleDbCommand(s2, cnn)
            dtc3 = New OleDb.OleDbCommand(s3, cnn)
            dtc4 = New OleDb.OleDbCommand(s4, cnn)
            dtc5 = New OleDb.OleDbCommand(s5, cnn)
            cnn.Open()
            dtr = dtc.ExecuteReader
            dtr1 = dtc1.ExecuteReader
            dtr2 = dtc2.ExecuteReader
            dtr3 = dtc3.ExecuteReader
            dtr4 = dtc4.ExecuteReader
            dtr5 = dtc5.ExecuteReader
            For i = 0 To 50
                If dtr.Read = True And dtr1.Read = True And dtr2.Read = True And dtr3.Read = True And dtr4.Read = True And dtr5.Read = True Then

                    a = TextBox2.Text
                    b = dtr(0).ToString
                    c = dtr(0).ToString
                    Trim(b)
                    b = Microsoft.VisualBasic.LCase(b)
                    j = InStr(1, b, a)
                    If (j <> 0) Then

                        b1 = dtr1(0).ToString
                        b2 = dtr2(0).ToString
                        b3 = dtr3(0).ToString
                        b4 = dtr4(0).ToString
                        b5 = dtr5(0).ToString
                        Form2.ListView1.Items.Add(b2)
                        Form2.ListView2.Items.Add(b1)
                        Form2.ListView3.Items.Add(c)
                        Form2.ListView4.Items.Add(b3)
                        Form2.ListView5.Items.Add(b4)
                        Form2.ListView6.Items.Add(b5)
                        q = 1
                        Form2.Show()
                    End If
                End If
            Next
            If q = 0 Then
                MsgBox("Book Not Available")
            End If
            TextBox2.Text = ""
            cnn.Close()

        ElseIf ComboBox1.Text = "Subject" Then
            Dim dtr, dtr1, dtr2, dtr3, dtr4, dtr5, dtr6 As OleDb.OleDbDataReader
            Dim dtc, dtc1, dtc2, dtc3, dtc4, dtc5, dtc6 As OleDb.OleDbCommand
            Dim i, j, q As Double
            Dim s, s1, s2, s3, s4, s5, s6, a, b, c, b1, b2, b3, b4, b5, b6 As String

            q = 0
            s = "SELECT Subject From Stock"
            s1 = "SELECT Book_Id From Stock "
            s2 = "SELECT Author From Stock "
            s3 = "SELECT Publisher From Stock "
            s4 = "SELECT Copies From Stock "
            s5 = "SELECT Price From Stock "
            s6 = "SELECT Book_Name From Stock"
            dtc = New OleDb.OleDbCommand(s, cnn)
            dtc1 = New OleDb.OleDbCommand(s1, cnn)
            dtc2 = New OleDb.OleDbCommand(s2, cnn)
            dtc3 = New OleDb.OleDbCommand(s3, cnn)
            dtc4 = New OleDb.OleDbCommand(s4, cnn)
            dtc5 = New OleDb.OleDbCommand(s5, cnn)
            dtc6 = New OleDb.OleDbCommand(s6, cnn)
            cnn.Open()
            dtr = dtc.ExecuteReader
            dtr1 = dtc1.ExecuteReader
            dtr2 = dtc2.ExecuteReader
            dtr3 = dtc3.ExecuteReader
            dtr4 = dtc4.ExecuteReader
            dtr5 = dtc5.ExecuteReader
            dtr6 = dtc6.ExecuteReader
            For i = 0 To 50
                If dtr.Read = True And dtr1.Read = True And dtr2.Read = True And dtr3.Read = True And dtr4.Read = True And dtr5.Read = True And dtr6.Read = True Then
                    a = TextBox2.Text
                    b = dtr(0).ToString
                    c = dtr(0).ToString
                    Trim(b)
                    b = Microsoft.VisualBasic.LCase(b)
                    j = InStr(1, b, a)
                    If (j <> 0) Then
                        b1 = dtr1(0).ToString
                        b2 = dtr2(0).ToString
                        b3 = dtr3(0).ToString
                        b4 = dtr4(0).ToString
                        b5 = dtr5(0).ToString
                        b6 = dtr6(0).ToString
                        Form2.ListView1.Items.Add(b6)
                        Form2.ListView2.Items.Add(b1)
                        Form2.ListView3.Items.Add(b2)
                        Form2.ListView4.Items.Add(b3)
                        Form2.ListView5.Items.Add(b4)
                        Form2.ListView6.Items.Add(b5)
                        q = 1
                        Form2.Show()
                    End If
                End If
            Next
            If q = 0 Then
                MsgBox("Book Not Available")
            End If
            TextBox2.Text = ""
            cnn.Close()
        ElseIf ComboBox1.Text = "Publisher" Then
            Dim dtr, dtr1, dtr2, dtr3, dtr4, dtr5 As OleDb.OleDbDataReader
            Dim dtc, dtc1, dtc2, dtc3, dtc4, dtc5 As OleDb.OleDbCommand
            Dim i, j, q As Double
            Dim s, s1, s2, s3, s4, s5, a, b, c, b1, b2, b3, b4, b5 As String
            q = 0
            s = "SELECT Publisher From Stock"
            s1 = "SELECT Book_Id From Stock "
            s2 = "SELECT Author From Stock "
            s3 = "SELECT Book_Name From Stock "
            s4 = "SELECT Copies From Stock "
            s5 = "SELECT Price From Stock "
            dtc = New OleDb.OleDbCommand(s, cnn)
            dtc1 = New OleDb.OleDbCommand(s1, cnn)
            dtc2 = New OleDb.OleDbCommand(s2, cnn)
            dtc3 = New OleDb.OleDbCommand(s3, cnn)
            dtc4 = New OleDb.OleDbCommand(s4, cnn)
            dtc5 = New OleDb.OleDbCommand(s5, cnn)
            cnn.Open()
            dtr = dtc.ExecuteReader
            dtr1 = dtc1.ExecuteReader
            dtr2 = dtc2.ExecuteReader
            dtr3 = dtc3.ExecuteReader
            dtr4 = dtc4.ExecuteReader
            dtr5 = dtc5.ExecuteReader
            For i = 0 To 50
                If dtr.Read = True And dtr1.Read = True And dtr2.Read = True And dtr3.Read = True And dtr4.Read = True And dtr5.Read = True Then
                    a = TextBox2.Text
                    b = dtr(0).ToString
                    c = dtr(0).ToString
                    Trim(b)
                    b = Microsoft.VisualBasic.LCase(b)
                    j = InStr(1, b, a)
                    If (j = 0) Then
                        b1 = dtr1(0).ToString
                        b2 = dtr2(0).ToString
                        b3 = dtr3(0).ToString
                        b4 = dtr4(0).ToString
                        b5 = dtr5(0).ToString
                        Form2.ListView1.Items.Add(b3)
                        Form2.ListView2.Items.Add(b1)
                        Form2.ListView3.Items.Add(b2)
                        Form2.ListView4.Items.Add(c)
                        Form2.ListView5.Items.Add(b4)
                        Form2.ListView6.Items.Add(b5)
                        q = 1
                        Form2.Show()
                    End If
                End If
            Next
            If q = 0 Then
                MsgBox("Book Not Available")
            End If
            TextBox2.Text = ""
            cnn.Close()
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Form3.Show()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        End
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form6.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Hide()
        Form8.Show()
    End Sub
End Class





