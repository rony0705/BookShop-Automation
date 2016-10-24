Public Class Form4
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If RadioButton1.Checked Then
            Dim c, tot, r As Double
            tot = Form3.TextBox3.Text
            c = InputBox("Enter The Amount Paid by Cust.")
            r = c - tot
            If r > 0 Then
                Me.Hide()
                MsgBox("Return Rs" + r.ToString + " To the Cust.")
                Form3.TextBox1.Text = ""
                Form3.TextBox2.Text = ""
                Form3.TextBox3.Text = ""
                Form3.TextBox5.Text = ""
                Form3.ListView1.Items.Clear()
                Form3.ListView2.Items.Clear()
                Form3.ListView3.Items.Clear()
                Form10.Show()
            ElseIf r < 0 Then
                MsgBox("Invalid Payment")
            End If
        ElseIf RadioButton2.Checked Then
            Form3.TextBox1.Text = ""
            Form3.TextBox2.Text = ""
            Form3.TextBox3.Text = ""
            Form3.TextBox5.Text = ""
            Form3.ListView1.Text = ""
            Form3.ListView2.Text = ""
            Form3.ListView3.Text = ""
            Me.Hide()
            Form5.Show()
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs)
        End
    End Sub
    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        End
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        Form3.Show()
    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class