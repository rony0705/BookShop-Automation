Public Class Form5
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Focus()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim s, s1 As String
        Dim l, l1, f As Double
        f = 0
        s = TextBox1.Text.ToString
        l = Len(s)
        If l <> 16 Then
            MsgBox("Invalid Card Number")
            f = 1
        Else
            s1 = TextBox2.Text.ToString
            l1 = Len(s1)
            If l1 <> 3 Then
                MsgBox("Invalid CVV")
                f = 1
            End If
            If f <> 1 Then
                Me.Hide()
                MsgBox("Transaction Complete")
            End If
        End If
        Form10.Show()
    End Sub
End Class