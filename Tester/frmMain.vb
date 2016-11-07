Imports hsFunctions.fn
Public Class frmMain
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load, MyBase.DoubleClick
        lblIP.Text = GetIPAddress()

        TextBox1_TextChanged(sender, e)
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        TextBox2.Text = Encrypt2(TextBox1.Text)
        TextBox3.Text = Decrypt2(TextBox2.Text)
    End Sub
End Class
