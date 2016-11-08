Imports hsFunctions.fns
Public Class frmMain
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load, MyBase.DoubleClick
        lblIP.Text = GetIPAddress()
        TextBox1_TextChanged(sender, e)
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) ' Handles TextBox1.TextChanged
        TextBox2.Text = EncryptString(TextBox1.Text, "12345678")
        TextBox3.Text = DecryptString(TextBox2.Text, "12345678")
    End Sub
End Class
