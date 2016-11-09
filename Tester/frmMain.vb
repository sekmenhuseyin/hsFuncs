Imports hsFunctions.fns
Public Class frmMain
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load, MyBase.DoubleClick
        lblIP.Text = GetIPAddress()
        TextBox1_TextChanged(sender, e)
    End Sub

    Private Sub frmMain_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        TextBox2.Width = Me.Width - 45
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) ' Handles TextBox1.TextChanged
        TextBox2.Text = Encrypt(TextBox1.Text, Application.CompanyName)
        TextBox3.Text = Decrypt(TextBox2.Text, Application.CompanyName)
    End Sub
End Class
