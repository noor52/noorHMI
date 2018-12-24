Public Class Form1


    Dim plc As New FA
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load



        plc.Additem("A", "R100")
        plc.Connect("192.168.1.5")
        Timer1.Enabled = True


    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label1.Text = plc.GetItem("A", "R100")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        plc.Disconnect()
        Application.Exit()
    End Sub
End Class
