Public Class Form1
    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If RadioButton1.Checked = True Then
            jk = RadioButton1.Text
        Else
            jk = RadioButton2.Text
        End If

        bagian1 = bagian1a + bagian1b + bagian1c
        bagian2 = bagian2a + bagian2b + bagian2c

        bagian1a = CheckedListBox1.CheckedItems.Count
        bagian1b = CheckedListBox2.CheckedItems.Count
        bagian1c = CheckedListBox3.CheckedItems.Count

        bagian2a = 10 - bagian1a
        bagian2b = 6 - bagian1b
        bagian2c = 5 - bagian1c

        If bagian1 >= 0 And bagian1 <= 7 Then
            hasil = "Rendah"
            Form2.Label1.Text = "Resiko Anda Terkena Covid - 19 Rendah"
        ElseIf bagian1 >= 8 And bagian1 <= 14 Then
            hasil = "Sedang"
            Form2.Label1.Text = "Resiko Anda Terkena Covid - 19 Sedang"
        Else
            hasil = "Tinggi"
            Form2.Label1.Text = "Resiko Anda Terkena Covid - 19 Tinggi"
        End If

        sqlnya = "insert into tb_survey(NIS,Nama,Umur,Jenis_Kelamin,Jumlah_Checklist,Risiko)values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & jk & "','" & bagian1 & "','" & hasil & "')"
        Call go()
        Call summon()

        Form2.Show()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub
End Class
