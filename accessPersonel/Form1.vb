Imports System.Data.OleDb
Public Class Form1
    Dim baglanti As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Onyxia\Desktop\Database1.accdb")
    Private Sub temizle()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Ekle Button
        If TextBox1.Text = "Ekle" Then
            baglanti.Open()
            Dim komut As New OleDbCommand("insert into bilgi(ad,soyad,sehir,telefon,email)values('" + TextBox1.Text + "','" + TextBox2.Text + "','" + TextBox3.Text + "','" + TextBox4.Text + "','" + TextBox5.Text + "')", baglanti)
            komut.ExecuteNonQueryAsync()
            baglanti.Close()
            MessageBox.Show("Kayıt Eklendi.", "Kayıt")
            temizle()
            tablo.Clear()
            listele()
        End If

        If TextBox1.Text = "Güncelle" Then
            baglanti.Open()
            Dim komut As New OleDbCommand("update bilgi set ad='" + TextBox1.Text + "', soyad='" + TextBox2.Text + "', sehir='" + TextBox3.Text + "', telefon='" + TextBox4.Text + "', email='" + TextBox5.Text + "'where telefon='" + TextBox4.Text + "'", baglanti)
            komut.ExecuteNonQueryAsync()
            baglanti.Close()
            MessageBox.Show("Kayıt Güncellendi.", "Kayıt")
            temizle()
            tablo.Clear()
            listele()
            Button1.Text = "Güncelle"
        End If
    End Sub

    Dim tablo As New DataTable
    Private Sub listele()
        'Listele Methodu
        baglanti.Open()
        Dim adtr As New OleDbDataAdapter("select *from bilgi", baglanti)
        adtr.Fill(tablo)
        DataGridView1.DataSource = tablo
        baglanti.Close()


    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        listele()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Silme Buton
        baglanti.Open()
        Dim komut As New OleDbCommand("delete *from bilgi where telefon='" + DataGridView1.CurrentRow.Cells("telefon").Value.ToString() + "'", baglanti)
            komut.ExecuteNonQueryAsync()
        baglanti.Close()
        tablo.Clear()
        listele()

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        'DataGrid Satır Bilgisi
        TextBox1.Text = DataGridView1.CurrentRow.Cells("ad").Value.ToString
        TextBox2.Text = DataGridView1.CurrentRow.Cells("soyad").Value.ToString
        TextBox3.Text = DataGridView1.CurrentRow.Cells("sehir").Value.ToString
        TextBox4.Text = DataGridView1.CurrentRow.Cells("telefon").Value.ToString
        TextBox5.Text = DataGridView1.CurrentRow.Cells("email").Value.ToString
        Button1.Text = "Güncelle"

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        'Arama Button
        baglanti.Open()
        Dim adtr1 As New OleDbDataAdapter("select *from bilgi where ad like'" + TextBox6.Text + "%'", baglanti)
        Dim tablo2 As New DataTable
        adtr1.Fill(tablo2)
        DataGridView1.DataSource = tablo2
        baglanti.Close()

    End Sub
End Class
