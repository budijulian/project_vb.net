Public Class MenghitungNilaiApp


    Private Sub MenghitungNilaiApp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'memberi inputan ke dalam combobox matkul
        cb_jumlahM.Text = 1
        Dim a As Integer = 1
        For a = 1 To 20
            cb_jumlahM.Items.Add(a)
        Next
        'mengisi column pada datagridview
        DataGridView1.Columns.Add(0, "Nama Mata Kuliah")
        DataGridView1.Columns.Add(1, "Nilai Sikap")
        DataGridView1.Columns.Add(2, "Nilai Tugas")
        DataGridView1.Columns.Add(3, "Nilai UTS")
        DataGridView1.Columns.Add(4, "Nilai UAS")
        DataGridView1.Columns.Add(5, "Total")
        DataGridView1.Columns.Add(6, "Grade")
        ' memberi nilai lebar di datagridview
        DataGridView1.Columns(0).Width = 200
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 100
        DataGridView1.Columns(3).Width = 100
        DataGridView1.Columns(4).Width = 100
        DataGridView1.Columns(5).Width = 100
        DataGridView1.Columns(6).Width = 80


    End Sub
    Sub NilaiAkhir()
        'memberi nilai akhir seluruh matkul
        Dim y, z As Double
        Dim hurufAk As String
        For x As Integer = 0 To DataGridView1.RowCount - 1
            y = y + DataGridView1.Rows(x).Cells(5).Value
        Next
        txt_totalN.Text = y
        'mencari rata nilai akhir
        z = y / Val(cb_jumlahM.Text)

        If z <= 100 And z >= 80 Then
            hurufAk = "A+"
            txt_ipk.Text = "4.0"
            txt_pesan.Text = "Nilai Anda Sangat Memuaskan"
        ElseIf z < 80 And z >= 75 Then
            hurufAk = "A-"
            txt_ipk.Text = "4.0"
            txt_pesan.Text = "Nilai Anda Memuaskan"
        ElseIf z < 75 And z >= 70 Then
            hurufAk = "A"
            txt_ipk.Text = "3.7"
            txt_pesan.Text = "Nilai Anda Cukup Memuaskan"
        ElseIf z < 70 And z >= 65 Then
            hurufAk = "B+"
            txt_ipk.Text = "3.5"
            txt_pesan.Text = "Nilai Anda Lebih dari Cukup"
        ElseIf z < 65 And z >= 60 Then
            hurufAk = "B-"
            txt_ipk.Text = "3.0"
            txt_pesan.Text = "Nilai Anda "
        ElseIf z < 60 And z >= 55 Then
            hurufAk = "B"
            txt_ipk.Text = "2.7"
            txt_pesan.Text = "Sangat Memuaskan"
        Else
            hurufAk = "C"
            txt_ipk.Text = "2.5"
        End If
        txt_grade.Text = hurufAk
    End Sub
    Private Sub btn_pilih_Click(sender As Object, e As EventArgs) Handles btn_pilih.Click
        'enable is true textbox
        txt_grade.Enabled = True
        txt_Nsikap.Enabled = True
        txt_Ntugas.Enabled = True
        txt_Nuas.Enabled = True
        txt_Nuts.Enabled = True
        Dim b As Integer = 0
        lb_matkulke.Text = b
        If Val(lb_matkulke.Text) = Val(cb_jumlahM.Text) Then
        Else
            'memberikan informasi banyak matkul yang diambil
            lb_matkulke.Text = lb_matkulke.Text + 1
            lb_totalM.Text = String.Copy(cb_jumlahM.Text)
        End If
    End Sub
    Public total As Double
    Public huruf As String
    Public Function Hitung(ByVal sikap As Double, ByVal tugas As Double, ByVal uts As Double, ByVal uas As Double)
        'format   10% * Nilai Sikap + 20% * Nilai Tugas + 30% * Nilai UTS + 40% * Nilai UAS
        total = (sikap * 0.1) + (tugas * 0.2) + (uts * 0.3) + (uas * 0.4)

        'memasukan nilai total ke dalam datagridview masing - masing column 
        DataGridView1.RowCount = DataGridView1.RowCount + 1
        DataGridView1(0, DataGridView1.RowCount - 2).Value = txt_matkul.Text
        DataGridView1(1, DataGridView1.RowCount - 2).Value = sikap
        DataGridView1(2, DataGridView1.RowCount - 2).Value = tugas
        DataGridView1(3, DataGridView1.RowCount - 2).Value = uts
        DataGridView1(4, DataGridView1.RowCount - 2).Value = uas
        DataGridView1(5, DataGridView1.RowCount - 2).Value = total
        'memanggil grade hasil kalkulasi nilai matkul
        Call Grade()
        DataGridView1(6, DataGridView1.RowCount - 2).Value = huruf

        Return total
    End Function
    Sub Clear()
        txt_Nsikap.Text = Nothing
        txt_Ntugas.Text = Nothing
        txt_Nuts.Text = Nothing
        txt_Nuas.Text = Nothing
        txt_matkul.Text = ""
    End Sub
    Sub Grade()
        If total <= 100 And total >= 80 Then
            huruf = "A+"
        ElseIf total < 80 And total >= 75 Then
            huruf = "A-"
        ElseIf total < 75 And total >= 70 Then
            huruf = "A"
        ElseIf total < 70 And total >= 65 Then
            huruf = "B+"
        ElseIf total < 65 And total >= 60 Then
            huruf = "B-"
        ElseIf total < 60 And total >= 55 Then
            huruf = "B"
        Else
            huruf = "C"
        End If
    End Sub

    Private Sub btn_insert_Click(sender As Object, e As EventArgs) Handles btn_insert.Click
        If txt_Nsikap.Text = Nothing And txt_Ntugas.Text = Nothing And txt_Nuas.Text = Nothing And txt_Nuts.Text = Nothing Then
            MessageBox.Show("Maaf, Borang ada yang kosong", "Warning")
        Else
            ' kondisi dimana jika jumlah matkul maximum maka tidak bisa lagi diinputkan
            If cb_jumlahM.Text = Val(lb_matkulke.Text) Then

                'memanggil function hitung nilai
                Hitung(Val(txt_Nsikap.Text), Val(txt_Ntugas.Text), Val(txt_Nuts.Text), Val(txt_Nuas.Text))
                'memberi nilai naik di matkul -ke -
                lb_matkulke.Text = Val(lb_matkulke.Text)

                'meng -enabled texbok setelah jumlah matkul maximum
                txt_grade.Enabled = False
                txt_Nsikap.Enabled = False
                txt_Ntugas.Enabled = False
                txt_Nuas.Enabled = False
                txt_Nuts.Enabled = False

            Else
                'memanggil function hitung nilai
                Hitung(Val(txt_Nsikap.Text), Val(txt_Ntugas.Text), Val(txt_Nuts.Text), Val(txt_Nuas.Text))
                'memberi nilai naik di matkul -ke -
                lb_matkulke.Text = Val(lb_matkulke.Text) + 1
                Call Clear()
            End If
        End If
    End Sub

    Private Sub btn_clear_Click(sender As Object, e As EventArgs) Handles btn_clear.Click
        Call Clear()
    End Sub
    Private Sub btn_hitung_Click(sender As Object, e As EventArgs) Handles btn_hitung.Click
        Call NilaiAkhir()
        Call statusstrip()
    End Sub
    Sub statusstrip()
        ts_nama.Text = String.Copy(txt_nama.Text)
        ts_grade.Text = String.Copy(txt_grade.Text)
        ts_ipk.Text = String.Copy(txt_ipk.Text)
        ts_Nmatkul.Text = String.Copy(lb_totalM.Text)
        ts_npm.Text = String.Copy(txt_npm.Text)
    End Sub
    Private Sub btn_delete_Click(sender As Object, e As EventArgs) Handles btn_delete.Click
        'kodingan ini berguna untuk menghapus data baris di datagridview di 
        If DataGridView1.CurrentRow.Index <> DataGridView1.NewRowIndex Then
            DataGridView1.Rows.RemoveAt(DataGridView1.CurrentRow.Index)
        End If
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim i As Integer
        i = DataGridView1.SelectedCells(0).RowIndex
        txt_matkul.Text = DataGridView1.Rows(i).Cells(0).Value.ToString()
        txt_Nsikap.Text = DataGridView1.Rows(i).Cells(1).Value.ToString()
        txt_Ntugas.Text = DataGridView1.Rows(i).Cells(2).Value.ToString()
        txt_Nuts.Text = DataGridView1.Rows(i).Cells(3).Value.ToString()
        txt_Nuas.Text = DataGridView1.Rows(i).Cells(4).Value.ToString()
    End Sub

    Private Sub btn_clearcgv_Click(sender As Object, e As EventArgs) Handles btn_clearcgv.Click
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub ToolStripStatusLabel1_Click(sender As Object, e As EventArgs)

    End Sub
End Class