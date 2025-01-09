Imports MySql.Data.MySqlClient

Public Class Form1

    Private counter As Integer = 1
    Private isDarkMode As Boolean = False

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Dim conn As MySqlConnection
        Dim COMMAND As MySqlCommand
        conn = New MySqlConnection
        conn.ConnectionString = "server=localhost;userid=root;password='';database=data_anggota"

        Try
            conn.Open()
            MessageBox.Show("Connection to MySQL test database was successful!!!!", "TESTING      CONNECTION TO MySQL DATABASE")
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            conn.Close()
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set KeyPreview to true
        Me.KeyPreview = True

        ' Refresh data grid view saat form dimuat
        RefreshDataGridView()

        TextBox1.Text = counter.ToString()
        TextBox1.ReadOnly = True

        ' Isi data ke dalam combobox
        FillComboBoxData()

        ' Deteksi waktu lokal dan aktifkan mode gelap jika diperlukan
        SetupDarkMode()
    End Sub

    Private Sub SetupDarkMode()
        Dim localTime As DateTime = DateTime.Now
        Dim centralIndonesiaTime As DateTime = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(localTime, "SE Asia Standard Time")

        ' Mode gelap otomatis diaktifkan antara pukul 18:00 dan 06:00
        If centralIndonesiaTime.Hour >= 18 OrElse centralIndonesiaTime.Hour < 6 Then
            EnableDarkMode()
        Else
            EnableLightMode()
        End If
    End Sub

    Private Sub EnableDarkMode()
        BackColor = Color.FromArgb(40, 40, 40)
        For Each ctrl As Control In Controls
            ctrl.BackColor = Color.FromArgb(40, 40, 40)
            ctrl.ForeColor = Color.White
        Next
        Button5.Text = "Light Mode"
        isDarkMode = True
    End Sub

    Private Sub EnableLightMode()
        BackColor = SystemColors.Control
        For Each ctrl As Control In Controls
            ctrl.BackColor = SystemColors.Window
            ctrl.ForeColor = SystemColors.ControlText
        Next
        Button5.Text = "Dark Mode"
        isDarkMode = False
    End Sub

    ' Menangani event KeyDown untuk menambahkan shortcut keyboard
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        ' Menangani shortcut keyboard
        If e.Control AndAlso e.KeyCode = Keys.S Then
            e.SuppressKeyPress = True ' Mencegah default behavior (Select All)
            Button1.PerformClick() ' Ctrl + S untuk penambahan data
        ElseIf e.Control AndAlso e.KeyCode = Keys.O Then
            e.SuppressKeyPress = True ' Mencegah default behavior
            Button2.PerformClick() ' Ctrl + O untuk Keluar
        ElseIf e.Control AndAlso e.KeyCode = Keys.C Then
            e.SuppressKeyPress = True ' Mencegah default behavior
            Button3.PerformClick() ' Ctrl + C untuk pengosongan kolom
        ElseIf e.Control AndAlso e.KeyCode = Keys.D Then
            e.SuppressKeyPress = True ' Mencegah default behavior
            Button5.PerformClick() ' Ctrl + D untuk mengganti Tema
        End If
    End Sub

    Private Sub FillComboBoxData()
        ComboBox1.Items.Add("XI PPLG 1")
        ComboBox1.Items.Add("XI RPL 1")
        ComboBox1.Items.Add("XII TKJ 2")

        ' Atur combobox agar hanya menampilkan opsi saat diklik
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ' Kode untuk menghandle event saat combobox dipilih
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        ' Hanya menerima input angka, maksimal 13 digit
        Dim input As String = TextBox7.Text
        If input.Length > 13 Then
            ' Potong input menjadi 13 digit
            TextBox7.Text = input.Substring(0, 13)
            ' Tempatkan kursor di akhir teks
            TextBox7.SelectionStart = TextBox7.Text.Length
        ElseIf Not IsNumeric(input) Then
            ' Hapus karakter non-angka
            TextBox7.Text = input.Replace(input.Last, "")
            ' Tempatkan kursor di akhir teks
            TextBox7.SelectionStart = TextBox7.Text.Length
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Validasi input
        If String.IsNullOrEmpty(TextBox1.Text) Or String.IsNullOrEmpty(TextBox2.Text) Or String.IsNullOrEmpty(ComboBox1.Text) Or String.IsNullOrEmpty(TextBox4.Text) Or String.IsNullOrEmpty(TextBox5.Text) Or String.IsNullOrEmpty(TextBox6.Text) Or String.IsNullOrEmpty(TextBox7.Text) Then
            MessageBox.Show("Semua kolom harus diisi.")
            Return
        End If

        ' Validasi nomor HP
        Dim noHP As String = TextBox7.Text
        If noHP.Length <> 13 OrElse Not IsNumeric(noHP) Then
            MessageBox.Show("Nomor HP harus terdiri dari 13 angka.")
            Return
        End If

        ' Validasi apakah nomor anggota sudah digunakan
        Dim noAnggota As String = TextBox1.Text
        If IsNoAnggotaUsed(noAnggota) Then
            MessageBox.Show("Nomor anggota sudah digunakan.")
            Return
        End If

        ' Konfirmasi penambahan data
        Dim result As DialogResult = MessageBox.Show("Apakah Kamu Ingin Menambahkan Data Ini?", "Konfirmasi Penambahan", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.No Then
            Return
        End If

        ' Simpan data ke database
        Dim conn As New MySqlConnection("server=localhost;userid=root;password='';database=data_anggota")
        Dim query As String = "INSERT INTO anggota (no_anggota, nama, kelas, alamat, jurusan, email, no_hp) VALUES (@no_anggota, @nama, @kelas, @alamat, @jurusan, @email, @no_hp)"
        Dim cmd As New MySqlCommand(query, conn)

        ' Ambil nilai dari kontrol input
        cmd.Parameters.AddWithValue("@no_anggota", TextBox1.Text)
        cmd.Parameters.AddWithValue("@nama", TextBox2.Text)
        cmd.Parameters.AddWithValue("@kelas", ComboBox1.SelectedItem)
        cmd.Parameters.AddWithValue("@alamat", TextBox4.Text)
        cmd.Parameters.AddWithValue("@jurusan", TextBox5.Text)
        cmd.Parameters.AddWithValue("@email", TextBox6.Text)

        ' Sensor nomor HP dengan bintang
        Dim sensoredNoHP As String = New String("*"c, noHP.Length)
        cmd.Parameters.AddWithValue("@no_hp", sensoredNoHP)

        Try
            conn.Open()
            cmd.ExecuteNonQuery()
            MessageBox.Show("Data berhasil disimpan.")

            ' Refresh data grid view
            RefreshDataGridView()
        Catch ex As Exception
            MessageBox.Show("Gagal menyimpan data: " & ex.Message)
        Finally
            conn.Close()
        End Try

        counter += 1
        TextBox1.Text = counter.ToString()
    End Sub

    Private Sub RefreshDataGridView()
        Dim conn As New MySqlConnection("server=localhost;userid=root;password='';database=data_anggota")
        Dim adapter As New MySqlDataAdapter("SELECT * FROM anggota", conn)
        Dim table As New DataTable()
        adapter.Fill(table)
        DataGridView1.DataSource = table
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Konfirmasi keluar dari aplikasi
        Dim result As DialogResult = MessageBox.Show("Anda yakin ingin keluar dari aplikasi?", "Konfirmasi Keluar", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Me.Close()
        End If
    End Sub

    Private Function IsValidEmail(email As String) As Boolean
        Dim emailRegex As New System.Text.RegularExpressions.Regex("^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$")
        Return emailRegex.IsMatch(email)
    End Function

    Private Function IsNoAnggotaUsed(noAnggota As String) As Boolean
        Dim conn As New MySqlConnection("server=localhost;userid=root;password='';database=data_anggota")
        Dim query As String = "SELECT COUNT(*) FROM anggota WHERE no_anggota = @no_anggota"
        Dim cmd As New MySqlCommand(query, conn)
        cmd.Parameters.AddWithValue("@no_anggota", noAnggota)

        Try
            conn.Open()
            Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
            Return count > 0
        Catch ex As Exception
            MessageBox.Show("Gagal memeriksa nomor anggota: " & ex.Message)
            Return True
        Finally
            conn.Close()
        End Try
    End Function

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        ' Kosongkan semua kolom kecuali kolom No Anggota
        TextBox2.Text = ""
        ComboBox1.SelectedIndex = -1
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Periksa mode saat ini
        If Me.BackColor = SystemColors.Control Then
            ' Mode terang saat ini, ubah ke mode gelap
            Me.BackColor = Color.FromArgb(35, 35, 35)
            Me.ForeColor = Color.White
            Button1.BackColor = Color.FromArgb(64, 64, 64)
            Button2.BackColor = Color.FromArgb(64, 64, 64)
            Button3.BackColor = Color.FromArgb(64, 64, 64)
            Button4.BackColor = Color.FromArgb(64, 64, 64)
            Button5.BackColor = Color.FromArgb(64, 64, 64)
            TextBox1.BackColor = Color.FromArgb(64, 64, 64)
            TextBox2.BackColor = Color.FromArgb(64, 64, 64)
            TextBox4.BackColor = Color.FromArgb(64, 64, 64)
            TextBox5.BackColor = Color.FromArgb(64, 64, 64)
            TextBox6.BackColor = Color.FromArgb(64, 64, 64)
            TextBox7.BackColor = Color.FromArgb(64, 64, 64)
            ComboBox1.BackColor = Color.FromArgb(64, 64, 64)
            ComboBox1.ForeColor = Color.White
            DataGridView1.BackgroundColor = Color.FromArgb(64, 64, 64)
            DataGridView1.ForeColor = Color.White
        Else
            ' Mode gelap saat ini, ubah ke mode terang
            Me.BackColor = SystemColors.Control
            Me.ForeColor = SystemColors.ControlText
            Button1.BackColor = SystemColors.Control
            Button2.BackColor = SystemColors.Control
            Button3.BackColor = SystemColors.Control
            Button4.BackColor = SystemColors.Control
            Button5.BackColor = SystemColors.Control
            TextBox1.BackColor = SystemColors.Window
            TextBox2.BackColor = SystemColors.Window
            TextBox4.BackColor = SystemColors.Window
            TextBox5.BackColor = SystemColors.Window
            TextBox6.BackColor = SystemColors.Window
            TextBox7.BackColor = SystemColors.Window
            ComboBox1.BackColor = SystemColors.Window
            ComboBox1.ForeColor = SystemColors.ControlText
            DataGridView1.BackgroundColor = SystemColors.Window
            DataGridView1.ForeColor = SystemColors.ControlText
        End If

        If isDarkMode Then
            EnableLightMode()
        Else
            EnableDarkMode()
        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim helpMessage As String = "Panduan Penggunaan Aplikasi:" & Environment.NewLine & Environment.NewLine &
                                 "1. Menambahkan Data:" & Environment.NewLine &
                                 "   - Isi semua kolom yang tersedia." & Environment.NewLine &
                                 "   - Klik tombol 'Simpan' untuk menambahkan data." & Environment.NewLine & Environment.NewLine &
                                 "2. Menghapus Data:" & Environment.NewLine &
                                 "   - Klik tombol 'Bersih' untuk mengosongkan kolom." & Environment.NewLine &
                                 "   - Pilih data yang ingin dihapus dari tabel." & Environment.NewLine &
                                 "   - Data akan dihapus saat disimpan." & Environment.NewLine & Environment.NewLine &
                                 "3. Memperbarui Data:" & Environment.NewLine &
                                 "   - Ubah informasi di kolom yang diperlukan." & Environment.NewLine &
                                 "   - Klik tombol 'Simpan' untuk menyimpan perubahan." & Environment.NewLine & Environment.NewLine &
                                 "Shortcut Keyboard:" & Environment.NewLine &
                                 "   - Ctrl + S: Menambahkan data" & Environment.NewLine &
                                 "   - Ctrl + O: Keluar" & Environment.NewLine &
                                 "   - Ctrl + C: Mengosongkan kolom" & Environment.NewLine &
                                 "   - Ctrl + D: Mengganti tema" & Environment.NewLine &
                                 "   - Ctrl + H: Menampilkan panduan ini" & Environment.NewLine & Environment.NewLine &
                                 "©Copyright All Reserved 2023 By Alpian. Semua hak dilindungi." & Environment.NewLine

        MessageBox.Show(helpMessage, "Panduan Penggunaan", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim feedbackFAQMessage As String = "Feedback dan FAQ:" & Environment.NewLine & Environment.NewLine &
                                        "1. Apa itu aplikasi ini?" & Environment.NewLine &
                                        "   - Aplikasi ini digunakan untuk mengelola data anggota." & Environment.NewLine & Environment.NewLine &
                                        "2. Bagaimana cara menambahkan data?" & Environment.NewLine &
                                        "   - Isi semua kolom yang tersedia dan klik tombol 'Simpan'." & Environment.NewLine & Environment.NewLine &
                                        "3. Apakah saya bisa menghapus data?" & Environment.NewLine &
                                        "   - Ya, pilih data yang ingin dihapus dan klik tombol 'Simpan'." & Environment.NewLine & Environment.NewLine &
                                        "4. Di mana saya bisa melihat statistik?" & Environment.NewLine &
                                        "   - Statistik dapat dilihat pada bagian laporan di aplikasi." & Environment.NewLine & Environment.NewLine &
                                        "5. Bagaimana cara menghubungi dukungan?" & Environment.NewLine &
                                        "   - Silakan kirim email ke support@contoh.com." & Environment.NewLine & Environment.NewLine &
                                        "Feedback:" & Environment.NewLine &
                                        "   - Kami sangat menghargai masukan Anda. Silakan beri tahu kami jika ada saran untuk perbaikan." & Environment.NewLine

        MessageBox.Show(feedbackFAQMessage, "Feedback/FAQ", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class