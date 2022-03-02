Attribute VB_Name = "Module1"
Option Private Module
Sub change_indikator_strategis()
Sheets("Risiko Strategis").Range("C6").value = "Indikator Inherent"
With Sheets("Risiko Strategis").Range("C27:C31")
    .value = "Ketidaksesuaian strategi bisnis dengan visi dan misi Perusahaan serta kondisi lingkungan usaha"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Strategis").Range("C32")
    .value = "Pilihan tingkat strategi bisnis yaitu strategi berisiko tinggi dan strategi berisiko rendah"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Strategis").Range("C33:C37")
    .value = "Posisi strategis (strategic position) Perusahaan di industri"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Strategis").Range("C38:C39")
    .value = "Pencapaian realisasi bisnis Perusahaan"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
End Sub
Sub change_indikator_operasional()
Sheets("Risiko Operasional").Range("C6").value = "Indikator Inherent"
With Sheets("Risiko Operasional").Range("C29:C34")
    .value = "Kompleksitas organisasi dan kegiatan usaha"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Operasional").Range("C35:C38")
    .value = "SDM"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Operasional").Range("C39:C44")
    .value = "Sistem teknologi informasi"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Operasional").Range("C45:C46")
    .value = "Risiko kecurangan (fraud)"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Operasional").Range("C47:C48")
    .value = "Gangguan terhadap bisnis dan organisasi"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Operasional").Range("C49:C50")
    .value = "Sistem administrasi"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
End Sub
Sub change_indikator_hukum()
Sheets("Risiko Hukum").Range("C6").value = "Indikator Inherent"
With Sheets("Risiko Hukum").Range("C25:C26")
    .value = "Ketiadaan atau perubahan peraturan perundangundangan"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Hukum").Range("C27:C31")
    .value = "Kelemahan dalam perikatan atau kerja sama"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Hukum").Range("C32:C38")
    .value = "Proses penyelesaian sengketa"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
End Sub
Sub change_indikator_kepatuhan()
Sheets("Risiko Kepatuhan").Range("C6").value = "Indikator Inherent"
With Sheets("Risiko Kepatuhan").Range("C25:C29")
    .value = "Jenis dan signifikansi pelanggaran yang dilakukan "
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Kepatuhan").Range("C30:C31")
    .value = "Frekuensi pelanggaraan (termasuk sanksi) yang dilakukan atau track record ketidakpatuhan perusahaan"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Kepatuhan").Range("C32")
    .value = "Pelanggaran terhadap ketentuan peraturan perundangundangan, ketentuan yang berlaku bagi Perusahaan, atau standar bisnis yang berlaku umum"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Kepatuhan").Range("C33")
    .value = "Tindak lanjut atas pelanggaran"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
End Sub
Sub change_indikator_reputasi()
Sheets("Risiko Reputasi").Range("C6").value = "Indikator Inherent"
With Sheets("Risiko Reputasi").Range("C28:C29")
    .value = "Pengaruh reputasi dari pengurus, pemilik, dan grup perusahaan"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Reputasi").Range("C30:C34")
    .value = "Pelanggaran etika bisnis"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Reputasi").Range("C35:C36")
    .value = ". Kompleksitas produk asuransi/ reasuransi yang diperantarai penempatannya atau dinilai kerugiannya dan kerja sama bisnis"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Reputasi").Range("C37:C38")
    .value = "Frekuensi, materialitas , dan eksposur pemberitaan negatif"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
With Sheets("Risiko Reputasi").Range("C39:C41")
    .value = "Frekuensi dan materialitas keluhan konsumen"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = True
End With
End Sub
Sub revert_indikator_strategis()
Sheets("Risiko Strategis").Range("C6").value = "Kategori"
With Sheets("Risiko Strategis").Range("C27:C39")
    .value = "Indikator Inherent"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = False
End With
End Sub
Sub revert_indikator_operasional()
Sheets("Risiko Operasional").Range("C6").value = "Kategori"
With Sheets("Risiko Operasional").Range("C29:C50")
    .value = "Indikator Inherent"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = False
End With
End Sub
Sub revert_indikator_hukum()
Sheets("Risiko Hukum").Range("C6").value = "Kategori"
With Sheets("Risiko Hukum").Range("C25:C38")
    .value = "Indikator Inherent"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = False
End With
End Sub
Sub revert_indikator_kepatuhan()
Sheets("Risiko Kepatuhan").Range("C6").value = "Kategori"
With Sheets("Risiko Kepatuhan").Range("C25:C33")
    .value = "Indikator Inherent"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = False
End With
End Sub
Sub revert_indikator_reputasi()
Sheets("Risiko Reputasi").Range("C6").value = "Kategori"
With Sheets("Risiko Reputasi").Range("C28:C41")
    .value = "Indikator Inherent"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = False
End With
End Sub
Sub save_as_pdf()
' Saves active sheet as PDF file.
    ThisWorkbook.Sheets("Report").Unprotect Password:="bdsb1sa"
    Dim Name As String
    Dim periodelapor As String
    Dim namadirektur As String
    Dim namakomisaris As String
    Dim shtreport As Worksheet
    Set shtreport = Sheets("Report")
    If shtreport.Range("D20") = "Tidak tersedia (data belum lengkap)" Or _
    shtreport.Range("D21") = "Tidak tersedia (data belum lengkap)" Or _
    shtreport.Range("D24") = "Tidak tersedia (data belum lengkap)" Or _
    shtreport.Range("D25") = "Tidak tersedia (data belum lengkap)" Or _
    shtreport.Range("D28") = "Tidak tersedia (data belum lengkap)" Or _
    shtreport.Range("D29") = "Tidak tersedia (data belum lengkap)" Or _
    shtreport.Range("D32") = "Tidak tersedia (data belum lengkap)" Or _
    shtreport.Range("D33") = "Tidak tersedia (data belum lengkap)" Or _
    shtreport.Range("D36") = "Tidak tersedia (data belum lengkap)" Or _
    shtreport.Range("D37") = "Tidak tersedia (data belum lengkap)" Or _
    shtreport.Range("C48") = "Peringkat tidak tersedia (data belum lengkap)" Or _
    shtreport.Range("C51") = "Tidak tersedia (data belum lengkap)" Then
    MsgBox "Mohon lengkapi data terlebih dahulu"
    Else
    'If Sheets
    periodelapor = _
    InputBox("Isi bulan atau periode pelaporan yang sudah dipilih:", "Laporan Risiko Komposit PT BDS", "Januari 2022/Q3 2021/Januari-April 2022")
    namadirektur = _
    InputBox("Isi nama direktur utama:", "Laporan Risiko Komposit PT BDS", "Nama lengkap direktur utama")
    namakomisaris = _
    InputBox("Isi nama komisaris utama:", "Laporan Risiko Komposit PT BDS", "Nama lengkap komisaris utama")
    If periodelapor = Isblank Or periodelapor = "Januari 2022/Q3 2021/Januari-April 2022" Then
    MsgBox "Isi periode pelaporan"
    ElseIf namadirektur = Isblank Or namadirektur = "Nama lengkap direktur utama" Then
    MsgBox "Isi nama direktur utama"
    ElseIf namakomisaris = Isblank Or namakomisaris = "Nama lengkap komisaris utama" Then
    MsgBox "Isi nama komisaris utama"
    Else
    shtreport.Range("D5") = ": " & periodelapor
    shtreport.Range("C75") = namadirektur
    With shtreport.Range("C75")
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    shtreport.Range("F75") = namakomisaris
    With shtreport.Range("F75")
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Name = ThisWorkbook.Path & "\" & shtreport.Name & " " & _
        Format(Now(), "ddmmyyyy hh.mm") & ".pdf"
    shtreport.ExportAsFixedFormat Type:=xlTypePDF, FileName:=Name, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    MsgBox "PDF successfully saved on " & ThisWorkbook.Path
    shtreport.Range("D5") = ": "
    shtreport.Range("C75") = ""
    shtreport.Range("F75") = ""
    End If
    End If
    Call protectworksheet
End Sub
Sub database()
' select sheet database
Call protectworksheet
    Call refresh_pivot
    Sheets("Database").Select
End Sub
Sub grafik()
' select sheet grafik
Call protectworksheet
    Call refresh_pivot
    Sheets("Graphs").Select
End Sub
Sub charts()
' select sheet charts
Call protectworksheet
    Call refresh_pivot
    Sheets("Chart").Select
End Sub
Sub report()
' select sheet report
Call protectworksheet
    Call refresh_pivot
    Sheets("Report").Select
End Sub
Sub risiko_komposit()
' select sheet risiko komposit
Call protectworksheet
    Call refresh_pivot
    Sheets("Risiko Komposit").Select
End Sub
Sub main_menu()
' select main menu
Call protectworksheet
    Call refresh_pivot
    Sheets("Main Menu").Select
End Sub
Sub risiko_strategis()
' select sheet risiko strategis
Call protectworksheet
    Call refresh_pivot
    Sheets("Risiko Strategis").Select
End Sub
Sub risiko_operasional()
' select sheet risiko operasional
Call protectworksheet
    Call refresh_pivot
    Sheets("Risiko Operasional").Select
End Sub
Sub risiko_hukum()
' select sheet risiko hukum
Call protectworksheet
    Call refresh_pivot
    Sheets("Risiko Hukum").Select
End Sub
Sub risiko_kepatuhan()
' select sheet risiko kepatuhan
Call protectworksheet
    Call refresh_pivot
    Sheets("Risiko Kepatuhan").Select
End Sub
Sub risiko_reputasi()
' select sheet risiko reputasi
Call protectworksheet
    Call refresh_pivot
    Sheets("Risiko Reputasi").Select
End Sub
Sub filter_kpmr_strategis()
' filter kpmr strategis
Call protectworksheet
With Sheets("Risiko Strategis")
    .ListObjects("RisikoStrategis").Range.AutoFilter Field:=10
    .ListObjects("RisikoStrategis").Range.AutoFilter Field:=1, _
        Criteria1:=Array("III.A.3.a", "III.A.3.b", "III.A.3.c", "III.A.3.d", "III.A.3.e", _
        "III.A.3.f", "III.A.3.g", "III.A.3.h", "III.A.3.i", "III.A.3.j", "III.A.3.k", _
        "III.A.3.l", "III.A.3.m", "III.A.3.n", "III.A.3.o"), Operator:=xlFilterValues
    .Range("E6").value = "1 (Kuat)"
    .Range("F6").value = "2 (Agak Kuat)"
    .Range("G6").value = "3 (Cukup)"
    .Range("H6").value = "4 (Agak Lemah)"
    .Range("I6").value = "5 (Lemah)"
End With
Call revert_indikator_strategis
End Sub
Sub filter_kpmr_operasional()
' filter kpmr operasional
Call protectworksheet
With Sheets("Risiko Operasional")
    .ListObjects("RisikoOperasional").Range.AutoFilter Field:=10
    .ListObjects("RisikoOperasional").Range.AutoFilter Field:=1, _
        Criteria1:=Array("III.B.3.a", "III.B.3.b", "III.B.3.c", "III.B.3.d", "III.B.3.e", _
        "III.B.3.f", "III.B.3.g", "III.B.3.h", "III.B.3.i", "III.B.3.j", "III.B.3.k", _
        "III.B.3.l", "III.B.3.m", "III.B.3.n", "III.B.3.o", "III.B.3.p"), Operator:= _
        xlFilterValues
    .Range("E6").value = "1 (Kuat)"
    .Range("F6").value = "2 (Agak Kuat)"
    .Range("G6").value = "3 (Cukup)"
    .Range("H6").value = "4 (Agak Lemah)"
    .Range("I6").value = "5 (Lemah)"
End With
Call revert_indikator_operasional
End Sub
Sub filter_kpmr_hukum()
' filter kpmr hukum
Call protectworksheet
With Sheets("Risiko Hukum")
    .ListObjects("RisikoHukum").Range.AutoFilter Field:=10
    .ListObjects("RisikoHukum").Range.AutoFilter Field:=1, Criteria1 _
        :=Array("III.C.3.a", "III.C.3.b", "III.C.3.c", "III.C.3.d", "III.C.3.e", _
        "III.C.3.f", "III.C.3.g", "III.C.3.h", "III.C.3.i", "III.C.3.j", "III.C.3.k", _
        "III.C.3.l", "III.C.3.m", "III.C.3.n", "III.C.3.o"), Operator:=xlFilterValues
    .Range("E6").value = "1 (Kuat)"
    .Range("F6").value = "2 (Agak Kuat)"
    .Range("G6").value = "3 (Cukup)"
    .Range("H6").value = "4 (Agak Lemah)"
    .Range("I6").value = "5 (Lemah)"
End With
Call revert_indikator_hukum
End Sub
Sub filter_kpmr_kepatuhan()
' filter kpmr kepatuhan
Call protectworksheet
With Sheets("Risiko Kepatuhan")
    .ListObjects("RisikoKepatuhan").Range.AutoFilter Field:=10
    .ListObjects("RisikoKepatuhan").Range.AutoFilter Field:=1, _
        Criteria1:=Array("III.D.3.a", "III.D.3.b", "III.D.3.c", "III.D.3.d", "III.D.3.e", _
        "III.D.3.f", "III.D.3.g", "III.D.3.h", "III.D.3.i", "III.D.3.j", "III.D.3.k", _
        "III.D.3.l", "III.D.3.m", "III.D.3.n", "III.D.3.o"), Operator:=xlFilterValues
    .Range("E6").value = "1 (Kuat)"
    .Range("F6").value = "2 (Agak Kuat)"
    .Range("G6").value = "3 (Cukup)"
    .Range("H6").value = "4 (Agak Lemah)"
    .Range("I6").value = "5 (Lemah)"
End With
Call revert_indikator_kepatuhan
End Sub
Sub filter_kpmr_reputasi()
' filter kpmr reputasi
Call protectworksheet
    Sheets("Risiko Reputasi").ListObjects("RisikoReputasi").Range.AutoFilter Field:=10
    Sheets("Risiko Reputasi").ListObjects("RisikoReputasi").Range.AutoFilter Field:=1, _
        Criteria1:=Array("III.E.3.a", "III.E.3.b", "III.E.3.c", "III.E.3.d", "III.E.3.e", _
        "III.E.3.f", "III.E.3.g", "III.E.3.h", "III.E.3.i", "III.E.3.j", "III.E.3.k", _
        "III.E.3.l", "III.E.3.m", "III.E.3.n", "III.E.3.o"), Operator:=xlFilterValues
    Sheets("Risiko Reputasi").Range("E6").value = "1 (Kuat)"
    Sheets("Risiko Reputasi").Range("F6").value = "2 (Agak Kuat)"
    Sheets("Risiko Reputasi").Range("G6").value = "3 (Cukup)"
    Sheets("Risiko Reputasi").Range("H6").value = "4 (Agak Lemah)"
    Sheets("Risiko Reputasi").Range("I6").value = "5 (Lemah)"
Call revert_indikator_reputasi
End Sub
Sub filter_inherent_indikator_strategis()
' filter indikator strategis
Call protectworksheet
Call change_indikator_strategis
    Sheets("Risiko Strategis").ListObjects("RisikoStrategis").Range.AutoFilter Field:=10
    Sheets("Risiko Strategis").ListObjects("RisikoStrategis").Range.AutoFilter Field:=1, _
        Criteria1:=Array("III.A.1.1.a", "III.A.1.1.b", "III.A.1.1.c", "III.A.1.1.d", _
        "III.A.1.1.e", "III.A.1.2.a", "III.A.1.3.a", "III.A.1.3.b", "III.A.1.3.c", _
        "III.A.1.3.d", "III.A.1.3.e", "III.A.1.4.a", "III.A.1.4.b"), Operator:= _
        xlFilterValues
    Sheets("Risiko Strategis").Range("E6").value = "1 (Rendah)"
    Sheets("Risiko Strategis").Range("F6").value = "2 (Sedang Rendah)"
    Sheets("Risiko Strategis").Range("G6").value = "3 (Sedang)"
    Sheets("Risiko Strategis").Range("H6").value = "4 (Sedang Tinggi)"
    Sheets("Risiko Strategis").Range("I6").value = "5 (Tinggi)"
End Sub
Sub filter_inherent_indikator_operasional()
' filter indikator operasional
Call protectworksheet
Call change_indikator_operasional
    Sheets("Risiko Operasional").ListObjects("RisikoOperasional").Range.AutoFilter Field:=10
    Sheets("Risiko Operasional").ListObjects("RisikoOperasional").Range.AutoFilter Field:=1, _
        Criteria1:=Array("III.B.1.1.a", "III.B.1.1.b", "III.B.1.1.c", "III.B.1.1.d", _
        "III.B.1.1.e", "III.B.1.1.f", "III.B.1.2.a", "III.B.1.2.b", "III.B.1.2.c", _
        "III.B.1.2.d", "III.B.1.3.a", "III.B.1.3.b", "III.B.1.3.c", "III.B.1.3.d", _
        "III.B.1.3.e", "III.B.1.3.f", "III.B.1.4.a", "III.B.1.4.b", "III.B.1.5.a", _
        "III.B.1.5.b", "III.B.1.6.a", "III.B.1.6.b"), Operator:=xlFilterValues
    Sheets("Risiko Operasional").Range("E6").value = "1 (Rendah)"
    Sheets("Risiko Operasional").Range("F6").value = "2 (Sedang Rendah)"
    Sheets("Risiko Operasional").Range("G6").value = "3 (Sedang)"
    Sheets("Risiko Operasional").Range("H6").value = "4 (Sedang Tinggi)"
    Sheets("Risiko Operasional").Range("I6").value = "5 (Tinggi)"
End Sub
Sub filter_inherent_indikator_hukum()
' filter indikator hukum
Call protectworksheet
Call change_indikator_hukum
    Sheets("Risiko Hukum").ListObjects("RisikoHukum").Range.AutoFilter Field:=10
    Sheets("Risiko Hukum").ListObjects("RisikoHukum").Range.AutoFilter Field:=1, Criteria1 _
        :=Array("III.C.1.1.a", "III.C.1.1.b", "III.C.1.2.a", "III.C.1.2.b", "III.C.1.2.c", _
        "III.C.1.2.d", "III.C.1.2.e", "III.C.1.3.a", "III.C.1.3.b", "III.C.1.3.c", _
        "III.C.1.3.d", "III.C.1.3.e", "III.C.1.3.f", "III.C.1.3.g"), Operator:= _
        xlFilterValues
    Sheets("Risiko Hukum").Range("E6").value = "1 (Rendah)"
    Sheets("Risiko Hukum").Range("F6").value = "2 (Sedang Rendah)"
    Sheets("Risiko Hukum").Range("G6").value = "3 (Sedang)"
    Sheets("Risiko Hukum").Range("H6").value = "4 (Sedang Tinggi)"
    Sheets("Risiko Hukum").Range("I6").value = "5 (Tinggi)"
End Sub
Sub filter_inherent_indikator_kepatuhan()
' filter indikator kepatuhan
Call protectworksheet
Call change_indikator_kepatuhan
    Sheets("Risiko Kepatuhan").ListObjects("RisikoKepatuhan").Range.AutoFilter Field:=10
    Sheets("Risiko Kepatuhan").ListObjects("RisikoKepatuhan").Range.AutoFilter Field:=1, _
        Criteria1:=Array("III.D.1.1.a", "III.D.1.1.b", "III.D.1.1.c", "III.D.1.1.d", _
        "III.D.1.1.e", "III.D.1.2.a", "III.D.1.2.b", "III.D.1.3.a", "III.D.1.4.a"), _
        Operator:=xlFilterValues
    Sheets("Risiko Kepatuhan").Range("E6").value = "1 (Rendah)"
    Sheets("Risiko Kepatuhan").Range("F6").value = "2 (Sedang Rendah)"
    Sheets("Risiko Kepatuhan").Range("G6").value = "3 (Sedang)"
    Sheets("Risiko Kepatuhan").Range("H6").value = "4 (Sedang Tinggi)"
    Sheets("Risiko Kepatuhan").Range("I6").value = "5 (Tinggi)"
End Sub
Sub filter_inherent_indikator_reputasi()
' filter indikator reputasi
Call protectworksheet
Call change_indikator_reputasi
    Sheets("Risiko Reputasi").ListObjects("RisikoReputasi").Range.AutoFilter Field:=10
    Sheets("Risiko Reputasi").ListObjects("RisikoReputasi").Range.AutoFilter Field:=1, _
        Criteria1:=Array("III.E.1.1.a", "III.E.1.1.b", "III.E.1.2.a", "III.E.1.2.b", _
        "III.E.1.2.c", "III.E.1.2.d", "III.E.1.2.e", "III.E.1.3.a", "III.E.1.3.b", _
        "III.E.1.4.a", "III.E.1.4.b", "III.E.1.5.a", "III.E.1.5.b", "III.E.1.5.c"), _
        Operator:=xlFilterValues
    Sheets("Risiko Reputasi").Range("E6").value = "1 (Rendah)"
    Sheets("Risiko Reputasi").Range("F6").value = "2 (Sedang Rendah)"
    Sheets("Risiko Reputasi").Range("G6").value = "3 (Sedang)"
    Sheets("Risiko Reputasi").Range("H6").value = "4 (Sedang Tinggi)"
    Sheets("Risiko Reputasi").Range("I6").value = "5 (Tinggi)"
End Sub
Sub filter_inherent_peringkat_strategis()
' filter peringkat strategis
Call protectworksheet
    Sheets("Risiko Strategis").ListObjects("RisikoStrategis").Range.AutoFilter Field:=10
    Sheets("Risiko Strategis").ListObjects("RisikoStrategis").Range.AutoFilter Field:=1, _
        Criteria1:=Array("III.A.2.a", "III.A.2.b", "III.A.2.c", "III.A.2.d", "III.A.2.e") _
        , Operator:=xlFilterValues
    Sheets("Risiko Strategis").Range("E6").value = "1 (Rendah)"
    Sheets("Risiko Strategis").Range("F6").value = "2 (Sedang Rendah)"
    Sheets("Risiko Strategis").Range("G6").value = "3 (Sedang)"
    Sheets("Risiko Strategis").Range("H6").value = "4 (Sedang Tinggi)"
    Sheets("Risiko Strategis").Range("I6").value = "5 (Tinggi)"
    Call revert_indikator_strategis
End Sub
Sub filter_inherent_peringkat_operasional()
' filter peringkat operasional
Call protectworksheet
    Sheets("Risiko Operasional").ListObjects("RisikoOperasional").Range.AutoFilter Field:=10
    Sheets("Risiko Operasional").ListObjects("RisikoOperasional").Range.AutoFilter Field:=1, _
        Criteria1:=Array("III.B.2.a", "III.B.2.b", "III.B.2.c", "III.B.2.d", "III.B.2.e", _
        "III.B.2.f"), Operator:=xlFilterValues
    Sheets("Risiko Operasional").Range("E6").value = "1 (Rendah)"
    Sheets("Risiko Operasional").Range("F6").value = "2 (Sedang Rendah)"
    Sheets("Risiko Operasional").Range("G6").value = "3 (Sedang)"
    Sheets("Risiko Operasional").Range("H6").value = "4 (Sedang Tinggi)"
    Sheets("Risiko Operasional").Range("I6").value = "5 (Tinggi)"
Call revert_indikator_operasional
End Sub
Sub filter_inherent_peringkat_hukum()
' filter peringkat hukum
Call protectworksheet
    Sheets("Risiko Hukum").ListObjects("RisikoHukum").Range.AutoFilter Field:=10
    Sheets("Risiko Hukum").ListObjects("RisikoHukum").Range.AutoFilter Field:=1, Criteria1 _
        :=Array("III.C.2.a", "III.C.2.b", "III.C.2.c"), Operator:=xlFilterValues
    Sheets("Risiko Hukum").Range("E6").value = "1 (Rendah)"
    Sheets("Risiko Hukum").Range("F6").value = "2 (Sedang Rendah)"
    Sheets("Risiko Hukum").Range("G6").value = "3 (Sedang)"
    Sheets("Risiko Hukum").Range("H6").value = "4 (Sedang Tinggi)"
    Sheets("Risiko Hukum").Range("I6").value = "5 (Tinggi)"
Call revert_indikator_hukum
End Sub
Sub filter_inherent_peringkat_kepatuhan()
' filter peringkat kepatuhan
Call protectworksheet
    Sheets("Risiko Kepatuhan").ListObjects("RisikoKepatuhan").Range.AutoFilter Field:=10
    Sheets("Risiko Kepatuhan").ListObjects("RisikoKepatuhan").Range.AutoFilter Field:=1, _
        Criteria1:=Array("III.D.2.a", "III.D.2.b", "III.D.2.c"), Operator:= _
        xlFilterValues
    Sheets("Risiko Kepatuhan").Range("E6").value = "1 (Rendah)"
    Sheets("Risiko Kepatuhan").Range("F6").value = "2 (Sedang Rendah)"
    Sheets("Risiko Kepatuhan").Range("G6").value = "3 (Sedang)"
    Sheets("Risiko Kepatuhan").Range("H6").value = "4 (Sedang Tinggi)"
    Sheets("Risiko Kepatuhan").Range("I6").value = "5 (Tinggi)"
Call revert_indikator_kepatuhan
End Sub
Sub filter_inherent_peringkat_reputasi()
' filter peringkat reputasi
Call protectworksheet
    Sheets("Risiko Reputasi").ListObjects("RisikoReputasi").Range.AutoFilter Field:=10
    Sheets("Risiko Reputasi").ListObjects("RisikoReputasi").Range.AutoFilter Field:=1, _
        Criteria1:=Array("III.E.2.a", "III.E.2.b", "III.E.2.c", "III.E.2.d", "III.E.2.e", _
        "III.E.2.f"), Operator:=xlFilterValues
    Sheets("Risiko Reputasi").Range("E6").value = "1 (Rendah)"
    Sheets("Risiko Reputasi").Range("F6").value = "2 (Sedang Rendah)"
    Sheets("Risiko Reputasi").Range("G6").value = "3 (Sedang)"
    Sheets("Risiko Reputasi").Range("H6").value = "4 (Sedang Tinggi)"
    Sheets("Risiko Reputasi").Range("I6").value = "5 (Tinggi)"
Call revert_indikator_reputasi
End Sub
Sub clear_filter_kategori_strategis()
' clear filter kategori strategis
Call protectworksheet
    Sheets("Risiko Strategis").ListObjects("RisikoStrategis").Range.AutoFilter Field:=10
    Sheets("Risiko Strategis").ListObjects("RisikoStrategis").Range.AutoFilter Field:=1
    Sheets("Risiko Strategis").Range("E6").value = "1 (Kuat/Rendah)"
    Sheets("Risiko Strategis").Range("F6").value = "2 (Agak Kuat/Sedang Rendah)"
    Sheets("Risiko Strategis").Range("G6").value = "3 (Cukup/Sedang)"
    Sheets("Risiko Strategis").Range("H6").value = "4 (Agak Lemah/Sedang Tinggi)"
    Sheets("Risiko Strategis").Range("I6").value = "5 (Lemah/Tinggi)"
    Call revert_indikator_strategis
End Sub
Sub clear_filter_kategori_operasional()
' clear filter kategori operasional
Call protectworksheet
    Sheets("Risiko Operasional").ListObjects("RisikoOperasional").Range.AutoFilter Field:=10
    Sheets("Risiko Operasional").ListObjects("RisikoOperasional").Range.AutoFilter Field:=1
    Sheets("Risiko Operasional").Range("E6").value = "1 (Kuat/Rendah)"
    Sheets("Risiko Operasional").Range("F6").value = "2 (Agak Kuat/Sedang Rendah)"
    Sheets("Risiko Operasional").Range("G6").value = "3 (Cukup/Sedang)"
    Sheets("Risiko Operasional").Range("H6").value = "4 (Agak Lemah/Sedang Tinggi)"
    Sheets("Risiko Operasional").Range("I6").value = "5 (Lemah/Tinggi)"
    Call revert_indikator_operasional
End Sub
Sub clear_filter_kategori_hukum()
' clear filter kategori hukum
Call protectworksheet
    Sheets("Risiko Hukum").ListObjects("RisikoHukum").Range.AutoFilter Field:=10
    Sheets("Risiko Hukum").ListObjects("RisikoHukum").Range.AutoFilter Field:=1
    Sheets("Risiko Hukum").Range("E6").value = "1 (Kuat/Rendah)"
    Sheets("Risiko Hukum").Range("F6").value = "2 (Agak Kuat/Sedang Rendah)"
    Sheets("Risiko Hukum").Range("G6").value = "3 (Cukup/Sedang)"
    Sheets("Risiko Hukum").Range("H6").value = "4 (Agak Lemah/Sedang Tinggi)"
    Sheets("Risiko Hukum").Range("I6").value = "5 (Lemah/Tinggi)"
End Sub
Sub clear_filter_kategori_kepatuhan()
' clear filter kategori kepatuhan
Call protectworksheet
    Sheets("Risiko Kepatuhan").ListObjects("RisikoKepatuhan").Range.AutoFilter Field:=10
    Sheets("Risiko Kepatuhan").ListObjects("RisikoKepatuhan").Range.AutoFilter Field:=1
    Sheets("Risiko Kepatuhan").Range("E6").value = "1 (Kuat/Rendah)"
    Sheets("Risiko Kepatuhan").Range("F6").value = "2 (Agak Kuat/Sedang Rendah)"
    Sheets("Risiko Kepatuhan").Range("G6").value = "3 (Cukup/Sedang)"
    Sheets("Risiko Kepatuhan").Range("H6").value = "4 (Agak Lemah/Sedang Tinggi)"
    Sheets("Risiko Kepatuhan").Range("I6").value = "5 (Lemah/Tinggi)"
Call revert_indikator_kepatuhan
End Sub
Sub clear_filter_kategori_reputasi()
' clear filter kategori reputasi
Call protectworksheet
    Sheets("Risiko Reputasi").ListObjects("RisikoReputasi").Range.AutoFilter Field:=10
    Sheets("Risiko Reputasi").ListObjects("RisikoReputasi").Range.AutoFilter Field:=1
    Sheets("Risiko Reputasi").Range("E6").value = "1 (Kuat/Rendah)"
    Sheets("Risiko Reputasi").Range("F6").value = "2 (Agak Kuat/Sedang Rendah)"
    Sheets("Risiko Reputasi").Range("G6").value = "3 (Cukup/Sedang)"
    Sheets("Risiko Reputasi").Range("H6").value = "4 (Agak Lemah/Sedang Tinggi)"
    Sheets("Risiko Reputasi").Range("I6").value = "5 (Lemah/Tinggi)"
Call revert_indikator_reputasi
End Sub
Sub filter_blank_data_strategis()
' filter blank data strategis
Call protectworksheet
    Sheets("Risiko Strategis").ListObjects("RisikoStrategis").Range.AutoFilter Field:=1
    Sheets("Risiko Strategis").ListObjects("RisikoStrategis").Range.AutoFilter Field:=10, Criteria1:="="
    Sheets("Risiko Strategis").Range("E6").value = "1 (Kuat/Rendah)"
    Sheets("Risiko Strategis").Range("F6").value = "2 (Agak Kuat/Sedang Rendah)"
    Sheets("Risiko Strategis").Range("G6").value = "3 (Cukup/Sedang)"
    Sheets("Risiko Strategis").Range("H6").value = "4 (Agak Lemah/Sedang Tinggi)"
    Sheets("Risiko Strategis").Range("I6").value = "5 (Lemah/Tinggi)"
    Call revert_indikator_strategis
End Sub
Sub filter_blank_data_operasional()
' filter blank data operasional
Call protectworksheet
    Sheets("Risiko Operasional").ListObjects("RisikoOperasional").Range.AutoFilter Field:=1
    Sheets("Risiko Operasional").ListObjects("RisikoOperasional").Range.AutoFilter Field:=10, Criteria1:="="
    Sheets("Risiko Operasional").Range("E6").value = "1 (Kuat/Rendah)"
    Sheets("Risiko Operasional").Range("F6").value = "2 (Agak Kuat/Sedang Rendah)"
    Sheets("Risiko Operasional").Range("G6").value = "3 (Cukup/Sedang)"
    Sheets("Risiko Operasional").Range("H6").value = "4 (Agak Lemah/Sedang Tinggi)"
    Sheets("Risiko Operasional").Range("I6").value = "5 (Lemah/Tinggi)"
    Call revert_indikator_operasional
End Sub
Sub filter_blank_data_hukum()
' filter blank data hukum
Call protectworksheet
    Sheets("Risiko Hukum").ListObjects("RisikoHukum").Range.AutoFilter Field:=1
    Sheets("Risiko Hukum").ListObjects("RisikoHukum").Range.AutoFilter Field:=10, Criteria1:="="
    Sheets("Risiko Hukum").Range("E6").value = "1 (Kuat/Rendah)"
    Sheets("Risiko Hukum").Range("F6").value = "2 (Agak Kuat/Sedang Rendah)"
    Sheets("Risiko Hukum").Range("G6").value = "3 (Cukup/Sedang)"
    Sheets("Risiko Hukum").Range("H6").value = "4 (Agak Lemah/Sedang Tinggi)"
    Sheets("Risiko Hukum").Range("I6").value = "5 (Lemah/Tinggi)"
End Sub
Sub filter_blank_data_kepatuhan()
' filter blank data kepatuhan
Call protectworksheet
    Sheets("Risiko Kepatuhan").ListObjects("RisikoKepatuhan").Range.AutoFilter Field:=1
    Sheets("Risiko Kepatuhan").ListObjects("RisikoKepatuhan").Range.AutoFilter Field:=10, Criteria1:="="
    Sheets("Risiko Kepatuhan").Range("E6").value = "1 (Kuat/Rendah)"
    Sheets("Risiko Kepatuhan").Range("F6").value = "2 (Agak Kuat/Sedang Rendah)"
    Sheets("Risiko Kepatuhan").Range("G6").value = "3 (Cukup/Sedang)"
    Sheets("Risiko Kepatuhan").Range("H6").value = "4 (Agak Lemah/Sedang Tinggi)"
    Sheets("Risiko Kepatuhan").Range("I6").value = "5 (Lemah/Tinggi)"
Call revert_indikator_kepatuhan
End Sub
Sub filter_blank_data_reputasi()
' filter blank data reputasi
Call protectworksheet
    Sheets("Risiko Reputasi").ListObjects("RisikoReputasi").Range.AutoFilter Field:=1
    Sheets("Risiko Reputasi").ListObjects("RisikoReputasi").Range.AutoFilter Field:=10, Criteria1:="="
    Sheets("Risiko Reputasi").Range("E6").value = "1 (Kuat/Rendah)"
    Sheets("Risiko Reputasi").Range("F6").value = "2 (Agak Kuat/Sedang Rendah)"
    Sheets("Risiko Reputasi").Range("G6").value = "3 (Cukup/Sedang)"
    Sheets("Risiko Reputasi").Range("H6").value = "4 (Agak Lemah/Sedang Tinggi)"
    Sheets("Risiko Reputasi").Range("I6").value = "5 (Lemah/Tinggi)"
Call revert_indikator_reputasi
End Sub
Sub get_data_strategis()
' get data strategis
Call protectworksheet
    Call clear_filter_kategori_strategis
    MsgBox "Selalu pilih tanggal satu pada bulan yang akan dipilih"
    Dim dateVariable As Date
    dateVariable = CalendarForm.GetDate
    Range("E2") = dateVariable
    Range("E2").NumberFormat = "dd/mm/yyyy"
    Sheets("Risiko Strategis").Range("K7:K" & Cells(Rows.Count, "A").End(xlUp).Row).SpecialCells(xlCellTypeVisible).FormulaR1C1 = Range("E2")
    Sheets("Risiko Strategis").Range("J7:J39").FormulaR1C1 = "=VLOOKUP(RC[-9]&(TEXT(RC[1],""dd/mm/yyyy"")),Database!C[-9]:C[1],11,0)"
    Sheets("Risiko Strategis").Range("J7:J39").Copy
    Sheets("Risiko Strategis").Range("J7:J39").PasteSpecial Paste:=xlPasteValues
    For Each Rng In Range("J7:J" & Cells(Rows.Count, "A").End(xlUp).Row)
    If Rng.Text = "#N/A" Then
    Rng.ClearContents
    End If
    If Rng.Value2 = "0" Then
    Rng.ClearContents
    End If
    Next
    Range("E2").NumberFormat = "MMM-YYYY"
End Sub
Sub get_data_operasional()
' get data operasional
Call protectworksheet
    Call clear_filter_kategori_operasional
    MsgBox "Selalu pilih tanggal satu pada bulan yang akan dipilih"
    Dim dateVariable As Date
    dateVariable = CalendarForm.GetDate
    Range("E2") = dateVariable
    Range("E2").NumberFormat = "dd/mm/yyyy"
    Sheets("Risiko Operasional").Range("K7:K" & Cells(Rows.Count, "A").End(xlUp).Row).SpecialCells(xlCellTypeVisible).FormulaR1C1 = Range("E2")
    Sheets("Risiko Operasional").Range("J7:J50").FormulaR1C1 = "=VLOOKUP(RC[-9]&(TEXT(RC[1],""dd/mm/yyyy"")),Database!C[-9]:C[1],11,0)"
    Sheets("Risiko Operasional").Range("J7:J50").Copy
    Sheets("Risiko Operasional").Range("J7:J50").PasteSpecial Paste:=xlPasteValues
    For Each Rng In Range("J7:J" & Cells(Rows.Count, "A").End(xlUp).Row)
    If Rng.Text = "#N/A" Then
    Rng.ClearContents
    End If
    If Rng.Value2 = "0" Then
    Rng.ClearContents
    End If
    Next
    Range("E2").NumberFormat = "MMM-YYYY"
End Sub
Sub get_data_hukum()
' get data hukum
Call protectworksheet
    Call clear_filter_kategori_hukum
    MsgBox "Selalu pilih tanggal satu pada bulan yang akan dipilih"
    Dim dateVariable As Date
    dateVariable = CalendarForm.GetDate
    Range("E2") = dateVariable
    Range("E2").NumberFormat = "dd/mm/yyyy"
    Sheets("Risiko Hukum").Range("K7:K" & Cells(Rows.Count, "A").End(xlUp).Row).SpecialCells(xlCellTypeVisible).FormulaR1C1 = Range("E2")
    Sheets("Risiko Hukum").Range("J7:J38").FormulaR1C1 = "=VLOOKUP(RC[-9]&(TEXT(RC[1],""dd/mm/yyyy"")),Database!C[-9]:C[1],11,0)"
    Sheets("Risiko Hukum").Range("J7:J38").Copy
    Sheets("Risiko Hukum").Range("J7:J38").PasteSpecial Paste:=xlPasteValues
    For Each Rng In Range("J7:J" & Cells(Rows.Count, "A").End(xlUp).Row)
    If Rng.Text = "#N/A" Then
    Rng.ClearContents
    End If
    If Rng.Value2 = "0" Then
    Rng.ClearContents
    End If
    Next
    Range("E2").NumberFormat = "MMM-YYYY"
End Sub
Sub get_data_kepatuhan()
' get data kepatuhan
Call protectworksheet
    Call clear_filter_kategori_kepatuhan
    MsgBox "Selalu pilih tanggal satu pada bulan yang akan dipilih"
    Dim dateVariable As Date
    dateVariable = CalendarForm.GetDate
    Range("E2") = dateVariable
    Range("E2").NumberFormat = "dd/mm/yyyy"
    Sheets("Risiko Kepatuhan").Range("K7:K" & Cells(Rows.Count, "A").End(xlUp).Row).SpecialCells(xlCellTypeVisible).FormulaR1C1 = Range("E2")
    Sheets("Risiko Kepatuhan").Range("J7:J33").FormulaR1C1 = "=VLOOKUP(RC[-9]&(TEXT(RC[1],""dd/mm/yyyy"")),Database!C[-9]:C[1],11,0)"
    Sheets("Risiko Kepatuhan").Range("J7:J33").Copy
    Sheets("Risiko Kepatuhan").Range("J7:J33").PasteSpecial Paste:=xlPasteValues
    For Each Rng In Range("J7:J" & Cells(Rows.Count, "A").End(xlUp).Row)
    If Rng.Text = "#N/A" Then
    Rng.ClearContents
    End If
    If Rng.Value2 = "0" Then
    Rng.ClearContents
    End If
    Next
    Range("E2").NumberFormat = "MMM-YYYY"
End Sub
Sub get_data_reputasi()
' get data reputasi
Call protectworksheet
    Call clear_filter_kategori_reputasi
    MsgBox "Selalu pilih tanggal satu pada bulan yang akan dipilih"
    Dim dateVariable As Date
    dateVariable = CalendarForm.GetDate
    Range("E2") = dateVariable
    Range("E2").NumberFormat = "dd/mm/yyyy"
    Sheets("Risiko Reputasi").Range("K7:K" & Cells(Rows.Count, "A").End(xlUp).Row).SpecialCells(xlCellTypeVisible).FormulaR1C1 = Range("E2")
    Sheets("Risiko Reputasi").Range("J7:J41").FormulaR1C1 = "=VLOOKUP(RC[-9]&(TEXT(RC[1],""dd/mm/yyyy"")),Database!C[-9]:C[1],11,0)"
    Sheets("Risiko Reputasi").Range("J7:J41").Copy
    Sheets("Risiko Reputasi").Range("J7:J41").PasteSpecial Paste:=xlPasteValues
    For Each Rng In Range("J7:J" & Cells(Rows.Count, "A").End(xlUp).Row)
    If Rng.Text = "#N/A" Then
    Rng.ClearContents
    End If
    If Rng.Value2 = "0" Then
    Rng.ClearContents
    End If
    Next
    Range("E2").NumberFormat = "MMM-YYYY"
End Sub
Sub submit_data_strategis()
' submit data strategis ke database
Call protectworksheet
    Application.ScreenUpdating = False
    Call revert_indikator_strategis
    Call remove_duplicate
    Call clear_blank_database
        answer = MsgBox("Jika pernah submit data sebelumnya akan terhapus, anda yakin?", vbQuestion + vbYesNo + vbDefaultButton2, "Submit data")
    If answer = vbYes And Sheets("Risiko Strategis").Range("I4") = "Data komplit" Then
            Call copy_strategis_database
    ElseIf answer = vbYes And Sheets("Risiko Strategis").Range("I4") = "Data belum komplit" Then
        answer = MsgBox("Data belum komplit, anda yakin?", vbQuestion + vbYesNo + vbDefaultButton2, "Data belum komplit")
        If answer = vbYes Then
            Call copy_strategis_database
        Else
        Sheets("Risiko Strategis").Select
        MsgBox ("Periksa Kembali data yang akan disubmit")
        End If
    ElseIf answer = vbNo Then
        Sheets("Risiko Strategis").Select
        MsgBox ("Periksa Kembali data yang akan disubmit")
    End If
    Application.ScreenUpdating = True
    Call refresh_pivot
End Sub
Sub copy_strategis_database()
Call clear_filter_database
        Sheets("Database").ListObjects("Table2").Range.AutoFilter Field:=3, Criteria1:= _
             "=Inherent Strategis", Operator:=xlOr, Criteria2:="=KPMR Strategis"
        Sheets("Risiko Strategis").Unprotect Password:="bdsb1sa"
        Sheets("Risiko Strategis").Range("E2").NumberFormat = "dd/mm/yyyy"
        Sheets("Database").ListObjects("Table2").Range.AutoFilter Field:=12, Criteria1:=Worksheets("Risiko Strategis").Range("E2").Text
        On Error Resume Next
        Application.DisplayAlerts = False
        Sheets("Database").ListObjects("Table2").DataBodyRange.SpecialCells(xlCellTypeVisible).delete
        Application.DisplayAlerts = True
        On Error GoTo 0
        Call clear_filter_database
        Call clear_filter_kategori_strategis
        Sheets("Database").Unprotect Password:="bdsb1sa"
        Sheets("Risiko Strategis").ListObjects("RisikoStrategis").DataBodyRange.Copy
        If Sheets("Database").ListObjects("Table2").DataBodyRange Is Nothing Then
        Sheets("Database").ListObjects("Table2").Range(Sheets("Database").ListObjects("Table2").Range.Rows.Count, "A").Offset(0, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        ElseIf Not Sheets("Database").ListObjects("Table2").DataBodyRange Is Nothing Then
        Sheets("Database").ListObjects("Table2").Range(Sheets("Database").ListObjects("Table2").Range.Rows.Count, "A").Offset(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        'Sheets("Database").ListObjects("Table2").DataBodyRange(1, 2).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        End If
        Sheets("Risiko Strategis").Range("E2").NumberFormat = "MMM-YYYY"
        Sheets("Risiko Strategis").Select
        MsgBox ("Submit data berhasil")
        Call key_database
End Sub
Sub submit_data_operasional()
' submit data operasional ke database
Call protectworksheet
    Application.ScreenUpdating = False
    Call remove_duplicate
    Call clear_blank_database
    Call revert_indikator_operasional
        answer = MsgBox("Jika pernah submit data sebelumnya akan terhapus, anda yakin?", vbQuestion + vbYesNo + vbDefaultButton2, "Submit data")
    If answer = vbYes And Sheets("Risiko Operasional").Range("I4") = "Data komplit" Then
        Call copy_operasional_database
   ElseIf answer = vbYes And Sheets("Risiko Operasional").Range("I4") = "Data belum komplit" Then
        answer = MsgBox("Data belum komplit, anda yakin?", vbQuestion + vbYesNo + vbDefaultButton2, "Data belum komplit")
        If answer = vbYes Then
            Call copy_operasional_database
        Else
        Sheets("Risiko Operasional").Select
        MsgBox ("Periksa Kembali data yang akan disubmit")
        End If
    ElseIf answer = vbNo Then
        Sheets("Risiko Operasional").Select
        MsgBox ("Periksa Kembali data yang akan disubmit")
    End If
    Application.ScreenUpdating = True
    Call refresh_pivot
End Sub
Sub copy_operasional_database()
Call clear_filter_database
        Sheets("Database").ListObjects("Table2").Range.AutoFilter Field:=3, Criteria1:= _
             "=Inherent Operasional", Operator:=xlOr, Criteria2:="=KPMR Operasional"
        Sheets("Risiko Operasional").Unprotect Password:="bdsb1sa"
        Sheets("Risiko Operasional").Range("E2").NumberFormat = "dd/mm/yyyy"
        Sheets("Database").ListObjects("Table2").Range.AutoFilter Field:=12, Criteria1:=Worksheets("Risiko Operasional").Range("E2").Text
        On Error Resume Next
        Application.DisplayAlerts = False
        Sheets("Database").ListObjects("Table2").DataBodyRange.SpecialCells(xlCellTypeVisible).delete
        Application.DisplayAlerts = True
        On Error GoTo 0
        Call clear_filter_database
        Call clear_filter_kategori_operasional
        Sheets("Database").Unprotect Password:="bdsb1sa"
        Sheets("Risiko Operasional").ListObjects("RisikoOperasional").DataBodyRange.Copy
        If Sheets("Database").ListObjects("Table2").DataBodyRange Is Nothing Then
        Sheets("Database").ListObjects("Table2").Range(Sheets("Database").ListObjects("Table2").Range.Rows.Count, "A").Offset(0, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        ElseIf Not Sheets("Database").ListObjects("Table2").DataBodyRange Is Nothing Then
        Sheets("Database").ListObjects("Table2").Range(Sheets("Database").ListObjects("Table2").Range.Rows.Count, "A").Offset(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        'Sheets("Database").ListObjects("Table2").DataBodyRange(1, 2).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        End If
        Sheets("Risiko Operasional").Range("E2").NumberFormat = "MMM-YYYY"
        Sheets("Risiko Operasional").Select
        MsgBox ("Submit data berhasil")
        Call key_database
End Sub
Sub submit_data_hukum()
' submit data hukum ke database
Call protectworksheet
    Application.ScreenUpdating = False
    Call remove_duplicate
    Call clear_blank_database
    Call revert_indikator_hukum
        answer = MsgBox("Jika pernah submit data sebelumnya akan terhapus, anda yakin?", vbQuestion + vbYesNo + vbDefaultButton2, "Submit data")
    If answer = vbYes And Sheets("Risiko Hukum").Range("I4") = "Data komplit" Then
        Call copy_hukum_database
    ElseIf answer = vbYes And Sheets("Risiko Hukum").Range("I4") = "Data belum komplit" Then
        answer = MsgBox("Data belum komplit, anda yakin?", vbQuestion + vbYesNo + vbDefaultButton2, "Data belum komplit")
        If answer = vbYes Then
            Call copy_hukum_database
        Else
        Sheets("Risiko Hukum").Select
        MsgBox ("Periksa Kembali data yang akan disubmit")
        End If
    ElseIf answer = vbNo Then
        Sheets("Risiko Hukum").Select
        MsgBox ("Periksa Kembali data yang akan disubmit")
    End If
    Application.ScreenUpdating = True
    Call refresh_pivot
End Sub
Sub copy_hukum_database()
Call clear_filter_database
        Sheets("Database").ListObjects("Table2").Range.AutoFilter Field:=3, Criteria1:= _
             "=Inherent Hukum", Operator:=xlOr, Criteria2:="=KPMR Hukum"
        Sheets("Risiko Hukum").Unprotect Password:="bdsb1sa"
        Sheets("Risiko Hukum").Range("E2").NumberFormat = "dd/mm/yyyy"
        Sheets("Database").ListObjects("Table2").Range.AutoFilter Field:=12, Criteria1:=Worksheets("Risiko Hukum").Range("E2").Text
        On Error Resume Next
        Application.DisplayAlerts = False
        Sheets("Database").ListObjects("Table2").DataBodyRange.SpecialCells(xlCellTypeVisible).delete
        Application.DisplayAlerts = True
        On Error GoTo 0
        Call clear_filter_database
        Call clear_filter_kategori_hukum
        Sheets("Database").Unprotect Password:="bdsb1sa"
        Sheets("Risiko Hukum").ListObjects("RisikoHukum").DataBodyRange.Copy
        If Sheets("Database").ListObjects("Table2").DataBodyRange Is Nothing Then
        Sheets("Database").ListObjects("Table2").Range(Sheets("Database").ListObjects("Table2").Range.Rows.Count, "A").Offset(0, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        ElseIf Not Sheets("Database").ListObjects("Table2").DataBodyRange Is Nothing Then
        Sheets("Database").ListObjects("Table2").Range(Sheets("Database").ListObjects("Table2").Range.Rows.Count, "A").Offset(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        'Sheets("Database").ListObjects("Table2").DataBodyRange(1, 2).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        End If
        Sheets("Risiko Hukum").Range("E2").NumberFormat = "MMM-YYYY"
        Sheets("Risiko Hukum").Select
        MsgBox ("Submit data berhasil")
        Call key_database
End Sub
Sub submit_data_kepatuhan()
' submit data kepatuhan ke database
Call protectworksheet
    Application.ScreenUpdating = False
    Call remove_duplicate
    Call clear_blank_database
    Call revert_indikator_kepatuhan
        answer = MsgBox("Jika pernah submit data sebelumnya akan terhapus, anda yakin?", vbQuestion + vbYesNo + vbDefaultButton2, "Submit data")
    If answer = vbYes And Sheets("Risiko Kepatuhan").Range("I4") = "Data komplit" Then
        Call copy_kepatuhan_database
    ElseIf answer = vbYes And Sheets("Risiko Kepatuhan").Range("I4") = "Data belum komplit" Then
        answer = MsgBox("Data belum komplit, anda yakin?", vbQuestion + vbYesNo + vbDefaultButton2, "Data belum komplit")
        If answer = vbYes Then
            Call copy_kepatuhan_database
        Else
        Sheets("Risiko Kepatuhan").Select
        MsgBox ("Periksa Kembali data yang akan disubmit")
        End If
    ElseIf answer = vbNo Then
        Sheets("Risiko Kepatuhan").Select
        MsgBox ("Periksa Kembali data yang akan disubmit")
    End If
    Application.ScreenUpdating = True
    Call refresh_pivot
End Sub
Sub copy_kepatuhan_database()
Call clear_filter_database
        Sheets("Database").ListObjects("Table2").Range.AutoFilter Field:=3, Criteria1:= _
             "=Inherent Kepatuhan", Operator:=xlOr, Criteria2:="=KPMR Kepatuhan"
        Sheets("Risiko Kepatuhan").Unprotect Password:="bdsb1sa"
        Sheets("Risiko Kepatuhan").Range("E2").NumberFormat = "dd/mm/yyyy"
        Sheets("Database").ListObjects("Table2").Range.AutoFilter Field:=12, Criteria1:=Worksheets("Risiko Kepatuhan").Range("E2").Text
        On Error Resume Next
        Application.DisplayAlerts = False
        Sheets("Database").ListObjects("Table2").DataBodyRange.SpecialCells(xlCellTypeVisible).delete
        Application.DisplayAlerts = True
        On Error GoTo 0
        Call clear_filter_database
        Call clear_filter_kategori_kepatuhan
        Sheets("Database").Unprotect Password:="bdsb1sa"
        Sheets("Risiko Kepatuhan").ListObjects("RisikoKepatuhan").DataBodyRange.Copy
        If Sheets("Database").ListObjects("Table2").DataBodyRange Is Nothing Then
        Sheets("Database").ListObjects("Table2").Range(Sheets("Database").ListObjects("Table2").Range.Rows.Count, "A").Offset(0, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        ElseIf Not Sheets("Database").ListObjects("Table2").DataBodyRange Is Nothing Then
        Sheets("Database").ListObjects("Table2").Range(Sheets("Database").ListObjects("Table2").Range.Rows.Count, "A").Offset(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        'Sheets("Database").ListObjects("Table2").DataBodyRange(1, 2).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        End If
        Sheets("Risiko Kepatuhan").Range("E2").NumberFormat = "MMM-YYYY"
        Sheets("Risiko Kepatuhan").Select
        MsgBox ("Submit data berhasil")
        Call key_database
End Sub
Sub submit_data_reputasi()
' submit data reputasi ke database
Call protectworksheet
    Application.ScreenUpdating = False
    Call remove_duplicate
    Call clear_blank_database
    Call revert_indikator_reputasi
        answer = MsgBox("Jika pernah submit data sebelumnya akan terhapus, anda yakin?", vbQuestion + vbYesNo + vbDefaultButton2, "Submit data")
    If answer = vbYes And Sheets("Risiko Reputasi").Range("I4") = "Data komplit" Then
        Call copy_reputasi_database
    ElseIf answer = vbYes And Sheets("Risiko Reputasi").Range("I4") = "Data belum komplit" Then
        answer = MsgBox("Data belum komplit, anda yakin?", vbQuestion + vbYesNo + vbDefaultButton2, "Data belum komplit")
        If answer = vbYes Then
            Call copy_reputasi_database
        Else
        Sheets("Risiko Reputasi").Select
        MsgBox ("Periksa Kembali data yang akan disubmit")
        End If
    ElseIf answer = vbNo Then
        Sheets("Risiko Reputasi").Select
        MsgBox ("Periksa Kembali data yang akan disubmit")
    End If
    Application.ScreenUpdating = True
    Call refresh_pivot
End Sub
Sub copy_reputasi_database()
Call clear_filter_database
        Sheets("Database").ListObjects("Table2").Range.AutoFilter Field:=3, Criteria1:= _
             "=Inherent Reputasi", Operator:=xlOr, Criteria2:="=KPMR Reputasi"
        Sheets("Risiko Reputasi").Unprotect Password:="bdsb1sa"
        Sheets("Risiko Reputasi").Range("E2").NumberFormat = "dd/mm/yyyy"
        Sheets("Database").ListObjects("Table2").Range.AutoFilter Field:=12, Criteria1:=Worksheets("Risiko Reputasi").Range("E2").Text
        On Error Resume Next
        Application.DisplayAlerts = False
        Sheets("Database").ListObjects("Table2").DataBodyRange.SpecialCells(xlCellTypeVisible).delete
        Application.DisplayAlerts = True
        On Error GoTo 0
        Call clear_filter_database
        Call clear_filter_kategori_reputasi
        Sheets("Database").Unprotect Password:="bdsb1sa"
        Sheets("Risiko Reputasi").ListObjects("RisikoReputasi").DataBodyRange.Copy
        If Sheets("Database").ListObjects("Table2").DataBodyRange Is Nothing Then
        Sheets("Database").ListObjects("Table2").Range(Sheets("Database").ListObjects("Table2").Range.Rows.Count, "A").Offset(0, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        ElseIf Not Sheets("Database").ListObjects("Table2").DataBodyRange Is Nothing Then
        Sheets("Database").ListObjects("Table2").Range(Sheets("Database").ListObjects("Table2").Range.Rows.Count, "A").Offset(1, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        'Sheets("Database").ListObjects("Table2").DataBodyRange(1, 2).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        End If
        Sheets("Risiko Reputasi").Range("E2").NumberFormat = "MMM-YYYY"
        Sheets("Risiko Reputasi").Select
        MsgBox ("Submit data berhasil")
        Call key_database
End Sub
Sub remove_duplicate()
    Sheets("Database").Select
        Sheets("Database").ListObjects("Table2").Range.AutoFilter Field:=1
        Sheets("Database").ListObjects("Table2").Range.AutoFilter Field:=3
        Sheets("Database").ListObjects("Table2").Range.AutoFilter Field:=12
        Sheets("Database").ListObjects("Table2").Range.RemoveDuplicates Columns:=1, Header:=xlYes
End Sub
Sub clear_blank_database()
    With Sheets("Database").ListObjects("Table2")
        .Range.AutoFilter Field:=1
        .Range.AutoFilter Field:=3
        .Range.AutoFilter Field:=12
        .Range.AutoFilter Field:=11, Criteria1:="="
        On Error Resume Next
        .DataBodyRange.SpecialCells(xlCellTypeVisible).EntireRow.delete
        On Error GoTo 0
        .Range.AutoFilter Field:=11
    End With
    Call key_database
End Sub
Sub refresh_pivot()
   Dim Ws As Worksheet
   Dim Pword As String
   Pword = "bdsb1sa"
   Application.ScreenUpdating = False
   Application.EnableEvents = False
   For Each Ws In Sheets(Array("Report", "Risiko Komposit", "Chart", "Graphs", "Database"))
      Ws.Unprotect Pword
   Next Ws
   Application.ScreenUpdating = True
Application.EnableEvents = False
  ThisWorkbook.RefreshAll
Application.EnableEvents = True
Call protectworksheet
End Sub
Sub protectworksheet()
'Protect, even if already protected
Dim pw As String
    pw = "bdsb1sa"
        ThisWorkbook.Sheets("Report").Protect Password:=pw, _
        DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowUsingPivotTables:=True
        ThisWorkbook.Sheets("Risiko Strategis").Protect Password:=pw, _
        UserInterfaceOnly:=True
        ThisWorkbook.Sheets("Risiko Operasional").Protect Password:=pw, _
        UserInterfaceOnly:=True
        ThisWorkbook.Sheets("Risiko Hukum").Protect Password:=pw, _
        UserInterfaceOnly:=True
        ThisWorkbook.Sheets("Risiko Kepatuhan").Protect Password:=pw, _
        UserInterfaceOnly:=True
        ThisWorkbook.Sheets("Risiko Reputasi").Protect Password:=pw, _
        UserInterfaceOnly:=True
        ThisWorkbook.Sheets("Database").Protect Password:=pw, _
        UserInterfaceOnly:=True
        ThisWorkbook.Sheets("Chart").Protect Password:=pw, _
        DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowUsingPivotTables:=True
        ThisWorkbook.Sheets("Graphs").Protect Password:=pw, _
        DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowUsingPivotTables:=True
        ThisWorkbook.Sheets("Risiko Komposit").Protect Password:=pw, _
        DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowUsingPivotTables:=True
End Sub
Sub key_database()
'Bikin key di database
On Error Resume Next
With Sheets("Database").ListObjects("Table2")
    .ListColumns("Key").DataBodyRange.FormulaR1C1 = "=[@No]&(TEXT([@Periode],""dd/mm/yyyy""))"
    .ListColumns("Period").DataBodyRange.FormulaR1C1 = "=TEXT([@Periode],""Mmm-YYYY"")"
    .ListColumns("Peringkat").DataBodyRange.NumberFormat = "General"
    .ListColumns("Periode").DataBodyRange.NumberFormat = "dd/mm/yyyy"
End With
On Error GoTo 0
End Sub
Sub clear_filter_database()
    With Sheets("Database").ListObjects("Table2").Range
        .AutoFilter Field:=1
        .AutoFilter Field:=3
        .AutoFilter Field:=11
        .AutoFilter Field:=12
    End With
End Sub
