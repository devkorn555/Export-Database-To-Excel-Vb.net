Imports Microsoft.Office.Interop
Imports System.Data.SqlClient
Imports System.IO
Public Class Form1
    Dim wBook As Excel.Workbook '' Workbook คือไฟล์ Excel
    Dim wSheet As Excel.Worksheet '' Worksheet คือ Sheet ใน WorkBook
    Dim _excel As New Excel.Application ''ตัวแปร _excel ประกาศว่าให้คือ โปรแกรม Excel

    Dim dt As DataTable
    Dim con As New SqlConnection("Data Source=DESKTOP-MPTIJRG\SQLEXPRESS;Initial Catalog=C0FFEE;User ID=sa;PWD=12345")

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        '' โหดลข้อมูล จากฐานข้อมูล ปกติ
        con.Open()
        Dim txt As String = "SELECT * FROM TBCate"
        dt = New DataTable
        Dim da As New SqlDataAdapter(txt, con)
        da.Fill(dt)
        DataGridView1.DataSource = dt
    End Sub

    Private Sub ExportEx()
        wBook = _excel.Workbooks.Add '' สร้างไฟล์ Excel
        wSheet = wBook.ActiveSheet '' ให้ Sheet เริ่มต้นทำการเปิด

        Dim Hcindex As Integer = 1 'ตัวแปรนับจำนวน หัว Column เพิ่มเฉพาะหัว Column เท่านั้น
        'ทำการเพิ่ม Columns ลงใน Excel File
        For Each dataCol As DataColumn In dt.Columns ' อ่านจำนวน Columns ใน Datable หรือ ใน DatagridView ก่ได้ในที่นี้ใช้ Datable
            _excel.Cells(1, Hcindex) = dataCol.ColumnName ' ให้ไฟล์ Excel เพิ่ม หัว Column ใน Row ที่ 1 และ Column 1
            Hcindex += 1 ' บวก Column ทีละ 1 ให้เท่ากับจำนวน Column ใน Datatable
        Next


        Dim eXcelColindex As Integer = 1 ' ตัวแปรนับจำนวน Column เมื่อ เพิ่มข้อมูล
        Dim eXcelRow As Integer = 2 'ตัวแปรเริ่มต้นจำนวน Row เมื่อเพิ่มข้อมูล

        For Each col As DataColumn In dt.Columns ' อ่านข้อมูลจำนวน Column ใน Datatable
            For Each row As DataRow In dt.Rows   ' อ่านค่าจำนวน Row ที่อยู่ใน Datatable
                _excel.Cells(eXcelRow, eXcelColindex) = row.Item(col) ' _excel.cells(แถวที่เริ่มต้นเขียนลงใน Excel , Column ที่เริ่มต้นใน Excel) = แถวที่กำลังอ่าน(Column ที่ 0)
                eXcelRow += 1 'บวก Row ขึ้นที่ละ 1
            Next

            'เมื่อเขียนข้อมูล Column ที่ 1 Row ที่ 1 เสร็จแล้ว
            eXcelRow = 2 'ให้ Row กลับเป็นค่าเดิม คือเริ่มต้นจาก 2 ใหม่
            eXcelColindex += 1 'ให้ Column ขึ้น Column ใหม่ โดยการบวก +1 
        Next


        Dim pathFile As String = Application.StartupPath + "\textExport.xlsx"

        If (File.Exists(pathFile)) Then 'ให้เช็คว่าถ้ามี ไฟล์ อยู่แล้ว 

            File.Delete(pathFile) ''ให้ลบออกก่อน

            wBook.SaveAs(pathFile) '' สร้างไฟล์ใหม่
            wBook.Close()
            _excel.Quit()

            Process.Start(pathFile) ''ให้เปิดไฟล์ Excel เมื่อบันทึกเสร็จ
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ExportEx()
    End Sub
End Class
