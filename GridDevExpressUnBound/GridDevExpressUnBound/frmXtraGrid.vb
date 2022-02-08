#Region "ABOUT"
' / --------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gsoft.com/
' /
' / Purpose: Sample code to used XtraGrid DevExpress.
' / Microsoft Visual Basic .NET (2010) + MS Access 2007+
' /
' / This is open source code under @Copyleft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------
#End Region

Imports System.Data.OleDb
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.Data
Imports DevExpress.XtraGrid

Public Class frmXtraGrid
    '// หากเป็นโปรเจคจริงๆ กลุ่มตัวแปรเหล่านี้ต้องนำไปวางไว้ใน Module 
    Dim Conn As OleDb.OleDbConnection
    'Dim DA As New System.Data.OleDb.OleDbDataAdapter()
    Dim Cmd As New System.Data.OleDb.OleDbCommand
    Dim strSQL As String

    '// Connect MS Access DataBase
    Function ConnectDataBase() As OleDb.OleDbConnection
        Return New OleDb.OleDbConnection( _
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
            MyPath(Application.StartupPath) & "dbFood.accdb;Persist Security Info=True")
    End Function

    ' / --------------------------------------------------------------------------------
    ' / Get my project path
    Function MyPath(ByVal AppPath As String) As String
        '/ MessageBox.Show(AppPath);
        MyPath = AppPath.ToLower.Replace("\bin\debug", "\").Replace("\bin\release", "\").Replace("\bin\x86\debug", "\")
        '/ Return Value
        '// If not found folder then put the \ (BackSlash ASCII Code = 92) at the end.
        If Microsoft.VisualBasic.Right(MyPath, 1) <> Chr(92) Then MyPath = MyPath & Chr(92)
    End Function

    '// Start Here.
    Private Sub frmXtraGrid_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        DevExpress.Utils.AppearanceObject.DefaultFont = New Font("Tahoma", 10, FontStyle.Regular, GraphicsUnit.Point, 0)
        Conn = ConnectDataBase()
        '// Anchor @Runtime
        Me.GridControl1.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
    End Sub

    Private Sub SetupGridView()
        GridView1.Columns.Clear()
        '// ตั้งค่าคุณสมบัติ XtraGrid
        '// Start Add Fields.
        Dim GC As New GridColumn    '// Imports DevExpress.XtraGrid.Columns
        GC = GridView1.Columns.AddField("FoodPK")
        With GC
            .Caption = "FoodPK"
            .UnboundType = DevExpress.Data.UnboundColumnType.Integer
            '// ซ่อนหลัก Index = 0 ซึ่งเป็นค่า Primary Key 
            '// เมื่อผู้ใช้กดดับเบิ้ลคลิ๊กเมาส์ หรือกด Enter ในแต่ละแถว เราจะนำค่านี้ไป Query เพื่อแสดงผลรายละเอียดอีกที
            .Visible = False
        End With
        '//
        GC = GridView1.Columns.AddField("FoodID")
        With GC
            .Caption = "รหัสอาหาร"
            .UnboundType = DevExpress.Data.UnboundColumnType.String
            .Visible = True
        End With
        '//
        GC = GridView1.Columns.AddField("FoodName")
        With GC
            .Caption = "ชื่ออาหาร"
            .UnboundType = DevExpress.Data.UnboundColumnType.String
            .Visible = True
        End With
        '//
        GC = GridView1.Columns.AddField("CategoryName")
        With GC
            .Caption = "ประเภท"
            .UnboundType = DevExpress.Data.UnboundColumnType.String
            .Visible = True
        End With
        '//
        GC = GridView1.Columns.AddField("PriceCash")
        With GC
            .Caption = "ราคา"
            .UnboundType = DevExpress.Data.UnboundColumnType.Decimal
            .DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
            .DisplayFormat.FormatString = "#,##.00;[#.0];0.00"
            .AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            .AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            .Visible = True
        End With
        '//
        GC = GridView1.Columns.AddField("PictureFood")
        With GC
            .Caption = "รูปภาพ"
            .UnboundType = DevExpress.Data.UnboundColumnType.Object
            .AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
            .Visible = True
        End With

        '// Setup XtraGrid Properties @Run Time.
        Dim view As DevExpress.XtraGrid.Views.Grid.GridView = CType(GridControl1.MainView, DevExpress.XtraGrid.Views.Grid.GridView)
        DevExpress.XtraEditors.WindowsFormsSettings.DefaultFont = New Font("Tahoma", 10)
        With view
            .OptionsBehavior.AutoPopulateColumns = False
            '// ไม่ให้เกิด GroupPanel เพื่อจับหลักลากมาวางในส่วนนี้ได้ จะใช้การเขียนโค้ดในการจัดกลุ่มแทน
            .OptionsView.ShowGroupPanel = False
            .OptionsCustomization.AllowFilter = True '//False
            .OptionsCustomization.AllowColumnMoving = False
            '// CheckBox Selector
            .OptionsSelection.ResetSelectionClickOutsideCheckboxSelector = True
        End With
        '//
        With GridView1
            .RowHeight = 25
            .GroupRowHeight = 30
            .Appearance.Row.Font = New Font("Tahoma", 10, FontStyle.Regular)
            .Appearance.HeaderPanel.Font = New Font("Tahoma", 10, FontStyle.Bold)
            .Appearance.GroupRow.Font = New Font("Tahoma", 10, FontStyle.Bold)
            .Appearance.FocusedRow.Font = New Font("Tahoma", 10, FontStyle.Bold)
            .Appearance.SelectedRow.Font = New Font("Tahoma", 10)
            .Appearance.FooterPanel.Font = New Font("Tahoma", 10)
            ' / Word wrap
            .OptionsView.RowAutoHeight = True
            .Appearance.Row.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap

            '// ปรับการเพิ่ม/แก้ไข/ลบข้อมูล
            .OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False
            .OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.False
            .OptionsBehavior.Editable = False '// True
            '// Display
            .FocusRectStyle = DrawFocusRectStyle.RowFocus
            .Appearance.FocusedCell.ForeColor = Color.Red
            .Appearance.FocusedCell.Options.UseTextOptions = True
            .OptionsSelection.EnableAppearanceFocusedCell = False '//True
            '// Make the group footers hidden.
            .OptionsView.ShowFooter = False
            .OptionsView.GroupFooterShowMode = GroupFooterShowMode.Hidden
        End With
    End Sub

    '// การค้นหาข้อมูล หรือแสดงผลข้อมูลทั้งหมด จะใช้เพียงโปรแกรมย่อยแบบฟังค์ชั่น (Function) เดียวกัน
    '// หากค่าที่ส่งมาเป็น False (หรือไม่ส่งมา จะถือเป็นค่า Default) นั่นคือให้แสดงผลข้อมูลทั้งหมด
    '// หากค่าที่ส่งมาเป็น True จะเป็นการค้นหาข้อมูล
    Private Function GetDataTable(Optional ByVal blnSearch As Boolean = False) As DataTable
        '// Join 2 Table between Food & Category.
        strSQL = _
            " SELECT Food.FoodPK, Food.FoodID, Food.FoodName, Category.CategoryName, Food.PriceCash, Food.PictureFood " & _
            " FROM Category INNER JOIN Food ON Category.CategoryPK = Food.CategoryFK "
        '// เป็นการค้นหา
        If blnSearch Then
            strSQL = strSQL & _
                    " WHERE " & _
                    " [FoodID] " & " Like '%" & Trim(txtSearch.Text) & "%'" & " OR " & _
                    " [FoodName] " & " Like '%" & Trim(txtSearch.Text) & "%'" & " OR " & _
                    " [CategoryName] " & " Like '%" & Trim(txtSearch.Text) & "%'"
            '// Else ไม่ต้องมี
        End If
        '// เอา strSQL มาเรียงต่อกัน
        strSQL = strSQL & " ORDER BY Food.FoodID "
        Dim DT As New DataTable
        With DT.Columns
            .Add("FoodPK", GetType(Int32))
            .Add("FoodID", GetType(String))
            .Add("FoodName", GetType(String))
            .Add("CategoryName", GetType(String))
            .Add("PriceCash", GetType(Double))
            .Add("PictureFood", GetType(Image))
        End With
        Try
            If Conn.State = ConnectionState.Closed Then Conn.Open()
            Cmd.Connection = Conn
            Cmd.CommandText = strSQL
            Dim DR As OleDbDataReader = Cmd.ExecuteReader
            ' / แสดงผลรูปภาพ
            Dim imgName As Image
            Dim strPath As String = MyPath(Application.StartupPath) & "Images\"
            While DR.Read()
                If DR.HasRows Then
                    '// เช็คจากฐานข้อมูลก่อนว่ามีชื่อไฟล์ภาพหรือไม่
                    If DR.Item("PictureFood").ToString <> "" Then
                        '// ทดสอบว่ามีไฟล์ภาพอยู่หรือไม่
                        If Not System.IO.File.Exists(strPath & DR.Item("PictureFood").ToString) Then
                            ' File dosn't exist. หากไม่เจอไฟล์ภาพให้แสดงผลภาพว่างเปล่าแทน
                            imgName = Image.FromFile(strPath & "NoImage.gif")
                        Else
                            imgName = Image.FromFile(strPath & DR.Item("PictureFood").ToString)
                        End If
                    Else
                        imgName = Image.FromFile(strPath & "NoImage.gif")
                    End If
                    '//
                    DT.Rows.Add(New Object() { _
                            DR.Item("FoodPK").ToString, _
                            DR.Item("FoodID").ToString, _
                            DR.Item("FoodName").ToString, _
                            DR.Item("CategoryName"), _
                            DR.Item("PriceCash"), _
                            imgName})
                End If
            End While
            DR.Close()
            Cmd.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        lblRecordCount.Text = "[จำนวน : " & DT.Rows.Count & " รายการ.]"
        '// Return DataTable
        Return DT
    End Function

    '// แสดงผลทั้งหมด
    Private Sub btnRefresh_Click(sender As System.Object, e As System.EventArgs) Handles btnRefresh.Click
        '// Refresh Data.
        Call SetupGridView()
        '// ไม่จำเป็นต้องส่งค่า Parameter (Argument) ออกไป เพราะค่า Default จะเท่ากับ False คือการแสดงผลทั้งหมด
        GridControl1.DataSource = GetDataTable()
    End Sub

    '// การค้นหา
    Private Sub txtSearch_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.KeyPress
        If Trim(txtSearch.Text.Trim) = "" Or Len(Trim(txtSearch.Text.Trim)) = 0 Then Return
        '// RetrieveData(True) It means searching for information.
        If e.KeyChar = Chr(13) Then '// Press Enter
            '// No beep.
            e.Handled = True
            '// Undesirable characters for the database ex.  ', * or %
            txtSearch.Text = txtSearch.Text.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
            Call SetupGridView()
            '// กำหนดพารามิเตอร์ให้เป็น True เพื่อแจ้งฟังค์ชั่นว่านี่คือการค้นหาข้อมูล
            GridControl1.DataSource = GetDataTable(True)
            txtSearch.Text = String.Empty
        End If
    End Sub

    '// เหตุการณ์ดับเบิ้ลคลิ๊กเมาส์ในแต่ละแถวของตารางกริด
    Private Sub GridControl1_DoubleClick(sender As Object, e As System.EventArgs) Handles GridControl1.DoubleClick
        If TryCast(GridControl1.FocusedView, GridView).RowCount = 0 Then Exit Sub
        '// การรับค่าในแต่ละหลักของตารางกริด
        Dim PK As Long = CLng(GridView1.GetRowCellValue(GridView1.GetSelectedRows(0), "FoodPK"))
        Dim FoodName As String = CStr(GridView1.GetRowCellValue(GridView1.GetSelectedRows(0), "FoodName"))
        Dim CategoryName As String = CStr(GridView1.GetRowCellValue(GridView1.GetSelectedRows(0), "CategoryName"))
        'MsgBox(PK)
        '// หาก PK = 0 จะเป็น Group Header
        '// เราจะเอาค่า Primary Key ไป Query อีกที เพื่อนำมาแก้ไขหรือลบข้อมูลออกไป
        If PK <> 0 Then DevExpress.XtraEditors.XtraMessageBox.Show( _
            "Primary Key is : " & PK & vbCrLf & _
            "Food Name is: " & FoodName & vbCrLf & _
            "Category Name is: " & CategoryName _
            )
    End Sub

    '// เหตุการณ์กด Enter ในแต่ละแถวของตารางกริด
    Private Sub GridControl1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles GridControl1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call GridControl1_DoubleClick(sender, e)
            '// ป้องกันไม่ให้เลื่อนโฟกัสไปแถวถัดไป
            e.SuppressKeyPress = True
        End If
    End Sub

    '// เปิดปิดการแสดงผลภาพ
    Private Sub chkShowPicture_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowPicture.CheckedChanged
        If GridView1.RowCount = 0 Then Return
        If chkShowPicture.Checked Then
            GridView1.Columns("PictureFood").Visible = True
        Else
            GridView1.Columns("PictureFood").Visible = False
        End If
    End Sub

    Private Sub frmXtraGrid_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Conn.State = ConnectionState.Open Then Conn.Close()
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub
End Class
