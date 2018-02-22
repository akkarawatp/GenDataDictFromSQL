Imports System.IO
Imports System.Data.SqlClient
Imports LinqDB.ConnectDB
Imports OfficeOpenXml
Imports wDoc = SautinSoft.Document
Imports SautinSoft.Document.Tables

Public Class Form1

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        GetSetting()
    End Sub

    Private Sub CreateSetting()
        Dim iniFile As String = Application.StartupPath & "\Setting.ini"
        Dim ini As New IniReader(iniFile)
        ini.Section = "Setting"

        ini.Write("DataSource", txtDataSource.Text)
        ini.Write("DbName", txtDatabaseName.Text)
        ini.Write("DbUserID", txtUserID.Text)
        ini.Write("DbPassword", txtPassword.Text)
        ini.Write("TableName", txtTableName.Text)
        ini.Write("OutputFile", txtOutputPath.Text)
        ini = Nothing

    End Sub

    Private Sub GetSetting()
        Dim iniFile As String = Application.StartupPath & "\Setting.ini"
        If File.Exists(iniFile) = True Then
            Dim ini As New IniReader(iniFile)
            ini.Section = "Setting"
            txtDataSource.Text = ini.ReadString("DataSource")
            txtDatabaseName.Text = ini.ReadString("DbName")
            txtUserID.Text = ini.ReadString("DbUserID")
            txtPassword.Text = ini.ReadString("DbPassword")
            txtTableName.Text = ini.ReadString("TableName")
            txtOutputPath.Text = ini.ReadString("OutputFile")
            ini = Nothing
        End If

    End Sub

    Private Sub btnBrowseFile_Click(sender As Object, e As EventArgs) Handles btnBrowseFile.Click
        Dim fle As New SaveFileDialog
        fle.Filter = "Excel File|*.xlsx"
        If txtOutputPath.Text.Trim <> "" Then
            fle.InitialDirectory = New DirectoryInfo(txtOutputPath.Text).FullName
        End If

        If fle.ShowDialog = DialogResult.OK Then
            txtOutputPath.Text = fle.FileName
        End If
    End Sub

    Private Function ValidateData() As Boolean
        Dim ret As Boolean = True
        If txtDataSource.Text.Trim() = "" Then
            MsgBox("กรุณาระบุ Data source", MsgBoxStyle.OkOnly + 48, "Validate Data")
            ret = False
            txtDataSource.Focus()
        ElseIf txtUserID.Text.Trim = "" Then
            MsgBox("กรุณาระบุ Database UserID", MsgBoxStyle.OkOnly + 48, "Validate Data")
            ret = False
            txtUserID.Focus()
        ElseIf txtDatabaseName.Text.Trim = "" Then
            MsgBox("กรุณาระบุ Database Name", MsgBoxStyle.OkOnly + 48, "Validate Data")
            ret = False
            txtDatabaseName.Focus()
        ElseIf txtPassword.Text.Trim = "" Then
            MsgBox("กรุณาระบุ Database Password", MsgBoxStyle.OkOnly + 48, "Validate Data")
            ret = False
            txtPassword.Focus()
            'ElseIf txtTableName.Text.Trim = "" Then
            '    MsgBox("กรุณาระบุ Table Name", MsgBoxStyle.OkOnly + 48, "Validate Data")
            '    ret = False
            '    txtTableName.Focus()
        ElseIf txtOutputPath.Text.Trim = "" Then
            MsgBox("กรุณาระบุ Output Path", MsgBoxStyle.OkOnly + 48, "Validate Data")
            ret = False
            txtOutputPath.Focus()
        ElseIf SqlDB.ChkConnection(GetConnectionString()) = False Then
            MsgBox("Database Connection Fail!", MsgBoxStyle.OkOnly + 48, "Validate Data")
            ret = False
            txtDataSource.Focus()
        End If

        Return ret
    End Function

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        If ValidateData() = True Then
            Dim tDt As DataTable = GetAllTable()
            If txtTableName.Text.Trim <> "" Then
                tDt.DefaultView.RowFilter = "table_name='" & txtTableName.Text & "'"
                tDt = tDt.DefaultView.ToTable
            End If

            If tDt.Rows.Count > 0 Then
                ProgressBar1.Maximum = tDt.Rows.Count + 3
                ProgressBar1.Value = 1
                Application.DoEvents()

                GenerateFileExcel(tDt)
                'GenerateFileWord(tDt)
            End If
            tDt.Dispose()
        End If
    End Sub



    Private Sub GenerateFileWord(tDt As DataTable)
        Dim docx As New wDoc.DocumentCore
        Dim s As New wDoc.Section(docx)
        docx.Sections.Add(s)


        For i As Integer = 0 To tDt.Rows.Count - 1
            Dim dr As DataRow = tDt.Rows(i)
            'Display Table Name
            Dim TableName As String = dr("table_name")
            Dim TableComment As String = GetTableComment(TableName)

            lblProgressText.Text = "Generate Table " & TableName.ToUpper & " ( " & i & "/" & tDt.Rows.Count & " )"
            Application.DoEvents()

            Dim pTableName As New wDoc.Paragraph(docx)
            pTableName.ParagraphFormat.Alignment = HorizontalAlignment.Left
            pTableName.ParagraphFormat.SpaceBefore = wDoc.LengthUnitConverter.Convert(3, wDoc.LengthUnit.Millimeter, wDoc.LengthUnit.Point)
            pTableName.ParagraphFormat.SpaceAfter = wDoc.LengthUnitConverter.Convert(3, wDoc.LengthUnit.Millimeter, wDoc.LengthUnit.Point)
            pTableName.Content.Start.Insert((i + 1) & ". " & TableName, New wDoc.CharacterFormat() With {
                                        .FontName = "Tahoma",
                                        .Size = 8.0
                                   })
            s.Blocks.Add(pTableName)

            Dim pTableComment As New wDoc.Paragraph(docx)
            pTableComment.ParagraphFormat.Alignment = HorizontalAlignment.Left
            pTableComment.ParagraphFormat.SpaceBefore = wDoc.LengthUnitConverter.Convert(3, wDoc.LengthUnit.Millimeter, wDoc.LengthUnit.Point)
            pTableComment.ParagraphFormat.SpaceAfter = wDoc.LengthUnitConverter.Convert(3, wDoc.LengthUnit.Millimeter, wDoc.LengthUnit.Point)
            pTableComment.Content.End.Insert(TableComment, New wDoc.CharacterFormat() With {
                                        .FontName = "Tahoma",
                                        .Size = 8.0
                                   })
            s.Blocks.Add(pTableComment)


            'Table Data
            Dim t As New Table(docx)
            Dim tWidth As Double = wDoc.LengthUnitConverter.Convert(160, wDoc.LengthUnit.Millimeter, wDoc.LengthUnit.Point)
            t.TableFormat.PreferredWidth = New TableWidth(tWidth, TableWidthUnit.Point)
            t.TableFormat.Alignment = HorizontalAlignment.Left
#Region "Header Row"
            Dim hRow As New TableRow(docx)
            'hRow.RowFormat.Height = New TableRowHeight(wDoc.LengthUnitConverter.Convert(7, wDoc.LengthUnit.Millimeter, wDoc.LengthUnit.Point), TableRowHeightRule.Auto)

            Dim hColNameCell As New TableCell(docx)
            '// Set some cell formatting
            hColNameCell.CellFormat.Borders.SetBorders(wDoc.MultipleBorderTypes.Outside, wDoc.BorderStyle.Single, wDoc.Color.Black, 1)
            'hColNameCell.CellFormat.PreferredWidth = New TableWidth(wDoc.LengthUnitConverter.Convert(30, wDoc.LengthUnit.Millimeter, wDoc.LengthUnit.Point), TableWidthUnit.Point)
            SetTextToTableCell("Column Name", docx, hColNameCell, New wDoc.CharacterFormat() With {.FontName = "Tahoma", .FontColor = wDoc.Color.Black, .Size = 8.0, .Bold = True})
            hRow.Cells.Add(hColNameCell)

            Dim hDataType As New TableCell(docx)
            hDataType.CellFormat.Borders.SetBorders(wDoc.MultipleBorderTypes.Outside, wDoc.BorderStyle.Single, wDoc.Color.Black, 1)
            hRow.Cells.Add(hDataType)
            SetTextToTableCell("Data Type", docx, hDataType, New wDoc.CharacterFormat() With {.FontName = "Tahoma", .FontColor = wDoc.Color.Black, .Size = 8.0, .Bold = True})

            Dim hComment As New TableCell(docx)
            hComment.CellFormat.Borders.SetBorders(wDoc.MultipleBorderTypes.Outside, wDoc.BorderStyle.Single, wDoc.Color.Black, 1)
            hRow.Cells.Add(hComment)
            SetTextToTableCell("Comment", docx, hComment, New wDoc.CharacterFormat() With {.FontName = "Tahoma", .FontColor = wDoc.Color.Black, .Size = 8.0, .Bold = True})

            Dim hPK As New TableCell(docx)
            hPK.CellFormat.Borders.SetBorders(wDoc.MultipleBorderTypes.Outside, wDoc.BorderStyle.Single, wDoc.Color.Black, 1)
            hRow.Cells.Add(hPK)
            SetTextToTableCell("PK", docx, hPK, New wDoc.CharacterFormat() With {.FontName = "Tahoma", .FontColor = wDoc.Color.Black, .Size = 8.0, .Bold = True})

            Dim hUQ As New TableCell(docx)
            hUQ.CellFormat.Borders.SetBorders(wDoc.MultipleBorderTypes.Outside, wDoc.BorderStyle.Single, wDoc.Color.Black, 1)
            hRow.Cells.Add(hUQ)
            SetTextToTableCell("UQ", docx, hUQ, New wDoc.CharacterFormat() With {.FontName = "Tahoma", .FontColor = wDoc.Color.Black, .Size = 8.0, .Bold = True})

            Dim hNotNull As New TableCell(docx)
            hNotNull.CellFormat.Borders.SetBorders(wDoc.MultipleBorderTypes.Outside, wDoc.BorderStyle.Single, wDoc.Color.Black, 1)
            hRow.Cells.Add(hNotNull)
            SetTextToTableCell("Not Null", docx, hNotNull, New wDoc.CharacterFormat() With {.FontName = "Tahoma", .FontColor = wDoc.Color.Black, .Size = 8.0, .Bold = True})

            Dim hDefault As New TableCell(docx)
            hDefault.CellFormat.Borders.SetBorders(wDoc.MultipleBorderTypes.Outside, wDoc.BorderStyle.Single, wDoc.Color.Black, 1)
            hRow.Cells.Add(hDefault)
            SetTextToTableCell("Default", docx, hDefault, New wDoc.CharacterFormat() With {.FontName = "Tahoma", .FontColor = wDoc.Color.Black, .Size = 8.0, .Bold = True})

            t.Rows.Add(hRow)
#End Region

#Region "Data Row"
            'หา Column ของ Table ที่ระบุ
            Dim cDt As DataTable = GetTableColumn(TableName)
            If cDt.Rows.Count > 0 Then
                For Each cDr As DataRow In cDt.Rows
                    Dim dRow As New TableRow(docx)
                    SetContentDataRow(cDr("COLUMN_NAME"), docx, dRow, 45)
                    SetContentDataRow(GetFormatColumnTypeName(cDr("TYPE_NAME"), cDr("LENGTH")), docx, dRow, 16)
                    SetContentDataRow(GetColumnComment(TableName, cDr("COLUMN_NAME")), docx, dRow, 38)
                    SetContentDataRow(GetPKColumn(TableName, cDr("COLUMN_NAME")), docx, dRow, 6)
                    SetContentDataRow(GetUQColumn(TableName, cDr("COLUMN_NAME")), docx, dRow, 6)
                    SetContentDataRow(IIf(cDr("NULLABLE") = 1, "", "Y"), docx, dRow, 6)

                    Dim DefalueValue As String = ""
                    If Convert.IsDBNull(cDr("COLUMN_DEF")) = False Then
                        DefalueValue = ReplaceBracket(cDr("COLUMN_DEF"))
                        SetContentDataRow(DefalueValue, docx, dRow, 25)
                    Else
                        SetContentDataRow(" ", docx, dRow, 25)
                    End If
                    t.Rows.Add(dRow)
                Next
            End If
            cDt.Dispose()
#End Region

            s.Blocks.Add(t)
            s.Content.End.Insert(" ")
            ProgressBar1.Value += 1
            Application.DoEvents()
        Next

        lblProgressText.Text = "Save output file..."
        ProgressBar1.Value += 1
        Application.DoEvents()
        docx.Save(txtOutputPath.Text)

        lblProgressText.Text = "Complete"
        ProgressBar1.Value = ProgressBar1.Maximum
        Application.DoEvents()
    End Sub

    Private Sub SetContentDataRow(txt As String, docx As wDoc.DocumentCore, dRow As TableRow, CellWidth As Integer)
        Dim dCell As New TableCell(docx)
        dCell.CellFormat.Borders.SetBorders(wDoc.MultipleBorderTypes.Outside, wDoc.BorderStyle.Single, wDoc.Color.Black, 1)
        Dim cWidth As Double = wDoc.LengthUnitConverter.Convert(CellWidth, wDoc.LengthUnit.Centimeter, wDoc.LengthUnit.Point)
        dCell.CellFormat.PreferredWidth = New TableWidth(cWidth, TableWidthUnit.Point)
        dRow.Cells.Add(dCell)

        'Dim p As New wDoc.Paragraph(docx)
        'p.ParagraphFormat.Alignment = HorizontalAlignment.Left
        'p.ParagraphFormat.SpaceBefore = wDoc.LengthUnitConverter.Convert(1, wDoc.LengthUnit.Millimeter, wDoc.LengthUnit.Point)
        'p.ParagraphFormat.SpaceAfter = wDoc.LengthUnitConverter.Convert(1, wDoc.LengthUnit.Millimeter, wDoc.LengthUnit.Point)
        'p.ParagraphFormat.LineSpacing = 1
        dCell.Content.Start.Insert(txt, New wDoc.CharacterFormat() With {.FontName = "Tahoma", .FontColor = wDoc.Color.Black, .Size = 8.0})

        'p.Content.End.Insert(txt, New wDoc.CharacterFormat() With {.FontName = "Tahoma", .FontColor = wDoc.Color.Black, .Size = 8.0})
    End Sub

    Private Sub SetTextToTableCell(txt As String, docx As wDoc.DocumentCore, cell As TableCell, txtStyle As wDoc.CharacterFormat)
        Dim p As New wDoc.Paragraph(docx)
        p.ParagraphFormat.Alignment = HorizontalAlignment.Left
        p.ParagraphFormat.SpaceBefore = wDoc.LengthUnitConverter.Convert(3, wDoc.LengthUnit.Millimeter, wDoc.LengthUnit.Point)
        p.ParagraphFormat.SpaceAfter = wDoc.LengthUnitConverter.Convert(3, wDoc.LengthUnit.Millimeter, wDoc.LengthUnit.Point)
        p.Content.Start.Insert(txt, txtStyle)
        cell.Blocks.Add(p)
    End Sub



    Private Sub GenerateFileExcel(tDt As DataTable)
        Using ep As New ExcelPackage
            Dim i As Integer = 1
            Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add("Output")
            Dim HeaderRow As Integer = 3   'ถ้า Table แรก ให้เริ่มที่ Row ที่ 3

            For Each tDr As DataRow In tDt.Rows
                Dim TableName As String = tDr("table_name")
                lblProgressText.Text = "Generate Table " & TableName.ToUpper & " ( " & i & "/" & tDt.Rows.Count & " )"
                Application.DoEvents()

                'หา Column ของ Table ที่ระบุ
                Dim cDt As DataTable = GetTableColumn(TableName)
                If cDt.Rows.Count > 0 Then
                    Dim dt As New DataTable
                    dt.Columns.Add("Column Name")
                    dt.Columns.Add("Data Type")
                    dt.Columns.Add("Comment")
                    dt.Columns.Add("PK")
                    dt.Columns.Add("UQ")
                    dt.Columns.Add("Not Null")
                    dt.Columns.Add("Default")

                    For Each cDr As DataRow In cDt.Rows
                        Dim dr As DataRow = dt.NewRow
                        dr("Column Name") = cDr("COLUMN_NAME")
                        dr("Data Type") = GetFormatColumnTypeName(cDr("TYPE_NAME"), cDr("LENGTH"))
                        dr("Comment") = GetColumnComment(TableName, cDr("COLUMN_NAME"))
                        dr("PK") = GetPKColumn(TableName, cDr("COLUMN_NAME"))
                        dr("UQ") = GetUQColumn(TableName, cDr("COLUMN_NAME"))
                        dr("Not Null") = IIf(cDr("NULLABLE") = 1, "", "Y")
                        If Convert.IsDBNull(cDr("COLUMN_DEF")) = False Then dr("Default") = ReplaceBracket(cDr("COLUMN_DEF"))

                        dt.Rows.Add(dr)
                    Next

                    If dt.Rows.Count > 0 Then
                        ExportDatatableToExcel(ep, ws, HeaderRow, i & ". " & TableName, dt, GetTableComment(TableName))
                        HeaderRow += (dt.Rows.Count + 5)
                    End If
                    dt.Dispose()
                End If
                cDt.Dispose()
                i += 1
                ProgressBar1.Value += 1
            Next

            lblProgressText.Text = "Save output file..."
            ProgressBar1.Value += 1
            Application.DoEvents()

            If File.Exists(txtOutputPath.Text) = True Then
                File.SetAttributes(txtOutputPath.Text, FileAttributes.Normal)
                File.Delete(txtOutputPath.Text)
            End If

            Dim f As New IO.FileInfo(txtOutputPath.Text)
            ep.SaveAs(f)
            Threading.Thread.Sleep(5000)

            If IO.File.Exists(f.FullName) = True Then
                CreateSetting()

                lblProgressText.Text = "Generate Complete"
                ProgressBar1.Value += 1
                Application.DoEvents()
                MessageBox.Show("Complete")
            End If
            f = Nothing
        End Using
    End Sub


    Private Sub ExportDatatableToExcel(ep As ExcelPackage, ws As ExcelWorksheet, HeaderRow As Integer, TableName As String, DT As DataTable, TableComment As String)
        ws.Cells("A" & (HeaderRow - 2)).Value = TableName.ToUpper
        ws.Cells("A" & (HeaderRow - 1)).Value = TableComment

        ws.Cells("A" & HeaderRow).LoadFromDataTable(DT, True)
        Dim hRow As Integer = HeaderRow

        Using RowHeader As ExcelRange = ws.Cells(hRow, 1, hRow, DT.Columns.Count)
            RowHeader.Style.Font.Bold = True
            RowHeader.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
            RowHeader.Style.Fill.BackgroundColor.SetColor(Color.Gray)
            RowHeader.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
            RowHeader.Style.Font.Color.SetColor(Color.Black)
            'RowHeader.AutoFitColumns()
        End Using

        Using RowContent As ExcelRange = ws.Cells(hRow + 1, 1, hRow + DT.Rows.Count + 1, DT.Columns.Count)
            RowContent.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin
            RowContent.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin
            RowContent.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin
            RowContent.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin
        End Using

        For i As Integer = 0 To DT.Columns.Count - 1
            Dim ColumType As String = DT.Columns(i).DataType.Name.ToLower
            If ColumType = "datetime" Then
                ws.Cells(hRow + 1, i + 1, hRow + DT.Rows.Count + 1, i + 1).Style.Numberformat.Format = "mmm dd yyyy HH:MM:ss"
            End If
        Next

        ws.Cells(hRow, 1, hRow + DT.Rows.Count, DT.Columns.Count).AutoFitColumns()
    End Sub
    Private Function ReplaceBracket(str As String) As String
        Dim ret As String = str

        Do
            If ret.StartsWith("(") = True And ret.EndsWith(")") = True Then
                ret = ret.Substring(1, ret.Length - 2)
            End If
        Loop While ret.StartsWith("(") = True And ret.EndsWith(")") = True

        Return ret
    End Function


#Region "Database Function"

    Private ReadOnly Property GetConnectionString() As String
        Get
            Return "Data Source=" & txtDataSource.Text & ";Initial Catalog=" & txtDatabaseName.Text & ";User ID=" & txtUserID.Text & ";Password=" & txtPassword.Text & ";"
        End Get
    End Property
    Private Function GetTableColumn(ByVal tbName As String) As DataTable
        Dim sql As String = "EXEC SP_COLUMNS " & SqlDB.SetString(tbName)
        Dim dt As DataTable = SqlDB.ExecuteTable(sql, SqlDB.GetConnection(GetConnectionString))
        If dt.Rows.Count > 0 Then
            If dt.Columns.Contains("msrepl_tran_version") = True Then
                dt.Columns.Remove("msrepl_tran_version")
            End If
            If dt.Columns.Contains("rowguid") = True Then
                dt.Columns.Remove("rowguid")
            End If
        End If

        Return dt
    End Function

    Private Function GetPKColumn(TableName As String, ColumnName As String) As String
        Dim ret As String = ""
        Try
            Dim conn As SqlConnection = SqlDB.GetConnection(GetConnectionString)
            Dim tmpTable As DataTable = SqlDB.ExecuteTable("EXEC SP_PKEYS " & SqlDB.SetString(TableName), conn)
            For Each dRow As DataRow In tmpTable.Rows
                If ColumnName = dRow("column_name").ToString() Then
                    ret = "Y"
                    Exit For
                End If
            Next

            conn.Close()
            conn.Dispose()

        Catch ex As Exception
        End Try

        Return ret
    End Function
    Private Function GetUQColumn(TableName As String, ColumnName As String) As String
        Dim ret As String = ""

        Dim conn As SqlConnection = SqlDB.GetConnection(GetConnectionString)
        Dim sql As String = "exec sp_indexcolumns_managed " & SqlDB.SetString(txtDatabaseName.Text) & ",null, " & SqlDB.SetString(TableName) & ",null,'" & ColumnName & "'"
        Dim dt As DataTable = SqlDB.ExecuteTable(sql, conn)
        If dt.Rows.Count > 0 Then
            ret = "Y"
        End If
        dt.Dispose()
        conn.Close()
        conn.Dispose()

        Return ret
    End Function

    Private Function GetTableComment(TableName As String) As String
        Dim ret As String = ""
        Dim sql As String = "SELECT sys.objects.name AS TableName, ep.name AS PropertyName, "
        sql += " ep.value AS table_desc"
        sql += " From sys.objects "
        sql += " CROSS APPLY fn_listextendedproperty(default,'SCHEMA', schema_name(schema_id),'TABLE', name, null, null) ep"
        sql += " WHERE sys.objects.name NOT IN ('sysdiagrams') "
        sql += " and sys.objects.name='" & TableName & "' "
        sql += " ORDER BY sys.objects.name"

        Dim conn As SqlConnection = SqlDB.GetConnection(GetConnectionString)
        Dim dt As DataTable = SqlDB.ExecuteTable(sql, conn)

        If dt.Rows.Count > 0 Then
            If Convert.IsDBNull(dt.Rows(0)("table_desc")) = False Then
                ret = dt.Rows(0)("table_desc")
            End If
        End If
        Return ret

    End Function

    Private Function GetColumnComment(TableName As String, ColumnName As String) As String
        Dim ret As String = ""
        Dim sql As String = "select sep.value column_desc "
        sql += " From sys.tables st "
        sql += " inner join sys.columns sc On st.object_id = sc.object_id "
        sql += " left join sys.extended_properties sep on st.object_id = sep.major_id "
        sql += "                                 And sc.column_id = sep.minor_id "
        sql += "                                  And sep.name = 'MS_Description'"
        sql += " where st.name ='" & TableName & "'"
        sql += " and sc.name = '" & ColumnName & "'"

        Dim conn As SqlConnection = SqlDB.GetConnection(GetConnectionString)
        Dim dt As DataTable = SqlDB.ExecuteTable(sql, conn)

        If dt.Rows.Count > 0 Then
            If Convert.IsDBNull(dt.Rows(0)("column_desc")) = False Then
                ret = dt.Rows(0)("column_desc")
            End If
        End If
        Return ret
    End Function
    Private Function GetFormatColumnTypeName(DataType As String, DataLength As Long) As String
        Dim ret As String = DataType

        Dim vTypeName As String = DataType.ToUpper
        If vTypeName = "VARCHAR" Then
            ret = DataType & "(" & DataLength & ")"
        ElseIf vTypeName = "CHAR" Then
            ret = DataType & "(" & DataLength & ")"
        ElseIf vTypeName = "NVARCHAR" Then
            ret = DataType & "(" & DataLength & ")"
        ElseIf vTypeName = "NCHAR" Then
            ret = DataType & "(" & DataLength & ")"
        ElseIf vTypeName = "TEXT" Then
            ret = DataType
        ElseIf vTypeName = "FLOAT" Then
            ret = DataType
        ElseIf vTypeName = "DOUBLE" Then
            ret = DataType
        ElseIf vTypeName = "DECIMAL" Then
            ret = DataType
        ElseIf vTypeName = "BIGINT" Then
            ret = DataType
        ElseIf vTypeName = "BIGINT IDENTITY" Then
            ret = DataType
        ElseIf vTypeName = "INT" Then
            ret = DataType
        ElseIf vTypeName = "INT IDENTITY" Then
            ret = DataType
        ElseIf vTypeName = "SMALLINT" Then
            ret = DataType
        ElseIf vTypeName = "DATETIME" Then
            ret = DataType
        ElseIf vTypeName = "DATETIME2" Then
            ret = DataType
        ElseIf vTypeName = "DATE" Then
            ret = DataType
        ElseIf vTypeName = "BIT" Then
            ret = DataType
        ElseIf vTypeName = "UNIQUEIDENTIFIER" Then  'uniqueidentifier
            ret = DataType
        ElseIf vTypeName = "IMAGE" Then
            ret = DataType
        End If

        Return ret
    End Function


    Private Function GetAllTable() As DataTable
        Dim ret As New DataTable
        Try
            Dim Sql As String = "EXEC SP_TABLES null,null,'" & txtDatabaseName.Text & "'"
            Dim dt As New DataTable
            dt = SqlDB.ExecuteTable(Sql, SqlDB.GetConnection(GetConnectionString))
            dt.DefaultView.RowFilter = "table_owner = 'dbo' and table_type='TABLE' and table_name<>'sysdiagrams'"

            ret = dt.DefaultView.ToTable
        Catch ex As Exception
            ret = New DataTable
        End Try

        Return ret
    End Function


#End Region
End Class
