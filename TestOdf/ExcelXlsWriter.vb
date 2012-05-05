Imports System.IO
Imports System.Data
Imports NPOI.HSSF.UserModel
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web
Imports NPOI.SS.UserModel
Imports NPOI.SS.Util
Imports NPOI.HSSF.Util
Imports NPOI.POIFS.FileSystem
Imports NPOI.HPSF


Public Class ExcelXlsWriter
    Public Sub WriteXls(ByVal ds As DataSet, ByVal savePath As String)
        Try
            Using exporter As New NpoiExport()

                For Each dt As DataTable In ds.Tables
                    exporter.ExportDataTableToWorkbook(dt, dt.TableName)
                Next

                Dim fs As New FileStream(savePath, FileMode.Create)
                Dim file() As Byte = exporter.GetBytes
                fs.Write(file, 0, file.Length)
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error creating xls file", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class


Public Class NpoiExport
    Implements IDisposable
    Const MaximumNumberOfRowsPerSheet As Integer = 65500
    Const MaximumSheetNameLength As Integer = 25
    Protected Property Workbook() As HSSFWorkbook
        Get
            Return m_Workbook
        End Get
        Set(ByVal value As HSSFWorkbook)
            m_Workbook = value
        End Set
    End Property
    Private m_Workbook As HSSFWorkbook

    Public Sub New()
        Me.Workbook = New HSSFWorkbook()
    End Sub

    Protected Function EscapeSheetName(ByVal sheetName As String) As String
        Dim escapedSheetName As String = sheetName.Replace("/", "-").Replace("\", " ").Replace("?", String.Empty).Replace("*", String.Empty).Replace("[", String.Empty).Replace("]", String.Empty).Replace(":", String.Empty)

        If escapedSheetName.Length > MaximumSheetNameLength Then
            escapedSheetName = escapedSheetName.Substring(0, MaximumSheetNameLength)
        End If

        Return escapedSheetName
    End Function

    Protected Function CreateExportDataTableSheetAndHeaderRow(ByVal exportData As DataTable, _
                                                              ByVal sheetName As String, _
                                                              ByVal headerRowStyle As HSSFCellStyle) As HSSFSheet


        Dim sheet As HSSFSheet = CType(Me.Workbook.CreateSheet(EscapeSheetName(sheetName)), HSSFSheet)

        ' Create the header row
        Dim row As HSSFRow = CType(sheet.CreateRow(0), HSSFRow)

        For colIndex As Integer = 0 To exportData.Columns.Count - 1
            Dim cell As HSSFCell = CType(row.CreateCell(colIndex), HSSFCell)
            cell.SetCellValue(exportData.Columns(colIndex).ColumnName)

            If headerRowStyle IsNot Nothing Then
                cell.CellStyle = headerRowStyle
            End If
        Next

        Return sheet
    End Function

    Public Sub ExportDataTableToWorkbook(ByVal exportData As DataTable, ByVal sheetName As String)
        ' Create the header row cell style
        Dim headerLabelCellStyle As HSSFCellStyle = CType(Me.Workbook.CreateCellStyle(), HSSFCellStyle)
        headerLabelCellStyle.BorderBottom = BorderStyle.THIN
        Dim headerLabelFont As HSSFFont = CType(Me.Workbook.CreateFont(), HSSFFont)
        headerLabelFont.Boldweight = CShort(FontBoldWeight.BOLD)
        headerLabelCellStyle.SetFont(headerLabelFont)

        Dim sheet As HSSFSheet = CreateExportDataTableSheetAndHeaderRow(exportData, sheetName, headerLabelCellStyle)
        Dim currentNPOIRowIndex As Integer = 1
        Dim sheetCount As Integer = 1

        For rowIndex As Integer = 0 To exportData.Rows.Count - 1
            If currentNPOIRowIndex >= MaximumNumberOfRowsPerSheet Then
                sheetCount += 1
                currentNPOIRowIndex = 1

                sheet = CreateExportDataTableSheetAndHeaderRow(exportData, sheetName & " - " & sheetCount, headerLabelCellStyle)
            End If

            Dim row As HSSFRow = CType(sheet.CreateRow(System.Math.Max(System.Threading.Interlocked.Increment(currentNPOIRowIndex), currentNPOIRowIndex - 1)), HSSFRow)

            For colIndex As Integer = 0 To exportData.Columns.Count - 1
                Dim cell As HSSFCell = CType(row.CreateCell(colIndex), HSSFCell)
                cell.SetCellValue(exportData.Rows(rowIndex)(colIndex).ToString())
            Next
        Next
    End Sub

    Public Function GetBytes() As Byte()
        Using buffer As New MemoryStream()
            Me.Workbook.Write(buffer)
            Return buffer.GetBuffer()
        End Using
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        If Me.Workbook IsNot Nothing Then
            Me.Workbook = Nothing
        End If
    End Sub
End Class
