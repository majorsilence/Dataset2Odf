Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports Ionic.Zip
Imports System.Text
Imports System.Text.RegularExpressions

Public NotInheritable Class OdsReaderWriter
    ' Namespaces. We need this to initialize XmlNamespaceManager so that we can search XmlDocument.
    Private Shared namespaces As String(,) = New String(,) {{"table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"}, {"office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"}, {"style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0"}, {"text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"}, {"draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"}, {"fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"}, _
     {"dc", "http://purl.org/dc/elements/1.1/"}, {"meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"}, {"number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"}, {"presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"}, {"svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"}, {"chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0"}, _
     {"dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"}, {"math", "http://www.w3.org/1998/Math/MathML"}, {"form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0"}, {"script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0"}, {"ooo", "http://openoffice.org/2004/office"}, {"ooow", "http://openoffice.org/2004/writer"}, _
     {"oooc", "http://openoffice.org/2004/calc"}, {"dom", "http://www.w3.org/2001/xml-events"}, {"xforms", "http://www.w3.org/2002/xforms"}, {"xsd", "http://www.w3.org/2001/XMLSchema"}, {"xsi", "http://www.w3.org/2001/XMLSchema-instance"}, {"rpt", "http://openoffice.org/2005/report"}, _
     {"of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2"}, {"rdfa", "http://docs.oasis-open.org/opendocument/meta/rdfa#"}, {"config", "urn:oasis:names:tc:opendocument:xmlns:config:1.0"}}

    ' Read zip stream (.ods file is zip file).
    Private Function GetZipFile(ByVal stream As Stream) As ZipFile
        Return ZipFile.Read(stream)
    End Function

    ' Read zip file (.ods file is zip file).
    Private Function GetZipFile(ByVal inputFilePath As String) As ZipFile
        Return ZipFile.Read(inputFilePath)
    End Function

    Private Function GetContentXmlFile(ByVal zipFile As ZipFile) As XmlDocument
        ' Get file(in zip archive) that contains data ("content.xml").
        Dim contentZipEntry As ZipEntry = zipFile("content.xml")

        ' Extract that file to MemoryStream.
        Dim contentStream As Stream = New MemoryStream()
        contentZipEntry.Extract(contentStream)
        contentStream.Seek(0, SeekOrigin.Begin)

        ' Create XmlDocument from MemoryStream (MemoryStream contains content.xml).
        Dim contentXml As New XmlDocument()
        contentXml.Load(contentStream)

        Return contentXml
    End Function

    Private Function InitializeXmlNamespaceManager(ByVal xmlDocument As XmlDocument) As XmlNamespaceManager
        Dim nmsManager As New XmlNamespaceManager(xmlDocument.NameTable)

        For i As Integer = 0 To namespaces.GetLength(0) - 1
            nmsManager.AddNamespace(namespaces(i, 0), namespaces(i, 1))
        Next

        Return nmsManager
    End Function

    ''' <summary>
    ''' Read .ods file and store it in DataSet.
    ''' </summary>
    ''' <param name="inputFilePath">Path to the .ods file.</param>
    ''' <returns>DataSet that represents .ods file.</returns>
    Public Function ReadOdsFile(ByVal inputFilePath As String) As DataSet
        Dim odsZipFile As ZipFile = Me.GetZipFile(inputFilePath)

        ' Get content.xml file
        Dim contentXml As XmlDocument = Me.GetContentXmlFile(odsZipFile)

        ' Initialize XmlNamespaceManager
        Dim nmsManager As XmlNamespaceManager = Me.InitializeXmlNamespaceManager(contentXml)

        Dim odsFile As New DataSet(Path.GetFileName(inputFilePath))

        For Each tableNode As XmlNode In Me.GetTableNodes(contentXml, nmsManager)
            odsFile.Tables.Add(Me.GetSheet(tableNode, nmsManager))
        Next

        Return odsFile
    End Function

    ' In ODF sheet is stored in table:table node
    Private Function GetTableNodes(ByVal contentXmlDocument As XmlDocument, ByVal nmsManager As XmlNamespaceManager) As XmlNodeList
        Return contentXmlDocument.SelectNodes("/office:document-content/office:body/office:spreadsheet/table:table", nmsManager)
    End Function

    Private Function GetSheet(ByVal tableNode As XmlNode, ByVal nmsManager As XmlNamespaceManager) As DataTable
        Dim sheet As New DataTable(tableNode.Attributes("table:name").Value)

        Dim rowNodes As XmlNodeList = tableNode.SelectNodes("table:table-row", nmsManager)

        Dim rowIndex As Integer = 0
        For Each rowNode As XmlNode In rowNodes
            Me.GetRow(rowNode, sheet, nmsManager, rowIndex)
        Next

        Return sheet
    End Function

    Private Sub GetRow(ByVal rowNode As XmlNode, ByVal sheet As DataTable, ByVal nmsManager As XmlNamespaceManager, ByRef rowIndex As Integer)
        Dim rowsRepeated As XmlAttribute = rowNode.Attributes("table:number-rows-repeated")
        If rowsRepeated Is Nothing OrElse Convert.ToInt32(rowsRepeated.Value, CultureInfo.InvariantCulture) = 1 Then
            While sheet.Rows.Count < rowIndex
                sheet.Rows.Add(sheet.NewRow())
            End While

            Dim row As DataRow = sheet.NewRow()

            Dim cellNodes As XmlNodeList = rowNode.SelectNodes("table:table-cell", nmsManager)

            Dim cellIndex As Integer = 0
            For Each cellNode As XmlNode In cellNodes
                Me.GetCell(cellNode, row, nmsManager, cellIndex)
            Next

            sheet.Rows.Add(row)

            rowIndex += 1
        Else
            rowIndex += Convert.ToInt32(rowsRepeated.Value, CultureInfo.InvariantCulture)
        End If

        ' sheet must have at least one cell
        If sheet.Rows.Count = 0 Then
            sheet.Rows.Add(sheet.NewRow())
            sheet.Columns.Add()
        End If
    End Sub

    Private Sub GetCell(ByVal cellNode As XmlNode, ByVal row As DataRow, ByVal nmsManager As XmlNamespaceManager, ByRef cellIndex As Integer)
        Dim cellRepeated As XmlAttribute = cellNode.Attributes("table:number-columns-repeated")

        If cellRepeated Is Nothing Then
            Dim sheet As DataTable = row.Table

            While sheet.Columns.Count <= cellIndex
                sheet.Columns.Add()
            End While

            row(cellIndex) = Me.ReadCellValue(cellNode)

            cellIndex += 1
        Else
            cellIndex += Convert.ToInt32(cellRepeated.Value, CultureInfo.InvariantCulture)
        End If
    End Sub

    Private Function ReadCellValue(ByVal cell As XmlNode) As String
        Dim cellVal As XmlAttribute = cell.Attributes("office:value")

        If cellVal Is Nothing Then
            Return If([String].IsNullOrEmpty(cell.InnerText), Nothing, cell.InnerText)
        Else
            Return cellVal.Value
        End If
    End Function

    ''' <summary>
    ''' Writes DataSet as .ods file.
    ''' </summary>
    ''' <param name="odsFile">DataSet that represent .ods file.</param>
    ''' <param name="outputFilePath">The name of the file to save to.</param>
    Public Sub WriteOdsFile(ByVal odsFile As DataSet, ByVal outputFilePath As String)

        Dim tempFileLocation As String = System.IO.Path.GetTempFileName()
        Try
            'ms.Close()
            Dim fs As New FileStream(tempFileLocation, FileMode.Create)
            Dim myWriter As New BinaryWriter(fs)
            myWriter.Write(My.Resources.template)
            myWriter.Close()

            Dim templateFile As ZipFile = Me.GetZipFile(tempFileLocation)

            Dim contentXml As XmlDocument = Me.GetContentXmlFile(templateFile)

            Dim nmsManager As XmlNamespaceManager = Me.InitializeXmlNamespaceManager(contentXml)

            Dim sheetsRootNode As XmlNode = Me.GetSheetsRootNodeAndRemoveChildrens(contentXml, nmsManager)

            For Each sheet As DataTable In odsFile.Tables
                Me.SaveSheet(sheet, sheetsRootNode)
            Next

            Me.SaveContentXml(templateFile, contentXml)

            templateFile.Save(outputFilePath)
            templateFile.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error Creating ODF File", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

            If System.IO.File.Exists(tempFileLocation) Then
                System.IO.File.Delete(tempFileLocation)
            End If
        End Try
    End Sub

    Private Function GetSheetsRootNodeAndRemoveChildrens(ByVal contentXml As XmlDocument, ByVal nmsManager As XmlNamespaceManager) As XmlNode
        Dim tableNodes As XmlNodeList = Me.GetTableNodes(contentXml, nmsManager)

        Dim sheetsRootNode As XmlNode = tableNodes.Item(0).ParentNode
        ' remove sheets from template file
        For Each tableNode As XmlNode In tableNodes
            sheetsRootNode.RemoveChild(tableNode)
        Next

        Return sheetsRootNode
    End Function

    Private Sub SaveSheet(ByVal sheet As DataTable, ByVal sheetsRootNode As XmlNode)
        Dim ownerDocument As XmlDocument = sheetsRootNode.OwnerDocument

        Dim sheetNode As XmlNode = ownerDocument.CreateElement("table:table", Me.GetNamespaceUri("table"))

        Dim sheetName As XmlAttribute = ownerDocument.CreateAttribute("table:name", Me.GetNamespaceUri("table"))
        sheetName.Value = sheet.TableName
        sheetNode.Attributes.Append(sheetName)

        Me.SaveColumnDefinition(sheet, sheetNode, ownerDocument)

        Me.SaveRows(sheet, sheetNode, ownerDocument)

        sheetsRootNode.AppendChild(sheetNode)
    End Sub

    Private Sub SaveColumnDefinition(ByVal sheet As DataTable, ByVal sheetNode As XmlNode, ByVal ownerDocument As XmlDocument)
        Dim columnDefinition As XmlNode = ownerDocument.CreateElement("table:table-column", Me.GetNamespaceUri("table"))

        Dim columnsCount As XmlAttribute = ownerDocument.CreateAttribute("table:number-columns-repeated", Me.GetNamespaceUri("table"))
        columnsCount.Value = sheet.Columns.Count.ToString(CultureInfo.InvariantCulture)
        columnDefinition.Attributes.Append(columnsCount)

        sheetNode.AppendChild(columnDefinition)
    End Sub

    Private Sub SaveRows(ByVal sheet As DataTable, ByVal sheetNode As XmlNode, ByVal ownerDocument As XmlDocument)
        Dim rows As DataRowCollection = sheet.Rows
        For i As Integer = 0 To rows.Count - 1
            Dim rowNode As XmlNode = ownerDocument.CreateElement("table:table-row", Me.GetNamespaceUri("table"))

            Me.SaveCell(rows(i), rowNode, ownerDocument)

            sheetNode.AppendChild(rowNode)
        Next
    End Sub

    Private Shared CELL_REF_REGEX As New Regex("((('[^']+')|([A-Za-z][A-Za-z\d_]+))[\.!])?[A-Za-z][A-Za-z]?\d{1,8}")

    Private Shared Function FixFormula(ByVal formulaStr As String) As String
        Dim result As New StringBuilder()
        Dim sourceStrIdx As Integer = 0
        For Each nextMatch As Match In CELL_REF_REGEX.Matches(formulaStr)
            Dim cellReference As String = formulaStr.Substring(nextMatch.Index, nextMatch.Length).Replace("!"c, "."c)
            If cellReference.IndexOf("."c) = -1 Then
                cellReference = "."c & cellReference
            End If
            result.Append(formulaStr.Substring(sourceStrIdx, nextMatch.Index - sourceStrIdx) & "["c & cellReference & "]"c)
            sourceStrIdx = nextMatch.Index + nextMatch.Length
        Next
        result.Append(formulaStr.Substring(sourceStrIdx))
        Return result.ToString()
    End Function


    Private Sub SaveCell(ByVal row As DataRow, ByVal rowNode As XmlNode, ByVal ownerDocument As XmlDocument)
        Dim cells As Object() = row.ItemArray

        For i As Integer = 0 To cells.Length - 1
            Dim cellNode As XmlElement = ownerDocument.CreateElement("table:table-cell", Me.GetNamespaceUri("table"))
            Dim cell As Object = cells(i)
            If cell IsNot DBNull.Value Then
                Dim isNumber As Boolean, isFormula As Boolean = False
                If (cell.ToString().Length = 0) OrElse (cell.ToString()(0) = "+"c) OrElse ((cell.ToString()(0) = "0"c) AndAlso (cell.ToString().Length > 1) AndAlso (Not cell.ToString().Contains("."))) Then
                    isNumber = False
                Else
                    isNumber = True
                    isFormula = cell.ToString()(0) = "="c
                    If Not isFormula Then
                        Try
                            isNumber = IsNumeric(cell.ToString)
                        Catch generatedExceptionName As FormatException
                            isNumber = False
                        End Try
                    End If
                End If
                Dim valueType As XmlAttribute = ownerDocument.CreateAttribute("office:value-type", Me.GetNamespaceUri("office"))
                valueType.Value = If(isNumber, "float", "string")
                cellNode.Attributes.Append(valueType)

                If isNumber Then
                    Dim value As XmlAttribute = ownerDocument.CreateAttribute("office:value", Me.GetNamespaceUri("office"))
                    value.Value = If(isFormula, "0", cell.ToString())
                    cellNode.Attributes.Append(value)
                    If isFormula Then
                        Dim formulaAttr As XmlAttribute = ownerDocument.CreateAttribute("table:formula", Me.GetNamespaceUri("table"))
                        formulaAttr.Value = "of:" & FixFormula(cell.ToString())
                        cellNode.Attributes.Append(formulaAttr)
                    End If
                Else
                    Dim cellValue As XmlElement = ownerDocument.CreateElement("text:p", Me.GetNamespaceUri("text"))
                    cellValue.InnerText = cell.ToString()
                    cellNode.AppendChild(cellValue)
                End If
            End If

            rowNode.AppendChild(cellNode)
        Next
    End Sub



    Private Sub SaveContentXml(ByVal templateFile As ZipFile, ByVal contentXml As XmlDocument)
        templateFile.RemoveEntry("content.xml")

        Dim memStream As New MemoryStream()
        contentXml.Save(memStream)
        memStream.Seek(0, SeekOrigin.Begin)

        templateFile.AddEntry("content.xml", memStream)
    End Sub

    Private Function GetNamespaceUri(ByVal prefix As String) As String
        For i As Integer = 0 To namespaces.GetLength(0) - 1
            If namespaces(i, 0) = prefix Then
                Return namespaces(i, 1)
            End If
        Next

        Throw New InvalidOperationException("Can't find that namespace URI")
    End Function
End Class

