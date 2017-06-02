Imports System.Drawing.Text
Imports Aspose.Words
Imports Aspose.Words.Tables
Imports System.Environment

Public Class MainForm
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ExportWord("E:\Test.doc")
        End
    End Sub

    Private Sub ExportWord(WordFilePath As String)
        Dim WordDocument As Document = New Document()
        Dim WordDocumentBuilder As DocumentBuilder = New DocumentBuilder(WordDocument)

        With WordDocumentBuilder.Font
            .Name = "微软雅黑"
            .Size = 15
            .Bold = True
        End With

        WordDocumentBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center
        WordDocumentBuilder.Writeln("XXX检查报告")
        WordDocumentBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left
        WordDocumentBuilder.Font.Size = 12

        WordDocumentBuilder.InsertParagraph()

        WordDocumentBuilder.ParagraphFormat.FirstLineIndent = 15.0F
        WordDocumentBuilder.Writeln("报告基本信息：")

        WordDocumentBuilder.ParagraphFormat.FirstLineIndent = 0.0F
        Dim BaseInfoTable As Table = WordDocumentBuilder.StartTable()
        WordDocumentBuilder.CellFormat.VerticalAlignment = CellVerticalAlignment.Top
        WordDocumentBuilder.RowFormat.HeightRule = HeightRule.Auto
        WordDocumentBuilder.CellFormat.Borders.LineStyle = LineStyle.None

        WordDocumentBuilder.InsertCell()
        BaseInfoTable.AutoFit(AutoFitBehavior.FixedColumnWidths)
        WordDocumentBuilder.CellFormat.Width = 180
        WordDocumentBuilder.Font.Bold = False
        WordDocumentBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Right
        WordDocumentBuilder.Write("标签一：")

        WordDocumentBuilder.InsertCell()
        WordDocumentBuilder.CellFormat.Width = 260
        WordDocumentBuilder.Font.Bold = True
        WordDocumentBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Left
        WordDocumentBuilder.Write("数据一")
        WordDocumentBuilder.EndRow()

        WordDocumentBuilder.EndTable()

        WordDocument.Save(WordFilePath, SaveFormat.Docx)
    End Sub

End Class
