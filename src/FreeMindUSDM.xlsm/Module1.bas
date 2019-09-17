Attribute VB_Name = "Module1"
Option Explicit

'    (c) 2019 Shoichi Hayashi(林 祥一)
'    このコードはGPLv3の元にライセンスします｡
'    (http://www.gnu.org/copyleft/gpl.html)

Const XMLDeclaration As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbLf

Sub FreeMindUSDM()
    Dim xdoc As New MSXML2.DOMDocument60
    Dim xstyle As New MSXML2.DOMDocument60
    Dim xmlss As String
    Dim xslSrc As String
    Dim freemindFileName As String
    Dim xmlssFileName As String
    Dim xlsxFileName As String
    Dim wb As Workbook
    Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' 中間形式であるXMLSSを出力する一時ファイル名を決める
    xmlssFileName = FSO.GetSpecialFolder(2) & "\" & FSO.GetTempName
    
    ' 変換元のFreeMind形式のファイルを選択させる
    freemindFileName = getFilePath("USDM記載ルールに従ったFreeMindファイル")
    If Right(freemindFileName, 3) = ".mm" Then
        ' 変換後の最終的なxlsxファイル名を決める
        xlsxFileName = Left(freemindFileName, Len(freemindFileName) - 3) & ".xlsx"
    Else
        MsgBox "選択されたファイルに問題があるので、処理を中止します。" & vbLf & freemindFileName
        Exit Sub
    End If
    
    ' 変換元のFreeMind形式のファイル内容をDOMとして読み込む
    xdoc.Load (freemindFileName)
    
    ' シートに保存してある変換用のXSLソースを文字列変数に読み込む
    xslSrc = ThisWorkbook.Worksheets("〔作業用記憶〕").Range("XSLソース").Value
    
    ' XSLソースをDOMに変換する
    If Not xstyle.LoadXML(xslSrc) Then
        MsgBox "xslの解析に失敗しました。内容を確認してください。処理を中止します。"
        Exit Sub
    End If
    
    ' XSLTを使用して変換元のFreeMind形式のファイル内容をXMLSS形式に変換する
    xmlss = xdoc.transformNode(xstyle)
    
    ' 変換されたXMLSSを先に決めておいた名前の一時ファイルに出力する
    Call OutputFile(xmlss, xmlssFileName)
    
    ' 不要になったDOM領域を解放
    xdoc.abort
    xstyle.abort

    ' 変換されたXMLSSの一時ファイルを新しいExcel Bookとして開く
    Set wb = Workbooks.Open(xmlssFileName)
    
    ' 変換され開かれたExcel Bookをxlsx形式で保存する
    On Error Resume Next
    Call wb.SaveAs(xlsxFileName, FileFormat:=xlOpenXMLWorkbook)
    If Err.Number > 0 Then MsgBox "未だ名前を付けて保存されていません"

End Sub

Function OutputFile(xmlss As String, fileName As String) As Boolean
    OutputFile = False
    Dim Reader As New SAXXMLReader60
    Dim writer As New MXXMLWriter60
    Dim xdoc As New MSXML2.DOMDocument60

    ' writer.indent = True
    writer.indent = False ' インデントのスペースがなぜかExcelのセル内の値に入ってしまうので
    writer.standalone = True
    writer.Encoding = "shift_jis"
    'writer.Encoding = "UTF-8"
    writer.omitXMLDeclaration = True ' XML宣言は文字コードについて整合的なものが出力されないので、用意した文字列を直接ストリームに出力する
    Set Reader.contentHandler = writer
    Call Reader.putProperty("http://xml.org/sax/properties/lexical-handler", writer)

    'ADODB.Streamオブジェクトを生成
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream") ' UTF8への変換のために使う
    
    Reader.Parse xmlss
    'インデントしたXMLを読み込みなおす
    xdoc.LoadXML (writer.output)

    With adoSt
        .Type = adTypeText
        .Charset = "UTF-8"
        .LineSeparator = adLF
        .Open
        .WriteText XMLDeclaration
        .LineSeparator = adCRLF
        .WriteText Replace(writer.output, vbCrLf, vbLf)
        .LineSeparator = adLF
        ' BOMを削除する
        Dim byteData() As Byte
        .Position = 0
        .Type = adTypeBinary
        .Position = 3
        byteData = adoSt.Read
        .Close
        .Open
        .Write byteData
    
        .SaveToFile fileName, adSaveCreateOverWrite
        .Close
    End With
    
    OutputFile = True
End Function

' ユーザに.mmファイルを選択させて、そのフルパスを得る
Function getFilePath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "FreeMindファイル", "*.mm"
    If fileDialog.Show = -1 Then
        getFilePath = fileDialog.SelectedItems(1)
    End If
End Function

