Attribute VB_Name = "Module1"
Option Explicit

'    (c) 2019 Shoichi Hayashi(�� �ˈ�)
'    ���̃R�[�h��GPLv3�̌��Ƀ��C�Z���X���܂��
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
    
    ' ���Ԍ`���ł���XMLSS���o�͂���ꎞ�t�@�C���������߂�
    xmlssFileName = FSO.GetSpecialFolder(2) & "\" & FSO.GetTempName
    
    ' �ϊ�����FreeMind�`���̃t�@�C����I��������
    freemindFileName = getFilePath("USDM�L�ڃ��[���ɏ]����FreeMind�t�@�C��")
    If Right(freemindFileName, 3) = ".mm" Then
        ' �ϊ���̍ŏI�I��xlsx�t�@�C���������߂�
        xlsxFileName = Left(freemindFileName, Len(freemindFileName) - 3) & ".xlsx"
    Else
        MsgBox "�I�����ꂽ�t�@�C���ɖ�肪����̂ŁA�����𒆎~���܂��B" & vbLf & freemindFileName
        Exit Sub
    End If
    
    ' �ϊ�����FreeMind�`���̃t�@�C�����e��DOM�Ƃ��ēǂݍ���
    xdoc.Load (freemindFileName)
    
    ' �V�[�g�ɕۑ����Ă���ϊ��p��XSL�\�[�X�𕶎���ϐ��ɓǂݍ���
    xslSrc = ThisWorkbook.Worksheets("�k��Ɨp�L���l").Range("XSL�\�[�X").Value
    
    ' XSL�\�[�X��DOM�ɕϊ�����
    If Not xstyle.LoadXML(xslSrc) Then
        MsgBox "xsl�̉�͂Ɏ��s���܂����B���e���m�F���Ă��������B�����𒆎~���܂��B"
        Exit Sub
    End If
    
    ' XSLT���g�p���ĕϊ�����FreeMind�`���̃t�@�C�����e��XMLSS�`���ɕϊ�����
    xmlss = xdoc.transformNode(xstyle)
    
    ' �ϊ����ꂽXMLSS���Ɍ��߂Ă��������O�̈ꎞ�t�@�C���ɏo�͂���
    Call OutputFile(xmlss, xmlssFileName)
    
    ' �s�v�ɂȂ���DOM�̈�����
    xdoc.abort
    xstyle.abort

    ' �ϊ����ꂽXMLSS�̈ꎞ�t�@�C����V����Excel Book�Ƃ��ĊJ��
    Set wb = Workbooks.Open(xmlssFileName)
    
    ' �ϊ�����J���ꂽExcel Book��xlsx�`���ŕۑ�����
    On Error Resume Next
    Call wb.SaveAs(xlsxFileName, FileFormat:=xlOpenXMLWorkbook)
    If Err.Number > 0 Then MsgBox "�������O��t���ĕۑ�����Ă��܂���"

End Sub

Function OutputFile(xmlss As String, fileName As String) As Boolean
    OutputFile = False
    Dim Reader As New SAXXMLReader60
    Dim writer As New MXXMLWriter60
    Dim xdoc As New MSXML2.DOMDocument60

    ' writer.indent = True
    writer.indent = False ' �C���f���g�̃X�y�[�X���Ȃ���Excel�̃Z�����̒l�ɓ����Ă��܂��̂�
    writer.standalone = True
    writer.Encoding = "shift_jis"
    'writer.Encoding = "UTF-8"
    writer.omitXMLDeclaration = True ' XML�錾�͕����R�[�h�ɂ��Đ����I�Ȃ��̂��o�͂���Ȃ��̂ŁA�p�ӂ���������𒼐ڃX�g���[���ɏo�͂���
    Set Reader.contentHandler = writer
    Call Reader.putProperty("http://xml.org/sax/properties/lexical-handler", writer)

    'ADODB.Stream�I�u�W�F�N�g�𐶐�
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream") ' UTF8�ւ̕ϊ��̂��߂Ɏg��
    
    Reader.Parse xmlss
    '�C���f���g����XML��ǂݍ��݂Ȃ���
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
        ' BOM���폜����
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

' ���[�U��.mm�t�@�C����I�������āA���̃t���p�X�𓾂�
Function getFilePath(title As String) As String
    Dim fileDialog As fileDialog
    
    Set fileDialog = Application.fileDialog(msoFileDialogOpen)
    fileDialog.AllowMultiSelect = False
    fileDialog.title = title
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "FreeMind�t�@�C��", "*.mm"
    If fileDialog.Show = -1 Then
        getFilePath = fileDialog.SelectedItems(1)
    End If
End Function

