Attribute VB_Name = "main"
Option Explicit

' �f�[�^���������ޗ�ԍ��B
' Excel�V�[�g�̃��C�A�E�g��ς�����A���̒l�����킹�ĕς���K�v������܂��B
Const colIsbn = 1
Const colTitle = 2
Const colAuthor = 3
Const colCreators = 4
Const colPublisher = 5
Const colPublicationDate = 6
Const colBinding = 7

' �X�e�[�^�X�o�[�ɕ\������i���󋵂̌�
Const progressDigit = 20

Public Sub setBookInfo()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim r As Range
    Set r = Selection
        
    Dim i As Integer
    Dim asin As String
    Dim xdoc As MSXML2.DOMDocument
    Dim itemAttributes As MSXML2.IXMLDOMNode
    
    For i = r.Row To (r.Row + r.Rows.Count - 1)
        If (progressDigit <= r.Rows.Count) Then ' ���Ȃ������Ȃ�킴�킴�\�����Ȃ�
            Call showProgress((i - r.Row + 1), r.Rows.Count)
        End If
        
        asin = toAsin(ws.Cells(i, colIsbn))
        If (asin = "") Then
            Call bgColor(ws.Cells(i, colIsbn), xlThemeColorAccent6)
            MsgBox ("�s [" & i & "] ISBN�����������͂���Ă��Ȃ��悤�ł��B��΂��܂��B")
            GoTo NEXT_ROW
        Else
            Dim maps() As Variant
            On Error GoTo ERROR_HANDLE
            maps = getAttributeMaps(load(signedUrlFor(asin:=asin)))
            On Error GoTo 0
            
            ws.Cells(i, colTitle).Value = maps(0)("title")
            ws.Cells(i, colAuthor).Value = maps(0)("author")
            ws.Cells(i, colCreators).Value = maps(0)("creators")
            ws.Cells(i, colPublisher).Value = maps(0)("publisher")
            ws.Cells(i, colPublicationDate).Value = maps(0)("publicationDate")
            ws.Cells(i, colBinding).Value = maps(0)("binding")
            
            Call bgColor(ws.Cells(i, colIsbn), Null)
        End If
NEXT_ROW:
    Next
    Exit Sub

ERROR_HANDLE:
    If Err.Number = 500 Then
        MsgBox ("�s [" & i & "] �f�[�^�擾�ł��܂���ł����B���R�F" & vbLf & Err.description)
        Call bgColor(ws.Cells(i, colIsbn), xlThemeColorAccent3)
        GoTo NEXT_ROW
    End If
    Error Err
End Sub

Public Sub searchBookInfo()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim r As Range
    Set r = Selection
        
    ' �^�C�g���`�o�ŎЁA���͂���Ă���������ɑ���
    Dim strTitle As String
    Dim strAuthor As String
    Dim strPublisher As String
    strTitle = Trim(ws.Cells(r.Row, colTitle).Value)
    strAuthor = Trim(ws.Cells(r.Row, colAuthor).Value)
    strPublisher = Trim(ws.Cells(r.Row, colPublisher).Value)
    If (strTitle = "" And strAuthor = "" And strPublisher = "") Then
        MsgBox ("�^�C�g���A��ҁA�o�Ŏ� �����ꂩ�͓��͂��Ă��������B")
        Exit Sub
    End If
    
    Dim maps() As Variant
    On Error GoTo ERROR_HANDLE
    maps = getAttributeMaps(load(signedUrlFor(title:=strTitle, author:=strAuthor, publisher:=strPublisher)))
    On Error GoTo 0
    
    Call searchResult.initialize(title:=strTitle, author:=strAuthor, publisher:=strPublisher, results:=maps)
    searchResult.Show
    If (searchResult.Tag = "cancel") Then
        Unload searchResult
        Exit Sub
    End If
    ws.Cells(r.Row, colIsbn).Value = maps(searchResult.Tag)("ean")
    ws.Cells(r.Row, colTitle).Value = maps(searchResult.Tag)("title")
    ws.Cells(r.Row, colAuthor).Value = maps(searchResult.Tag)("author")
    ws.Cells(r.Row, colCreators).Value = maps(searchResult.Tag)("creators")
    ws.Cells(r.Row, colPublisher).Value = maps(searchResult.Tag)("publisher")
    ws.Cells(r.Row, colPublicationDate).Value = maps(searchResult.Tag)("publicationDate")
    ws.Cells(r.Row, colBinding).Value = maps(searchResult.Tag)("binding")
    Call bgColor(ws.Cells(r.Row, colIsbn), Null)
    Unload searchResult
    Exit Sub

ERROR_HANDLE:
    If Err.Number = 500 Then
        MsgBox ("�f�[�^�擾�ł��܂���ł����B���R�F" & vbLf & Err.description)
        Call bgColor(ws.Cells(r.Row, colIsbn), xlThemeColorAccent3)
        On Error GoTo 0
        Exit Sub
    End If
    Error Err
End Sub
