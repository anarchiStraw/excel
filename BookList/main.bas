Attribute VB_Name = "main"
Option Explicit

' データを書き込む列番号。
' Excelシートのレイアウトを変えたら、この値も合わせて変える必要があります。
Const colIsbn = 1
Const colTitle = 2
Const colAuthor = 3
Const colCreators = 4
Const colPublisher = 5
Const colPublicationDate = 6
Const colBinding = 7

' ステータスバーに表示する進捗状況の桁
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
        If (progressDigit <= r.Rows.Count) Then ' 少ない件数ならわざわざ表示しない
            Call showProgress((i - r.Row + 1), r.Rows.Count)
        End If
        
        asin = toAsin(ws.Cells(i, colIsbn))
        If (asin = "") Then
            Call bgColor(ws.Cells(i, colIsbn), xlThemeColorAccent6)
            MsgBox ("行 [" & i & "] ISBNが正しく入力されていないようです。飛ばします。")
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
        MsgBox ("行 [" & i & "] データ取得できませんでした。理由：" & vbLf & Err.description)
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
        
    ' タイトル～出版社、入力されていたら条件に足す
    Dim strTitle As String
    Dim strAuthor As String
    Dim strPublisher As String
    strTitle = Trim(ws.Cells(r.Row, colTitle).Value)
    strAuthor = Trim(ws.Cells(r.Row, colAuthor).Value)
    strPublisher = Trim(ws.Cells(r.Row, colPublisher).Value)
    If (strTitle = "" And strAuthor = "" And strPublisher = "") Then
        MsgBox ("タイトル、作者、出版社 いずれかは入力してください。")
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
        MsgBox ("データ取得できませんでした。理由：" & vbLf & Err.description)
        Call bgColor(ws.Cells(r.Row, colIsbn), xlThemeColorAccent3)
        On Error GoTo 0
        Exit Sub
    End If
    Error Err
End Sub
