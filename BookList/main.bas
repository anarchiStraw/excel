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
Const colNote = 8
Const colPages = 9
Const colCurrencyCode = 10
Const colListPrice = 11
Const colLowestNewPrice = 12
Const colLowestUsedPrice = 13
Const colLowestCollectiblePrice = 14
Const colSalesRank = 15

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
    
    For i = r.row To (r.row + r.Rows.Count - 1)
        If (progressDigit <= r.Rows.Count) Then ' 少ない件数ならわざわざ表示しない
            Call showProgress((i - r.row + 1), r.Rows.Count)
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
            
            pasteValues ws, i, maps(0)
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
    strTitle = Trim(ws.Cells(r.row, colTitle).Value)
    strAuthor = Trim(ws.Cells(r.row, colAuthor).Value)
    strPublisher = Trim(ws.Cells(r.row, colPublisher).Value)
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
    pasteValues ws, r.row, maps(searchResult.Tag)
    Unload searchResult
    Exit Sub

ERROR_HANDLE:
    If Err.Number = 500 Then
        MsgBox ("データ取得できませんでした。理由：" & vbLf & Err.description)
        Call bgColor(ws.Cells(r.row, colIsbn), xlThemeColorAccent3)
        On Error GoTo 0
        Exit Sub
    End If
    Error Err
End Sub

Private Function pasteValues(ws As Worksheet, row As Integer, map As Variant)
    ws.Cells(row, colIsbn).Value = map("ean")
    ws.Cells(row, colTitle).Value = map("title")
    ws.Cells(row, colAuthor).Value = map("author")
    ws.Cells(row, colCreators).Value = map("creators")
    ws.Cells(row, colPublisher).Value = map("publisher")
    ws.Cells(row, colPublicationDate).Value = map("publicationDate")
    ws.Cells(row, colBinding).Value = map("binding")

    ws.Cells(row, colPages).Value = map("pages")
    ws.Cells(row, colCurrencyCode).Value = map("currencyCode")
    ws.Cells(row, colListPrice).Value = map("listPrice")
    ws.Cells(row, colLowestNewPrice).Value = map("lowestNewPrice")
    ws.Cells(row, colLowestUsedPrice).Value = map("lowestUsedPrice")
    ws.Cells(row, colLowestCollectiblePrice).Value = map("lowestCollectiblePrice")
    ws.Cells(row, colSalesRank).Value = map("salesRank")
    
    Call bgColor(ws.Cells(row, colIsbn), Null)
End Function
