Attribute VB_Name = "main"
Option Explicit

' データを書き込む列番号。
' Excelシートのレイアウトを変えたら、この値も合わせて変える必要があります。
Const colIsbn = 1
Const colTitle = 2
Const colDirector = 3
Const colActors = 4
Const colPublisher = 5
Const colReleaseDate = 6
Const colBinding = 7

' ステータスバーに表示する進捗状況の桁
Const progressDigit = 20

Public Sub setBookInfo()
debugPrint "setBookInfo START -----------------"
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
        
        'asin = toAsin(ws.Cells(i, colIsbn))
        asin = ws.Cells(i, colIsbn)
        If (asin = "") Then
            Call bgColor(ws.Cells(i, colIsbn), xlThemeColorAccent6)
            MsgBox ("行 [" & i & "] ISBNが正しく入力されていないようです。飛ばします。")
            GoTo NEXT_ROW
        Else
            Dim maps() As Variant
            On Error GoTo ERROR_HANDLE
            maps = getAttributeMaps(load(signedUrlFor(asin:=asin)))
            On Error GoTo 0
debugPrint "creating attribute map done."

            pasteValues ws, i, maps(0)
debugPrint "pasteValues done."
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

Public Sub searchBookInfoFromAmazonCom(dummy As Integer)
    searchBookInfo (amazonCom)
End Sub

Public Sub searchBookInfoFromAmazonFr(dummy As Integer)
    searchBookInfo (amazonFr)
End Sub

Public Sub searchBookInfoFromAmazonEs(dummy As Integer)
    searchBookInfo (amazonEs)
End Sub

Public Sub searchBookInfo(Optional endpoint As Variant)
    If IsMissing(endpoint) Then
        endpoint = amazonJp
    End If
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim r As Range
    Set r = Selection
        
    ' タイトル～出版社、入力されていたら条件に足す
    Dim strTitle As String
    Dim strDirector As String
    Dim strActor As String
    strTitle = Trim(ws.Cells(r.row, colTitle).Value)
    strDirector = Trim(ws.Cells(r.row, colDirector).Value)
    strActor = Trim(ws.Cells(r.row, colPublisher).Value)
    If (strTitle = "" And strDirector = "" And strActor = "") Then
        MsgBox ("タイトル、作者、出版社 いずれかは入力してください。")
        Exit Sub
    End If
    
    Dim maps() As Variant
    On Error GoTo ERROR_HANDLE
    maps = getAttributeMaps(load(signedUrlFor(endpoint:=endpoint, title:=strTitle, director:=strDirector, actor:=strActor)))
    On Error GoTo 0
    
    Call searchResult.initialize(title:=strTitle, director:=strDirector, actor:=strActor, results:=maps)
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
    ws.Cells(row, colDirector).Value = map("director")
    ws.Cells(row, colActors).Value = map("actors")
    ws.Cells(row, colPublisher).Value = map("publisher")
    Dim releaseDate As String
    ' 年4ケタのみ、などの場合、Excelが「日付値」と勘違いするのでハイフンをくっつける
    releaseDate = map("releaseDate")
    ws.Cells(row, colReleaseDate).Value = IIf(isNumber(releaseDate), releaseDate & "-", releaseDate)
    ws.Cells(row, colBinding).Value = map("binding")

    
    Call bgColor(ws.Cells(row, colIsbn), Null)
End Function

Private Function isNumber(var As String) As Boolean
    On Error GoTo NOT_A_NUMBER
    Dim i As Integer
    i = CInt(var)
    isNumber = True
    Exit Function
NOT_A_NUMBER:
    On Error GoTo 0
    isNumber = False
End Function

Public Sub autoExecSetBookInfo()
    isbnInputForm.Show
End Sub
