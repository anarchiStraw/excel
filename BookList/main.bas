Attribute VB_Name = "main"
Option Explicit

Const colIsbn = 1
Const colTitle = 2
Const colAuthor = 3
Const colCreators = 4
Const colManufacturer = 5
Const colPublicationDate = 6
Const colBinding = 7

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
    Dim errorNodes As MSXML2.IXMLDOMNodeList
    
    For i = r.Row To (r.Row + r.Rows.Count - 1)
'        If (progressDigit <= r.Rows.Count) Then
'            Call showProgress((i - r.Row + 1), r.Rows.Count)
'        End If
        asin = toAsin(ws.Cells(i, colIsbn))
        If (asin <> "") Then
            Set xdoc = load(signedUrlFor(asin))
            If (Not xdoc.SelectSingleNode("/ItemLookupResponse/Items/Request/Errors") Is Nothing) Then
                ws.Cells(i, colIsbn).Select
                With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent3
                        .TintAndShade = 0.799981688894314
                        .PatternTintAndShade = 0
                End With
                
                MsgBox (i & "行 データ取得できませんでした。理由：" & vbLf _
                        & xdoc.SelectSingleNode("/ItemLookupResponse/Items/Request/Errors/Error[0]/Message").text)
                GoTo NEXT_ROW
            End If
            
            Set itemAttributes = xdoc.SelectSingleNode("/ItemLookupResponse/Items/Item/ItemAttributes")
            
            ws.Cells(i, colAuthor).Value = itemAttributes.SelectSingleNode("Author[0]").text
            ws.Cells(i, colManufacturer).Value = itemAttributes.SelectSingleNode("Manufacturer").text
            ws.Cells(i, colTitle).Value = itemAttributes.SelectSingleNode("Title").text
            ws.Cells(i, colPublicationDate).Value = itemAttributes.SelectSingleNode("PublicationDate").text
            ws.Cells(i, colBinding).Value = itemAttributes.SelectSingleNode("Binding").text
            
            Dim n As MSXML2.IXMLDOMNode
            Dim creators As String
            For Each n In itemAttributes.SelectNodes("Creator")
                creators = creators & n.text & "(" & n.Attributes.getNamedItem("Role").text & "), "
            Next
            If (0 < Len(creators)) Then
                ws.Cells(i, colCreators).Value = Left(creators, Len(creators) - 2) ' 最後のカンマとスペース不要
            End If
            ws.Cells(i, colIsbn).Select
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Else
            ws.Cells(i, colIsbn).Select
            With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
            End With
            
            MsgBox (i & "行" & " ISBNが正しく入力されていないようです。飛ばします。")
            GoTo NEXT_ROW
        End If
NEXT_ROW:
    Next

    Application.StatusBar = False
End Sub

Function showProgress(current As Integer, all As Integer)
    Dim progress As Integer
    progress = WorksheetFunction.Round(progressDigit * (current / all), 0)
    Application.StatusBar = "処理中(" & current & "/" & all & ") " _
        & WorksheetFunction.Rept("|", progress) _
        & WorksheetFunction.Rept("-", (progressDigit - progress))
    
End Function

Private Function toAsin(isbn As String) As String
    isbn = Replace(Trim(isbn), "-", "")
    
    Select Case Len(isbn)
    Case 10
        If (Val(Left(isbn, 9)) = 0) Then
            toAsin = ""
            Exit Function
        End If
        
        toAsin = isbn
    Case 13
        isbn = Mid(isbn, 4, 9)
        
        Dim sum As Integer
        Dim i As Integer
        For i = 1 To 9
            sum = sum + Mid(isbn, i, 1) * (11 - i)
        Next
        
        Dim checkDigit As String
        Select Case (sum Mod 11)
        Case 0
            checkDigit = "0"
        Case 1
            checkDigit = "X"
        Case Else
            checkDigit = CStr(11 - (sum Mod 11))
        End Select
        
        toAsin = isbn & checkDigit
    Case Else
        toAsin = ""
    End Select
End Function

Function testAsin()
    Debug.Assert (asin("") = "")
    Debug.Assert (asin("a123-45-6789") = "")
    Debug.Assert (asin("4-86011-202-4") = "4860112024")
    Debug.Assert (asin("978-4-86011-202-8") = "4860112024")
End Function

Private Function signedUrlFor(asin As String) As String
    Dim accessKey As String
    Dim secretKey As String
    Dim associateTag As String
    Dim host As String
    Dim path As String
    Dim params As String
    Dim stringToSign As String
    Dim signedUrl As String
    
    accessKey = "accessKey"
    secretKey = "secretKey"
    associateTag = "associateTag"
    
    host = "ecs.amazonaws.jp"
    path = "/onca/xml"
    params = "AWSAccessKeyId=" & accessKey _
        & "&AssociateTag=" & associateTag _
        & "&ItemId=" & asin _
        & "&Operation=ItemLookup" _
        & "&ResponseGroup=ItemAttributes" _
        & "&Service=AWSECommerceService" _
        & "&Timestamp=" & urlEncode(Format(Now, "yyyy-mm-ddThh:MM:ss+0900")) _
        & "&Version=2009-03-31"
    
    stringToSign = "GET" & vbLf & host & vbLf & path & vbLf & params
    signedUrlFor = "http://" & host & path & "?" & params & "&Signature=" & getSignature(stringToSign, secretKey)
    Debug.Print signedUrlFor

End Function

Function testSignedUrlFor()
    Debug.Assert (signedUrlFor("4860112024") = "http://ecs.amazonaws.jp/onca/xml?AWSAccessKeyId=AKIAIL7NZCKP32A32LQQ&AssociateTag=attentiveada-20&ItemId=4860112024&Operation=ItemLookup&ResponseGroup=ItemAttributes&Service=AWSECommerceService&Timestamp=2011-11-30T11%3A12%3A33%2B0900&Version=2009-03-31&Signature=loqn3GV5ASf4HTLMxUh2GhGtQYAIJIYN21k6bykNT3A%3D")
End Function

Function getSignature(stringToSign As String, secretKey As String) As String
    Dim i As Integer
    Dim hash As String
    Dim arKey() As Byte
    Dim ipad As String
    Dim opad As String
    Dim buff() As Byte, offset As Integer
    
    '初期化
    ipad = ""
    opad = ""
    ReDim arKey(0 To 63)
    
    '秘密鍵から1文字づつ読込み、文字コードへ変換後配列へ格納
    For i = 0 To Len(secretKey) - 1
        arKey(i) = Asc(Mid(secretKey, i + 1, 1))
    Next
    
    '64文字に満たない分は、ゼロセット
    For i = Len(secretKey) To 63
        arKey(i) = 0
    Next
    
    'innerpad及びouterpad作成
    For i = 0 To 63
        ipad = ipad & Chr(arKey(i) Xor &H36)
        opad = opad & Chr(arKey(i) Xor &H5C)
    Next
    
    'ハッシュ処理1回目
    '(innerpad+メッセージ文字列)をハッシュ・・・ハッシュ結果1
    hash = CreateSHA256HashString(ipad & stringToSign)
    
    'ハッシュ処理2回目(modify by YU-TANGさん)
    '(outerpad+ハッシュ結果1)をハッシュ・・・メッセージ認証コード作成完了
    buff = StrConv(opad, vbFromUnicode)
    offset = UBound(buff)
    ReDim Preserve buff(offset + Len(hash) / 2)
    
    For i = 1 To (Len(hash) \ 2)
        buff(offset + i) = CByte("&H" & Mid(hash, (i - 1) * 2 + 1, 2))
    Next
    hash = CreateSHA256Hash(buff)
    
    getSignature = urlEncode(EncodeBase64(hex2byte(hash)))
    
End Function

Private Function hex2byte(hexStr As String) As Byte()
    Dim buff() As Byte
    
    Dim offset As Integer
    offset = 0
    ReDim Preserve buff(offset + (Len(hexStr) / 2) - 1)
    
    Dim i As Integer
    For i = 0 To (Len(hexStr) / 2) - 1
        buff(offset + i) = Val("&H" & Mid(hexStr, (i) * 2 + 1, 2))
    Next
    hex2byte = buff
End Function

Private Function EncodeBase64(src() As Byte) As String
  Dim objXML As MSXML2.DOMDocument
  Dim objNode As MSXML2.IXMLDOMElement

  Set objXML = New MSXML2.DOMDocument
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = src
  EncodeBase64 = objNode.text

  Set objNode = Nothing
  Set objXML = Nothing
End Function

Function urlEncode(str As String) As String
    Dim sc, js As Variant
    Set sc = CreateObject("ScriptControl")
    sc.Language = "Jscript"
    Set js = sc.CodeObject
    urlEncode = js.encodeURIComponent(str)
End Function

Function load(url As String) As MSXML2.DOMDocument
    Dim xdoc As MSXML2.DOMDocument
    Set xdoc = New MSXML2.DOMDocument
    
    xdoc.async = False
    'エラー対策 http://support.microsoft.com/kb/281142/ja
    xdoc.setProperty "ServerHTTPRequest", True
 
    xdoc.load (url)
    Set load = xdoc
End Function
