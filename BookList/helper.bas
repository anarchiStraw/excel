Attribute VB_Name = "helper"
Option Explicit

Const DEBUG_LOG = False
Function debugPrint(message As String)
    If DEBUG_LOG Then
        Debug.Print message
    End If
End Function

Function showProgress(current As Integer, all As Integer)
    Dim progress As Integer
    Dim progressDigit As Integer
    progressDigit = 10
    progress = WorksheetFunction.Round(progressDigit * (current / all), 0)
    Application.StatusBar = "処理中(" & current & "/" & all & ") " _
        & WorksheetFunction.Rept("|", progress) _
        & WorksheetFunction.Rept("-", (progressDigit - progress))
End Function

Function toAsin(isbn As String) As String
    isbn = Replace(Trim(isbn), "-", "")
    
    Select Case Len(isbn)
    Case 10
        toAsin = IIf((Val(Left(isbn, 9)) = 0), "", isbn)
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

' Amazonデータ取得URL生成
' asin または (title, author, publisher) を指定する。
' asinが指定されればItemLookup、指定されなければItemSearch のURLを返す。
'
' accessKey以降はテスト用。
' 実際使うときは引数ではなくfunction内の yourAccessKey, yourSecretKey, yourAssociateTag を正しい値に書き換えてください。
Function signedUrlFor( _
        Optional endpoint As Variant, _
        Optional asin As Variant, _
        Optional title As Variant, Optional author As Variant, Optional publisher As Variant, _
        Optional accessKey As Variant, Optional secretKey As Variant, Optional associateTag As Variant, Optional timestamp As Variant) As String
    If IsMissing(endpoint) Then
        endpoint = endpoints(IIf(IsMissing(asin), "4", CStr(asin)))
    End If

    Dim path As String
    path = "/onca/xml"
    
    Dim params As String
    params = "AWSAccessKeyId=" & IIf(IsMissing(accessKey), "yourAccessKey", accessKey) _
        & "&AssociateTag=" & IIf(IsMissing(associateTag), "yourAssociateTag", associateTag) _
        & IIf(IsMissing(author), "", "&Author=" & urlEncode(CStr(author))) _
        & IIf(IsMissing(asin), "", "&ItemId=" & CStr(asin)) _
        & "&Operation=" & IIf(IsMissing(asin), "ItemSearch", "ItemLookup") _
        & IIf(IsMissing(publisher), "", "&Publisher=" & urlEncode(CStr(publisher))) _
        & "&ResponseGroup=ItemAttributes" _
        & IIf(IsMissing(asin), "&SearchIndex=Books", "") _
        & "&Service=AWSECommerceService" _
        & "&Timestamp=" & urlEncode(IIf(IsMissing(timestamp), Format(Now, "yyyy-mm-ddThh:MM:ss+0900"), timestamp)) _
        & IIf(IsMissing(title), "", "&Title=" & urlEncode(CStr(title))) _
        & "&Version=2011-08-01"
    
    Dim stringToSign As String
    stringToSign = "GET" & vbLf & endpoint & vbLf & path & vbLf & params
    
    signedUrlFor = "http://" & endpoint & path & "?" & params _
                & "&Signature=" & getSignature(stringToSign, IIf(IsMissing(secretKey), "yourSecretKey", secretKey))
    Debug.Print signedUrlFor

End Function

Function endpoints(asin As String) As String
    Dim countryNumber As Integer
    countryNumber = CInt(IIf(CInt(Left(asin, 1)) <= 7, Left(asin, 1), Left(asin, 2)))
    Select Case countryNumber
    Case 2
        endpoints = amazonFr
    Case 3
        endpoints = amazonDe
    Case 4
        endpoints = amazonJp
    Case 7
        endpoints = amazonCn
    Case 84
        endpoints = amazonEs
    Case 88
        endpoints = amazonIt
    Case Else
        endpoints = amazonCom
    End Select
End Function

Function amazonFr() As String
    amazonFr = "ecs.amazonaws.fr"
End Function

Function amazonDe() As String
    amazonDe = "ecs.amazonaws.de"
End Function

Function amazonJp() As String
    amazonJp = "ecs.amazonaws.jp"
End Function

Function amazonCn() As String
    amazonCn = "webservices.amazon.cn"
End Function

Function amazonEs() As String
    amazonEs = "webservices.amazon.es"
End Function

Function amazonIt() As String
    amazonIt = "webservices.amazon.it"
End Function

Function amazonCom() As String
    amazonCom = "webservices.amazon.com"
End Function

' この関数はプラプラさんが
' http://plus-sys.jugem.jp/?eid=220
' で公開されているものを、ほぼそのまま使いました。
' (対象文字列と秘密鍵を引数にし、結果をDebug.printでなく返り値にしました)
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

Function hex2byte(hexStr As String) As Byte()
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

Function EncodeBase64(src() As Byte) As String
  Dim objXML As MSXML2.DOMDocument
  Dim objNode As MSXML2.IXMLDOMElement

  Set objXML = New MSXML2.DOMDocument
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = src
  EncodeBase64 = objNode.Text

  Set objNode = Nothing
  Set objXML = Nothing
End Function

Function urlEncode(str As String) As String
    Dim sc As Variant
    Set sc = CreateObject("ScriptControl")
    sc.Language = "Jscript"
    urlEncode = Replace(Replace(Replace(Replace(sc.CodeObject.encodeURIComponent(str), "!", "%21"), "(", "%28"), ")", "%29"), "*", "%2A")
End Function

Function load(url As String) As MSXML2.DOMDocument
    Dim xdoc As MSXML2.DOMDocument
    Set xdoc = New MSXML2.DOMDocument
    Dim tried As Integer
    tried = 0
    
    xdoc.async = False
    'エラー対策 http://support.microsoft.com/kb/281142/ja
    xdoc.setProperty "ServerHTTPRequest", True
    Do
        If (tried > 0) Then
            debugPrint "trying " & tried
        End If
        xdoc.load (url)
        tried = tried + 1
    Loop While (xdoc.XML = "" And tried < 3)
    If xdoc.XML = "" Then
        Err.Raise Number:=500, Description:="XMLを取得できませんでした(再実行すれば取得できるかもしれません)。"
    End If
    Set load = xdoc
End Function

Function getAttributeMaps(xdoc As MSXML2.DOMDocument) As Variant

    If (0 < InStr(1, xdoc.SelectSingleNode("/*/Items/Request").xml, "<Error")) Then
        Dim code, message As String
        code = xdoc.SelectSingleNode("/*/Items/Request/Errors/Error[0]/Code").Text
        If (0 < InStr(1, code, "AWS.ECommerceService.NoExactMatches")) Then
            message = "検索結果がありません。"
        ElseIf (0 < InStr(1, code, "AWS.InvalidParameterValue")) Then
            message = "このISBNは正しくないか、Amazonに登録されていません。"
        Else
            message = xdoc.SelectSingleNode("/*/Items/Request/Errors/Error[0]/Message").Text
        End If
        Err.Raise Number:=500, Description:=message
    End If
    
    Dim itemNodes As MSXML2.IXMLDOMNodeList
    Set itemNodes = xdoc.SelectNodes("/*/Items/Item")
    
    Dim maps() As Variant
    ReDim maps(itemNodes.Length - 1)
    
    Dim itemNode As MSXML2.IXMLDOMNode
    Dim i As Integer
    For i = 0 To (itemNodes.Length - 1)
        Dim attributesNode As MSXML2.IXMLDOMNode
        Set attributesNode = itemNodes(i).SelectSingleNode("ItemAttributes")
        
        On Error Resume Next ' ノードが欠落している場合のエラーを無視
        
        Dim map As Object
        Set map = CreateObject("Scripting.Dictionary")
        
        map.Add "title", attributesNode.SelectSingleNode("Title").Text
        map.Add "author", attributesNode.SelectSingleNode("Author[0]").Text
        Dim creators As String
        creators = ""
        Dim n As MSXML2.IXMLDOMNode
        For Each n In attributesNode.SelectNodes("Creator")
            creators = creators & n.Text & "(" & n.attributes.getNamedItem("Role").Text & "), "
        Next
        If (0 < Len(creators)) Then
            creators = Left(creators, (Len(creators) - 2)) ' 最後のカンマとスペース不要
        End If
        map.Add "creators", creators
        
        map.Add "publisher", attributesNode.SelectSingleNode("Publisher").Text
        map.Add "publicationDate", attributesNode.SelectSingleNode("PublicationDate").Text
        map.Add "binding", attributesNode.SelectSingleNode("Binding").Text
        map.Add "ean", attributesNode.SelectSingleNode("EAN").Text
        
        On Error GoTo 0
        Set maps(i) = map
    Next
    getAttributeMaps = maps
End Function

Function bgColor(r As Range, color As Variant)
    r.Select
    With Selection.Interior
        If IsEmpty(color) Then
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        Else
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = color
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End If
    End With

End Function
