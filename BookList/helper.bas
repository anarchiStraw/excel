Attribute VB_Name = "helper"
Option Explicit

Function showProgress(current As Integer, all As Integer)
    Dim progress As Integer
    progress = WorksheetFunction.Round(progressDigit * (current / all), 0)
    Application.StatusBar = "������(" & current & "/" & all & ") " _
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

' Amazon�f�[�^�擾URL����
' asin �܂��� (title, author, publisher) ���w�肷��B
' asin���w�肳����ItemLookup�A�w�肳��Ȃ����ItemSearch ��URL��Ԃ��B
'
' accessKey�ȍ~�̓e�X�g�p�B
' ���ێg���Ƃ��͈����ł͂Ȃ�function���� yourAccessKey, yourSecretKey, yourAssociateTag �𐳂����l�ɏ��������Ă��������B
Function signedUrlFor( _
        Optional asin As Variant, _
        Optional title As Variant, Optional author As Variant, Optional publisher As Variant, _
        Optional accessKey As Variant, Optional secretKey As Variant, Optional associateTag As Variant, Optional timestamp As Variant) As String
    Dim endpoint As String
    endpoint = endpoints(IIf(IsMissing(asin), "4", CStr(asin)))
    
    Dim path As String
    path = "/onca/xml"
    
    Dim params As String
    params = "AWSAccessKeyId=" & IIf(IsMissing(accessKey), "yourAccessKey", accessKey) _
        & "&AssociateTag=" & IIf(IsMissing(associateTag), "yourAssociateTag", associateTag) _
        & IIf(IsMissing(author), "", "&Author=" & urlEncode(removeUnwantedChars(CStr(author)))) _
        & IIf(IsMissing(asin), "", "&ItemId=" & CStr(asin)) _
        & "&Operation=" & IIf(IsMissing(asin), "ItemSearch", "ItemLookup") _
        & IIf(IsMissing(publisher), "", "&Publisher=" & urlEncode(removeUnwantedChars(CStr(publisher)))) _
        & "&ResponseGroup=ItemAttributes" _
        & IIf(IsMissing(asin), "&SearchIndex=Books", "") _
        & "&Service=AWSECommerceService" _
        & "&Timestamp=" & urlEncode(IIf(IsMissing(timestamp), Format(Now, "yyyy-mm-ddThh:MM:ss+0900"), timestamp)) _
        & IIf(IsMissing(title), "", "&Title=" & urlEncode(removeUnwantedChars(CStr(title)))) _
        & "&Version=2011-08-01"
    
    Dim stringToSign As String
    stringToSign = "GET" & vbLf & endpoint & vbLf & path & vbLf & params
    
    signedUrlFor = "http://" & endpoint & path & "?" & params _
                & "&Signature=" & getSignature(stringToSign, IIf(IsMissing(secretKey), "aIBZUSvQZ8Y15hTbHmZjAcTQDveTdBvaCv1uyBdj", secretKey))
    Debug.Print signedUrlFor

End Function

Function endpoints(asin As String) As String
    Dim countryNumber As Integer
    countryNumber = CInt(IIf(CInt(Left(asin, 1)) <= 7, Left(asin, 1), Left(asin, 2)))
    Select Case countryNumber
    Case 2
        endpoints = "ecs.amazonaws.fr"
    Case 3
        endpoints = "ecs.amazonaws.de"
    Case 4
        endpoints = "ecs.amazonaws.jp"
    Case 7
        endpoints = "webservices.amazon.cn"
    Case 84
        endpoints = "webservices.amazon.es"
    Case 88
        endpoints = "webservices.amazon.it"
    Case Else
        endpoints = "webservices.amazon.com"
    End Select
End Function

' ���̊֐��̓v���v������
' http://plus-sys.jugem.jp/?eid=220
' �Ō��J����Ă�����̂��A�قڂ��̂܂܎g���܂����B
' (�Ώە�����Ɣ閧���������ɂ��A���ʂ�Debug.print�łȂ��Ԃ�l�ɂ��܂���)
Function getSignature(stringToSign As String, secretKey As String) As String
    Dim i As Integer
    Dim hash As String
    Dim arKey() As Byte
    Dim ipad As String
    Dim opad As String
    Dim buff() As Byte, offset As Integer
    
    '������
    ipad = ""
    opad = ""
    ReDim arKey(0 To 63)
    
    '�閧������1�����ÂǍ��݁A�����R�[�h�֕ϊ���z��֊i�[
    For i = 0 To Len(secretKey) - 1
        arKey(i) = Asc(Mid(secretKey, i + 1, 1))
    Next
    
    '64�����ɖ����Ȃ����́A�[���Z�b�g
    For i = Len(secretKey) To 63
        arKey(i) = 0
    Next
    
    'innerpad�y��outerpad�쐬
    For i = 0 To 63
        ipad = ipad & Chr(arKey(i) Xor &H36)
        opad = opad & Chr(arKey(i) Xor &H5C)
    Next
    
    '�n�b�V������1���
    '(innerpad+���b�Z�[�W������)���n�b�V���E�E�E�n�b�V������1
    hash = CreateSHA256HashString(ipad & stringToSign)
    
    '�n�b�V������2���(modify by YU-TANG����)
    '(outerpad+�n�b�V������1)���n�b�V���E�E�E���b�Z�[�W�F�؃R�[�h�쐬����
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
  EncodeBase64 = objNode.text

  Set objNode = Nothing
  Set objXML = Nothing
End Function

Function urlEncode(str As String) As String
    Dim sc As Variant
    Set sc = CreateObject("ScriptControl")
    sc.Language = "Jscript"
    urlEncode = sc.CodeObject.encodeURIComponent(str)
End Function

Function load(url As String) As MSXML2.DOMDocument
    Dim xdoc As MSXML2.DOMDocument
    Set xdoc = New MSXML2.DOMDocument
    
    xdoc.async = False
    '�G���[�΍� http://support.microsoft.com/kb/281142/ja
    xdoc.setProperty "ServerHTTPRequest", True
 
    xdoc.load (url)
    Set load = xdoc
End Function

Function getAttributeMaps(xdoc As MSXML2.DOMDocument) As Variant

    If (0 < InStr(1, xdoc.SelectSingleNode("/*/Items/Request").xml, "<Error")) Then
        Dim code, message As String
        code = xdoc.SelectSingleNode("/*/Items/Request/Errors/Error[0]/Code").text
        If (0 < InStr(1, code, "AWS.ECommerceService.NoExactMatches")) Then
            message = "�������ʂ�����܂���B"
        ElseIf (0 < InStr(1, code, "AWS.InvalidParameterValue")) Then
            message = "����ISBN�͐������Ȃ����AAmazon�ɓo�^����Ă��܂���B"
        Else
            message = xdoc.SelectSingleNode("/*/Items/Request/Errors/Error[0]/Message").text
        End If
        Err.Raise Number:=500, description:=message
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
        
        On Error Resume Next ' �m�[�h���������Ă���ꍇ�̃G���[�𖳎�
        
        Dim map As Object
        Set map = CreateObject("Scripting.Dictionary")
        
        map.Add "title", attributesNode.SelectSingleNode("Title").text
        map.Add "author", attributesNode.SelectSingleNode("Author[0]").text
                
        Dim creators As String
        creators = ""
        Dim n As MSXML2.IXMLDOMNode
        For Each n In attributesNode.SelectNodes("Creator")
            creators = creators & n.text & "(" & n.attributes.getNamedItem("Role").text & "), "
        Next
        If (0 < Len(creators)) Then
            creators = Left(creators, (Len(creators) - 2)) ' �Ō�̃J���}�ƃX�y�[�X�s�v
        End If
        map.Add "creators", creators
        
        map.Add "publisher", attributesNode.SelectSingleNode("Publisher").text
        map.Add "publicationDate", attributesNode.SelectSingleNode("PublicationDate").text
        map.Add "binding", attributesNode.SelectSingleNode("Binding").text
        map.Add "ean", attributesNode.SelectSingleNode("EAN").text
        
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

Function removeUnwantedChars(str As String) As String
    Dim RE As Variant
    Set RE = CreateObject("VBScript.RegExp")
    With RE
        .Pattern = "[()*!']"
        .Global = True
    End With
    removeUnwantedChars = RE.Replace(str, "")
End Function