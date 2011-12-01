Attribute VB_Name = "main"
Option Explicit

' �f�[�^���������ޗ�ԍ��B
' Excel�V�[�g�̃��C�A�E�g��ς�����A���̒l�����킹�ĕς���K�v������܂��B
Const colIsbn = 1
Const colTitle = 2
Const colAuthor = 3
Const colCreators = 4
Const colManufacturer = 5
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
    Dim errorNodes As MSXML2.IXMLDOMNodeList
    
    For i = r.Row To (r.Row + r.Rows.Count - 1)
        If (progressDigit <= r.Rows.Count) Then ' ���Ȃ������Ȃ�킴�킴�\�����Ȃ�
            Call showProgress((i - r.Row + 1), r.Rows.Count)
        End If
        
        asin = toAsin(ws.Cells(i, colIsbn))
        If (asin = "") Then
            Call bgColor(ws.Cells(i, colIsbn), xlThemeColorAccent6)
            MsgBox (i & "�s" & " ISBN�����������͂���Ă��Ȃ��悤�ł��B��΂��܂��B")
            GoTo NEXT_ROW
        Else
            Set xdoc = load(signedUrlFor(asin))
            If (Not xdoc.SelectSingleNode("/ItemLookupResponse/Items/Request/Errors") Is Nothing) Then
                Call bgColor(ws.Cells(i, colIsbn), xlThemeColorAccent3)
                
                MsgBox (i & "�s �f�[�^�擾�ł��܂���ł����B���R�F" & vbLf _
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
                ws.Cells(i, colCreators).Value = Left(creators, Len(creators) - 2) ' �Ō�̃J���}�ƃX�y�[�X�s�v
                creators = ""
            End If
            Call bgColor(ws.Cells(i, colIsbn), Null)
        End If
NEXT_ROW:
    Next

    Application.StatusBar = False
End Sub

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

' asin�ɑ΂���f�[�^�擾URL
' optional�����̓e�X�g�p�B
' ���ێg���Ƃ��� yourAccessKey, yourSecretKey, yourAssociateTag �𐳂����l�ɏ��������Ă��������B
Function signedUrlFor(asin As String, _
        Optional accessKey As Variant, Optional secretKey As Variant, _
        Optional associateTag As Variant, Optional timestamp As Variant) As String
    
    Dim endpoint As String
    endpoint = "ecs.amazonaws.jp"
    
    Dim path As String
    path = "/onca/xml"
    
    Dim params As String
    params = "AWSAccessKeyId=" & IIf(IsMissing(accessKey), "yourAccessKey", accessKey) _
        & "&AssociateTag=" & IIf(IsMissing(associateTag), "yourAssociateTag", associateTag) _
        & "&ItemId=" & asin _
        & "&Operation=ItemLookup" _
        & "&ResponseGroup=ItemAttributes" _
        & "&Service=AWSECommerceService" _
        & "&Timestamp=" & urlEncode(IIf(IsMissing(timestamp), Format(Now, "yyyy-mm-ddThh:MM:ss+0900"), timestamp)) _
        & "&Version=2011-08-01"
    
    Dim stringToSign As String
    stringToSign = "GET" & vbLf & endpoint & vbLf & path & vbLf & params
    
    signedUrlFor = "http://" & endpoint & path & "?" & params _
                & "&Signature=" & getSignature(stringToSign, IIf(IsMissing(secretKey), "yourSecretKey", secretKey))
    Debug.Print signedUrlFor

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
