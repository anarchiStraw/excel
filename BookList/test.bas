Attribute VB_Name = "test"
Option Explicit

Function testShowProgress()
    Call showProgress(1, 20)
    Debug.Assert ("ˆ—’†(1/20) |-------------------" = Application.StatusBar)
    
    Call showProgress(2, 20)
    Debug.Assert ("ˆ—’†(2/20) ||------------------" = Application.StatusBar)
    
    Call showProgress(2, 30)
    Debug.Assert ("ˆ—’†(2/30) |-------------------" = Application.StatusBar)
    Call showProgress(3, 30)
    Debug.Assert ("ˆ—’†(3/30) ||------------------" = Application.StatusBar)
    Call showProgress(4, 30)
    Debug.Assert ("ˆ—’†(4/30) |||-----------------" = Application.StatusBar)
    
    Application.StatusBar = False
End Function

Function testToAsin()
    Debug.Assert (asin("") = "")
    Debug.Assert (asin("a123-45-6789") = "")
    Debug.Assert (asin("4-86011-202-4") = "4860112024")
    Debug.Assert (asin("978-4-86011-202-8") = "4860112024")
End Function

Function testSignedUrlFor()
    Debug.Assert ( _
        signedUrlFor("4860112024", "00000000000000000000", "1234567890", "dummy", "2011-12-01T12:30:03+0900") _
        = "http://ecs.amazonaws.jp/onca/xml?AWSAccessKeyId=00000000000000000000&AssociateTag=dummy&ItemId=4860112024&Operation=ItemLookup&ResponseGroup=ItemAttributes&Service=AWSECommerceService&Timestamp=2011-12-01T12%3A30%3A03%2B0900&Version=2011-08-01&Signature=GzsoqdSW%2FxEOioChJC%2FHxbmS18Khp%2FLId0pfmO%2FYfo8%3D")
End Function

Function testBgColor()
    Call bgColor(ActiveSheet.Range("A1"), xlThemeColorAccent2)
    With ActiveSheet.Range("A1").Interior
        Debug.Assert (.Pattern = xlSolid)
        Debug.Assert (.PatternColorIndex = xlAutomatic)
        Debug.Assert (.ThemeColor = xlThemeColorAccent2)
'        Debug.Assert (.TintAndShade = CDbl(0.799951170384838)) ' ˆê’v‚³‚¹‚é•û–@‚ª‚í‚©‚ç‚È‚¢
        Debug.Assert (.PatternTintAndShade = 0)
    End With
    
    Call bgColor(ActiveSheet.Range("A1"), Null)
    With ActiveSheet.Range("A1").Interior
        Debug.Assert (.Pattern = xlNone)
        Debug.Assert (.TintAndShade = 0)
        Debug.Assert (.PatternTintAndShade = 0)
    End With
End Function

Sub aaa()
    application.
End Sub

