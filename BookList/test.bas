Attribute VB_Name = "test"
Option Explicit

Function testShowProgress()
    Call showProgress(1, 20)
    Debug.Assert ("èàóùíÜ(1/20) |-------------------" = Application.StatusBar)
    
    Call showProgress(2, 20)
    Debug.Assert ("èàóùíÜ(2/20) ||------------------" = Application.StatusBar)
    
    Call showProgress(2, 30)
    Debug.Assert ("èàóùíÜ(2/30) |-------------------" = Application.StatusBar)
    Call showProgress(3, 30)
    Debug.Assert ("èàóùíÜ(3/30) ||------------------" = Application.StatusBar)
    Call showProgress(4, 30)
    Debug.Assert ("èàóùíÜ(4/30) |||-----------------" = Application.StatusBar)
    
    Application.StatusBar = False
End Function

Function testToAsin()
    Debug.Assert (toAsin("") = "")
    Debug.Assert (toAsin("a123-45-6789") = "")
    Debug.Assert (toAsin("4-86011-202-4") = "4860112024")
    Debug.Assert (toAsin("978-4-86011-202-8") = "4860112024")
End Function

Function testSignedUrlFor()
    Debug.Assert ( _
        signedUrlFor(asin:="4860112024", accessKey:="00000000000000000000", secretKey:="1234567890", associateTag:="dummy", timestamp:="2012-08-18T01:45:03+0900") _
        = "http://ecs.amazonaws.jp/onca/xml?AWSAccessKeyId=00000000000000000000&AssociateTag=dummy&ItemId=4860112024&Operation=ItemLookup&ResponseGroup=Large&Service=AWSECommerceService&Timestamp=2012-08-18T01%3A45%3A03%2B0900&Version=2011-08-01&Signature=Pdv%2FXkqw0rxGd2mvEiitZZ6zyq2g6Ezsxex1f0GXV9c%3D")
    Debug.Assert ( _
        signedUrlFor(title:="Ç†Ç†Ç†", accessKey:="00000000000000000000", secretKey:="1234567890", associateTag:="dummy", timestamp:="2012-08-18T01:45:03+0900") _
        = "http://ecs.amazonaws.jp/onca/xml?AWSAccessKeyId=00000000000000000000&AssociateTag=dummy&Operation=ItemSearch&ResponseGroup=Large&SearchIndex=Books&Service=AWSECommerceService&Timestamp=2012-08-18T01%3A45%3A03%2B0900&Title=%E3%81%82%E3%81%82%E3%81%82&Version=2011-08-01&Signature=zF8SI0YiOeL9qMLU0EPhhf6nMq4JKv%2FYnFWPjSxSe%2B4%3D")
    Debug.Assert ( _
        signedUrlFor(title:="Ç†Ç†Ç†", author:="Ç¢Ç¢Ç¢", publisher:="Ç§Ç§Ç§", accessKey:="00000000000000000000", secretKey:="1234567890", associateTag:="dummy", timestamp:="2012-08-18T01:45:03+0900") _
        = "http://ecs.amazonaws.jp/onca/xml?AWSAccessKeyId=00000000000000000000&AssociateTag=dummy&Author=%E3%81%84%E3%81%84%E3%81%84&Operation=ItemSearch&Publisher=%E3%81%86%E3%81%86%E3%81%86&ResponseGroup=Large&SearchIndex=Books&Service=AWSECommerceService&Timestamp=2012-08-18T01%3A45%3A03%2B0900&Title=%E3%81%82%E3%81%82%E3%81%82&Version=2011-08-01&Signature=7jnrKxKrc8L39Wt%2BMhpCYrPRIWLuC7Ze0dvxdTPl%2Fgg%3D")
End Function

Function testBgColor()
    Call bgColor(ActiveSheet.Range("A1"), xlThemeColorAccent2)
    With ActiveSheet.Range("A1").Interior
        Debug.Assert (.Pattern = xlSolid)
        Debug.Assert (.PatternColorIndex = xlAutomatic)
        Debug.Assert (.ThemeColor = xlThemeColorAccent2)
'        Debug.Assert (.TintAndShade = CDbl(0.799951170384838)) ' àÍívÇ≥ÇπÇÈï˚ñ@Ç™ÇÌÇ©ÇÁÇ»Ç¢
        Debug.Assert (.PatternTintAndShade = 0)
    End With
    
    Call bgColor(ActiveSheet.Range("A1"), Null)
    With ActiveSheet.Range("A1").Interior
        Debug.Assert (.Pattern = xlNone)
        Debug.Assert (.TintAndShade = 0)
        Debug.Assert (.PatternTintAndShade = 0)
    End With
End Function

