Attribute VB_Name = "Module1"
Sub getHandles()

Dim request As Object
Set request = CreateObject("MSXML2.XMLHTTP")

For b = 6 To Sheets(1).Cells(Rows.count, 8).End(xlUp).Row
    ublnum = Sheets(1).Cells(b, 2)
    shelpos = 0
    a = 0
    Do While shelpos = 0
        strlen = Len(Sheets(1).Cells(b, 8)) - a
        If InStr(strlen, Sheets(1).Cells(b, 8), ",") = 0 Then
            a = a + 1
        Else
            shelpos = Len(Sheets(1).Cells(b, 8)) - InStr(strlen, Sheets(1).Cells(b, 8), ",") - 1
        End If
        DoEvents
    Loop
    shelfmark = Right(Sheets(1).Cells(b, 8), shelpos)
    'Indien in Islandora een shelfmark met een : werkt ipv een -
    'BV "D H 797-2" in het sheet maar "D H 797: 2" in Islandora
    'haal dan het apostrof weg bij de onderstaande regel
    
    'shelfmark = Replace(shelfmark, "-", ": ")
    With request
        .Open "GET", "https://islandora7-solr.universiteitleiden.nl/solr/collection1/select?q=mods_relatedItem_otherFormat_identifier_ms%3A%22" & shelfmark & "%22&fl=mods_identifier_hdl_ms&wt=csv&indent=true", False, "", ""
        .send
        handlevar = Mid(.responseText, InStr(.responseText, "http"), Len(.responseText) - InStr(.responseText, "http"))
        Sheets(1).Cells(b, 7) = handlevar
        Sheets(1).Cells(b, 10) = handlevar
        pidvar = Right(Sheets(1).Cells(b, 7), Len(Sheets(1).Cells(b, 7)) - InStr(Sheets(1).Cells(b, 7), "m:") - 1)
        Sheets(1).Cells(b, 20) = "https://digitalcollections.universiteitleiden.nl/view/item/" & pidvar & "/datastream/TN/view"
    End With
    DoEvents
Next b

End Sub
