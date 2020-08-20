Attribute VB_Name = "Data_Functions"
Public Function Get_SoldTo_Of(ByVal Client As String) As String
    'Recherche de la ligne correspondante dans la BDDClients
    Set FoundCell = BDDClients.Range("A:A").Find(What:=Client)
    
    If Not FoundCell Is Nothing Then
        clientLine = FoundCell.Row
        Get_SoldTo_Of = CStr(BDDClients.Cells(clientLine, 2).Value)
    Else
        Get_SoldTo_Of = "Not Found"
    End If
    
End Function
Public Function Get_ShipTo_Of(ByVal Client As String) As String
    'Recherche de la ligne correspondante dans la BDDClients
    Set FoundCell = BDDClients.Range("A:A").Find(What:=Client)
    
    If Not FoundCell Is Nothing Then
        clientLine = FoundCell.Row
        Get_ShipTo_Of = CStr(BDDClients.Cells(clientLine, 2).Value)
    Else
        Get_ShipTo_Of = "Not Found"
    End If
End Function
Public Function RemoveWhiteSpace(ByVal target As String) As String
    With New RegExp
        .Pattern = "\s"
        .MultiLine = True
        .Global = True
        RemoveWhiteSpace = .Replace(target, vbNullString)
    End With
End Function
Public Function RemoveSlash(ByVal target As String) As String
    With New RegExp
        .Pattern = "/"
        .MultiLine = True
        .Global = True
        RemoveSlash = .Replace(target, vbNullString)
    End With
End Function
