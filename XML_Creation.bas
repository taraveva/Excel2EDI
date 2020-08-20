Attribute VB_Name = "XML_Creation"
Sub CHD()
Attribute CHD.VB_ProcData.VB_Invoke_Func = " \n14"
        
    'Capture du nom du fichier template au cas où il change dans le temps ...
    NOMTEMPLATE = ActiveWorkbook().Name
    'Quantité totale pour le check de fin d'intégration
    TOTALQTE = Cells(1, 7).Value
    
    
        'Ferme le fichier et l'enrregistre dans l'archive - Format "Date+Heure+Nom du client"
        Sheets("CHD").Select
        Calculate
        H = Cells(1, 17)
        M = Cells(2, 17)
        S = Cells(3, 17)
        J = Cells(4, 17)
        Mth = Cells(5, 17)
        A = Cells(6, 17)
        
        NewName = J & "_" & Mth & "_" & A & " a " & H & "-" & M & "-" & S & " " & Name & ".xlsx"
        Windows(NOM).Activate
        ActiveWorkbook.SaveAs Filename:= _
                "\\Nead.danet/fr_shares/DPFF_BurSupply/CUSTOMER_SERVICE/FRONT OFFICE/CLIENTS/DANONE PRO/$ Fichiers opérationnels quotidiens\COMMANDES EXCEL TO EDI/Archives/" & NewName _
                , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
          
        ActiveWindow.Close
    Next File
    
    Sheets("Template").Select
    Rows(Y).Select
    Selection.Delete Shift:=xlUp
    Cells(1, 1).Select
    
    'Corrige la date
    Sheets("Template").Select
    Range("P2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(DAY(RC[-3])<10," & Chr(10) & "IF(MONTH(RC[-3])<10,1*(YEAR(RC[-3])&""0""&MONTH(RC[-3])&""0""&DAY(RC[-3]))," & Chr(10) & "1*(YEAR(RC[-3])&MONTH(RC[-3])&""0""&DAY(RC[-3])))," & Chr(10) & "IF(MONTH(RC[-3])<10,1*(YEAR(RC[-3])&""0""&MONTH(RC[-3])&DAY(RC[-3])),1*(YEAR(RC[-3])&MONTH(RC[-3])&DAY(RC[-3]))))"

    
    'Extract format XML sur le réseau - Format "FR1_DUMBCP_Date_Heure"

    ActiveWorkbook.XmlMaps("TO_COR_BCP_ORDER_Mappage").Export URL:= _
            "\\wfrstob067/dpff_bcp$/" & "FR1_DUMBCP_" & A & Mth & J & "_" & M & S & ".xml"
            

End Sub
Sub Upload_Orders()
    Dim CommandeClient As Worksheet
    Dim listProduct As Scripting.Dictionary
    
    If Not function_Variables = "Activated" Then Variables
    
    ChDir "C:\Commandes Excel\"
    FichierCommande = Dir("")
    FilePath = "C:\Commandes Excel\" & FichierCommande
    Line = 4
    nbOrders = 0
    'Récapitulatif des commandes dans l'onglet principal
    While FichierCommande <> ""
        Workbooks.Open Filename:=FilePath, UpdateLinks:=True, ReadOnly:=True
        Set CommandeClient = Sheets("Fiche Commande Danone")
        CommandeClient.Unprotect ""
        
        Client = CStr(CommandeClient.Cells(2, 1).Value)
        PO = CStr(CommandeClient.Cells(2, 2).Value)
        If PO = "" Then Commande.Cells(Line, 2).Value = Left(CStr(Commande.Cells(Line, 1).Value), 5) & Day(CDate(Commande.Cells(Line, 3).Value)) & Month(CDate(Commande.Cells(Line, 3).Value))
        
        delivDate = CStr(CommandeClient.Cells(2, 3).Value)
        If delivDate = "" Then delivDate = jourJ1

        SoldTo = Get_SoldTo_Of(Client)
        ShipTo = Get_ShipTo_Of(Client)
        OrderType = "ZSO"
        If Not CommandeClient.Cells(2, 10).Value = "" Then OrderType = CStr(CommandeClient.Cells(2, 10).Value)
        OrderReason = ""
        If Not CommandeClient.Cells(2, 12).Value = "" Then OrderReason = CStr(CommandeClient.Cells(2, 12).Value)
        DelivBlock = ""
        Channel = "01"
        If Not CommandeClient.Cells(2, 11).Value = "" Then Channel = CStr(CommandeClient.Cells(2, 11).Value)
        Note = ""
        If Not CommandeClient.Cells(2, 14).Value = "" Then Note = CStr(CommandeClient.Cells(2, 14).Value)
        Plant = ""
        If Not CommandeClient.Cells(2, 13).Value = "" Then Plant = CStr(CommandeClient.Cells(2, 13).Value)
        OtherPartner = ""
        PartnerFunction = ""
        TypeOfOtherPartner = ""
 
        If SoldTo = 150052659 Or SoldTo = 150035933 Then 'Si il s'agit d'une commande cactus ou mistral alors rajout d'un 3rd party Y6
            If SoldTo = 150052659 Then 'Commande Cactus
                OrderType = "ZODO"
                Channel = "00"
                OtherPartner = "150051245"
                PartnerFunction = "Y6"
                TypeOfOtherPartner = "B"
            End If
            If SoldTo = 150035933 Then 'Commande Mistral
                OrderType = "ZODO"
                Channel = "00"
                OtherPartner = "150051245"
                PartnerFunction = "Y6"
                TypeOfOtherPartner = "B"
            End If
        End If
        
        firstLine = 2
        lastLine = CommandeClient.Cells(Rows.Count, 4).End(xlUp).Row
        totalQty = 0
        Set listProduct = New Scripting.Dictionary
        
        For i = firstLine To lastLine
            ProductCode = CStr(CommandeClient.Cells(i, columnProductCode).Value)
            Quantity = CInt(CommandeClient.Cells(i, columnQty).Value)
            totalQty = totalQty + Quantity
            If Quantity > 0 Then listProduct.Add ProductCode, CStr(Quantity)
        Next i
        
        Commande.Cells(Line, columnClient).Value = Client
        Commande.Cells(Line, columnOrderQty).Value = CStr(totalQty)
        Commande.Cells(Line, columnDelivDate).Value = delivDate
        Commande.Cells(Line, columnPO).Value = PO
        Commande.Cells(Line, columnSoldTo).Value = SoldTo
        Commande.Cells(Line, columnShipTo).Value = ShipTo
        Commande.Cells(Line, columnOrderType).Value = OrderType
        Commande.Cells(Line, columnDelivBlock).Value = DelivBlock
        Commande.Cells(Line, columnChannel).Value = Channel
        Commande.Cells(Line, columnNote).Value = Note
        Commande.Cells(Line, columnPartnerFunction).Value = PartnerFunction
        Commande.Cells(Line, columnPartner).Value = OtherPartner
        Commande.Cells(Line, columnPlant).Value = Plant
        Commande.Cells(Line, columnOrderReason).Value = OrderReason
        
        Line = Line + 2
        nbOrders = nbOrders + 1
        
        listOrders.Add nbOrders, listProduct
        Workbooks(FichierCommande).Close SaveChanges:=False
        Kill FichierCommande
        FichierCommande = Dir("")
        FilePath = "C:\Commandes Excel\" & FichierCommande
 
    Wend
    function_Upload_Orders = "Activated"
    End Sub
    Sub XML_File_Creation()
    
    Dim xmlFile  As DOMDocument
    Dim objRootElem As IXMLDOMElement
    Dim objMemberElem As IXMLDOMElement
    Dim objMemberName As IXMLDOMElement
    
    Dim champ As Variant
    Dim Product As Variant
    
    Set xmlFile = New DOMDocument
   
   ' Creates root element
    Set objRootElem = xmlFile.createElement("TO_COR_BCP_ORDER")
    xmlFile.appendChild objRootElem
    
    firstOrder = 4
    lastOrder = Commande.Cells(Rows.Count, columnClient).End(xlUp).Row
    Order = 0
    For Line = firstOrder To lastOrder Step 2
        Order = Order + 1
        For Each Product In listOrders(Order)
            
            dateSAP = Right(Commande.Cells(Line, columnDelivDate).Value, 4) & Mid(Commande.Cells(Line, columnDelivDate).Value, 4, 2) & Left(Commande.Cells(Line, columnDelivDate).Value, 2)
            deliveryDate = CStr(dateSAP)
            PO = CStr(Commande.Cells(Line, columnPO).Value)
            SoldTo = CStr(Commande.Cells(Line, columnSoldTo).Value)
            ShipTo = CStr(Commande.Cells(Line, columnShipTo).Value)
            OrderType = CStr(Commande.Cells(Line, columnOrderType).Value)
            DeliveryBlock = CStr(Commande.Cells(Line, columnDelivBlock).Value)
            Channel = CStr(Commande.Cells(Line, columnChannel).Value)
            DetailOfText = CStr(Commande.Cells(Line, columnNote).Value)
            PartnerFunction = CStr(Commande.Cells(Line, columnPartnerFunction).Value)
            OtherPartner = CStr(Commande.Cells(Line, columnPartner).Value)
            If OtherPartner = "" Then TypeOfOtherPartnerCode = ""
            If Not OtherPartner = "" Then TypeOfOtherPartnerCode = "B"
            If DetailOfText = "" Then TextType = ""
            If Not DetailOfText = "" Then TextType = "Z001"
            OrderReason = CStr(Commande.Cells(Line, columnOrderReason).Value)
            Plant = CStr(Commande.Cells(Line, columnPlant).Value)
            Material = CStr(Product)
            Quantity = CStr(listOrders(Order)(Product))
            
            Set objMemberElem = xmlFile.createElement("BCPOrder")
            objRootElem.appendChild objMemberElem
            
            For Each champ In champs_BCP.Keys
                
                Set objMemberName = xmlFile.createElement(CStr(champ))
                objMemberElem.appendChild objMemberName
                    
                Select Case champ

                    Case "RequestedDeliveryDate"
                        objMemberName.Text = deliveryDate
                    Case "PONumber"
                        objMemberName.Text = PO
                    Case "SoldToCode"
                        objMemberName.Text = SoldTo
                    Case "ShipToCode"
                        objMemberName.Text = ShipTo
                    Case "OrderType"
                        objMemberName.Text = OrderType
                    Case "DeliveryBlock"
                        objMemberName.Text = DeliveryBlock
                    Case "Channel"
                        objMemberName.Text = Channel
                    Case "TextType"
                        objMemberName.Text = TextType
                    Case "DetailOfText"
                        objMemberName.Text = DetailOfText
                    Case "PartnerFunctionOfOtherPartner"
                        objMemberName.Text = PartnerFunction
                    Case "TypeOfOtherPartnerCode"
                        objMemberName.Text = TypeOfOtherPartnerCode
                    Case "OtherPartnerCode"
                        objMemberName.Text = OtherPartner
                    Case "OrderReason"
                        objMemberName.Text = OrderReason
                    Case "Material"
                        objMemberName.Text = Material
                    Case "Quantity"
                        objMemberName.Text = Quantity
                    Case "Plant"
                        objMemberName.Text = Plant
                    Case Else
                        objMemberName.Text = CStr(champs_BCP(champ))
                End Select
            Next champ
        Next Product
    Next Line
    
   ' Saves XML data to disk.
   POclean = RemoveWhiteSpace(PO)
   POclean = RemoveSlash(POclean)
   xmlFile.Save ("\\wfrstob067/dpff_bcp$/" & "FR1_DUMBCP_" & POclean & "_" & deliveryDate & ".xml")
End Sub
