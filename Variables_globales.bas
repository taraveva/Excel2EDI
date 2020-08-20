Attribute VB_Name = "Variables_globales"
'Déclaration de l'onglet principal du fichier
Public Commande As Worksheet
Public BDDClients As Worksheet

'Déclaration des champs du message BCP
Public champs_BCP As Scripting.Dictionary
Public listOrders As Scripting.Dictionary

'Déclaration de la date du jour et du lendemain
Public jourJ As Date
Public jourJ1 As Date

'Déclaration des colonnesde l'onglet principal
Public columnClient As Integer
Public columnOrderQty As Integer
Public columnDelivDate As Integer
Public columnPO As Integer
Public columnSoldTo As Integer
Public columnShipTo As Integer
Public columnOrderType As Integer
Public columnDelivBlock As Integer
Public columnChannel As Integer
Public columnNote As Integer
Public columnPartnerFunction As Integer
Public columnPartner As Integer
Public columnOrderReason As Integer
Public columnPlant As Integer

'Déclaration des colonnes de l'onglet commande client
Public columnProductCode As Integer
Public columnQty As Integer

'Déclaration du témoin d'initialisation des variables
Public function_Variables As String
Public function_Upload_Orders As String

Sub Variables()
    
    'Déclaration de l'onglet principal du fichier
    Set Commande = Sheets("Commandes")
    Set BDDClients = Sheets("BDDClients")
    
    'Déclaration des champs du message BCP
    Set champs_BCP = New Scripting.Dictionary
    champs_BCP.Add "OrderCBU", "FR7"
    champs_BCP.Add "OrderLogicalMessage", "ZFR7ORDERS_07"
    champs_BCP.Add "OrderMessageFunction", "BCP"
    champs_BCP.Add "SoldToCode", "" 'variable
    champs_BCP.Add "TypeOfSoldToCode", "B"
    champs_BCP.Add "ShipToCode", "" 'variable
    champs_BCP.Add "TypeofShipToCode", "B"
    champs_BCP.Add "PONumber", "" 'variable
    champs_BCP.Add "RequestedDeliveryDate", "" 'variable
    champs_BCP.Add "OrderType", "" 'variable
    champs_BCP.Add "SalesOrg", "0024"
    champs_BCP.Add "Channel", "" 'variable
    champs_BCP.Add "Division", "00"
    champs_BCP.Add "Material", "" 'variable
    champs_BCP.Add "TypeOfMaterialCode", "B"
    champs_BCP.Add "Quantity", "" 'variable
    champs_BCP.Add "UnitOfQuantity", "CT"
    champs_BCP.Add "ConditionCurrency", "EUR"
    champs_BCP.Add "OtherPartnerCode", "" 'variable
    champs_BCP.Add "PartnerFunctionOfOtherPartner", "" 'variable
    champs_BCP.Add "TypeOfOtherPartnerCode", "" 'variable
    champs_BCP.Add "DetailOfText", "" 'variable
    champs_BCP.Add "TextType", "" 'variable
    champs_BCP.Add "DeliveryBlock", "" 'variable
    champs_BCP.Add "PurchaseOrderType", "DFUE"
    champs_BCP.Add "OrderReason", "" 'variable
    champs_BCP.Add "Plant", "" 'variable
    
    Set listOrders = New Scripting.Dictionary
    
    jourJ = Date
    jourJ1 = DateAdd("d", 1, jourJ)
    
    columnClient = 3
    columnOrderQty = 5
    columnDelivDate = 7
    columnPO = 9
    columnSoldTo = 11
    columnShipTo = 13
    columnOrderType = 15
    columnDelivBlock = 17
    columnChannel = 19
    columnNote = 21
    columnPartnerFunction = 23
    columnPartner = 25
    columnOrderReason = 27
    columnPlant = 29
    columnProductCode = 4
    columnQty = 6
    function_Upload_Orders = ""
    function_Variables = "Activated"
End Sub
