Attribute VB_Name = "Fonctions_globales"
Sub Validation()

    If Not function_Variables = "Activated" Then Variables
        
    If Not function_Upload_Orders = "Activated" Then Upload_Orders
    
    XML_File_Creation
    
    
End Sub

Sub Charger()

    If Not function_Variables = "Activated" Then Variables
        
    Upload_Orders
      
End Sub

Sub ViderCommandes()

    'Nettoyage du tableau de RECAP du check
    Range("C4:AC22").Select
    Selection.ClearContents
    
    'vider dossier sur C:/
    Const dossier As String = "C:\Commandes Excel\"
    Dim Fichier As String
    Fichier = Dir(dossier)
    Do While Fichier <> ""
        Kill dossier & Fichier
        Fichier = Dir
    Loop
    function_Variables = "Deactivated"
    End
End Sub



