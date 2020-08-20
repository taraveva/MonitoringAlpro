Attribute VB_Name = "Monitoring__Validation"
Sub Monitoring_validation()
     'Première ligne du tableau
    Dim emptyLine As Integer
    emptyLine = firstRowMonitoring
    Dim listOrders As Scripting.Dictionary
    Set listOrders = New Scripting.Dictionary
    
    If Not functionVariables = "activated" Then Variables
    If Not functionEstablish_listSchema = "activated" Then Establish_listSchema
    'If Not functionEstablish_listRuptures = "activated" Then Establish_listRuptures
    If Not functionEstablish_listFrequence = "activated" Then Establish_listFrequence
    If Not functionEstablish_listFranco = "activated" Then Establish_listFranco
    If Not functionEstablish_listCouche = "activated" Then Establish_listCouche
                
    'Check de chaque ligne de commande respectant date de création = date du jour
    For Each ligneDeCommande In commandesDuJour.Keys()
    
        Order = CLng(sheetExtract.Cells(ligneDeCommande, columnOrder_SAP).Value)
        SoldTo = CLng(sheetExtract.Cells(ligneDeCommande, columnSoldTo_SAP).Value)
        
        If Not listFranco.Exists(SoldTo) And Not listSchema.Exists(SoldTo) And Not listFrequence.Exists(SoldTo) And Not listCouche.Exists(SoldTo) Then
            If Not listOrders.Exists(Order) Then
                listOrders.Add Order, SoldTo
                Copy_OrderLine_from_SAP sheetValidation, emptyLine, Order, True
                emptyLine = emptyLine + 1
            End If
        End If
    Next ligneDeCommande
    
End Sub
Sub prepaMailValidation()
    
    If Not functionVariables = "activated" Then Variables
    If Not functionEstablish_listValidation = "activated" Then Establish_listValidation
    
    For Each Client In listValidation
        Call mail(Get_Contact_Of(Client), "DANONE - Prise de rendez-vous livraison", Mail_Validation_Of(Client))
    Next Client
    
End Sub

