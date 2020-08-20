Attribute VB_Name = "Monitoring__Schema"
Sub Monitoring_schema()
    
    Dim emptyLine As Integer
    emptyLine = firstRowMonitoring
    Dim nowDate As Date
    nowDate = Date
            
    Dim listOrder As New Scripting.Dictionary
        
    'Check de chaque ligne de commande respectant les conditions du CLEAN EXTRACT
    For Each ligneDeCommande In commandesDuJour.Keys()
        
        SoldTo = sheetExtract.Cells(ligneDeCommande, columnSoldTo_SAP).Value
        Order = sheetExtract.Cells(ligneDeCommande, columnOrder_SAP).Value
        deliveryDate = sheetExtract.Cells(ligneDeCommande, columnRequestedDeliveryDate_SAP).Value
        
        If Not listOrder.Exists(Order) Then
            listOrder.Add Order, deliveryDate
        End If
        
    Next ligneDeCommande
    
    For Each Order In listOrder
        If listOrder(Order) < DateAdd("d", 7, nowDate) Then
            Copy_OrderLine_from_SAP sheetSchema, emptyLine, Order, False
            emptyLine = emptyLine + 1
        End If
    Next Order
    
End Sub
Sub prepaMailSchema()
    
    If Not functionVariables = "activated" Then Variables
    If Not functionEstablish_listSchema = "activated" Then Establish_listSchema
    
    For Each Client In listSchema.Keys
        Call mail(Get_Contact_Of(Client), "DANONE - Rappel Schéma de livraison", Mail_Schema_Of(Client))
    Next Client
    
End Sub
