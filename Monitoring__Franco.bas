Attribute VB_Name = "Monitoring__Franco"
Sub Monitoring_franco()
    
    Dim emptyLine As Integer
    emptyLine = firstRowMonitoring
    Dim listSoldToDelivery As New Scripting.Dictionary
    Dim listDelivery As New Scripting.Dictionary
    Dim listDeliveryDetails As New Scripting.Dictionary
        
    'Check de chaque ligne de commande identifiées comme commandes créees le jour-même
    For Each ligneDeCommande In commandesDuJour.Keys()
        
        SoldTo = CLng(sheetExtract.Cells(ligneDeCommande, columnSoldTo_SAP).Value)
        Order = CLng(sheetExtract.Cells(ligneDeCommande, columnOrder_SAP).Value)
        orderedQty = CLng(sheetExtract.Cells(ligneDeCommande, columnOrderQty_SAP).Value)
        deliveryDate = CDate(sheetExtract.Cells(ligneDeCommande, columnRequestedDeliveryDate_SAP).Value)
             
        'Check si le client n'est pas exempté de cette CGV
        If Not listExceptionsClient.Exists(SoldTo) Then
            
            If Not listSoldToDelivery.Exists(SoldTo) Then
                Set listDeliveryDetails = New Scripting.Dictionary
                Set listDelivery = New Scripting.Dictionary
                listDeliveryDetails.Add Order, orderedQty
                listDelivery.Add deliveryDate, listDeliveryDetails
                listSoldToDelivery.Add SoldTo, listDelivery
            Else
                If Not listSoldToDelivery(SoldTo).Exists(deliveryDate) Then
                    Set listDeliveryDetails = New Scripting.Dictionary
                    listDeliveryDetails.Add Order, orderedQty
                    listSoldToDelivery(SoldTo).Add deliveryDate, listDeliveryDetails
                Else
                    If Not listSoldToDelivery(SoldTo)(deliveryDate).Exists(Order) Then
                        listSoldToDelivery(SoldTo)(deliveryDate).Add Order, orderedQty
                    Else
                        listSoldToDelivery(SoldTo)(deliveryDate)(Order) = listSoldToDelivery(SoldTo)(deliveryDate)(Order) + orderedQty
                    End If
                End If
            End If
        End If
            
    Next ligneDeCommande
    
    For Each SoldTo In listSoldToDelivery
        For Each deliveryDate In listSoldToDelivery(SoldTo)
            arrayOrders = listSoldToDelivery(SoldTo)(deliveryDate).Items
            orderedQty = 0
            For i = LBound(arrayOrders) To UBound(arrayOrders)
                orderedQty = orderedQty + arrayOrders(i)
            Next
            If orderedQty < 95 Then
                For Each Order In listSoldToDelivery(SoldTo)(deliveryDate)
                    Set FoundCell = sheetExtract.Range("A:A").Find(What:=Order)
                    If Not FoundCell Is Nothing Then
                        orderLine = FoundCell.row
                        For j = 1 To 15
                            Select Case j
                            Case 5 To 9
                            Case 13
                                 sheetFranco.Cells(emptyLine, j + 1).Value = listSoldToDelivery(SoldTo)(deliveryDate)(Order)
                            Case Else
                                sheetFranco.Cells(emptyLine, j + 1).Value = sheetExtract.Cells(orderLine, j).Value
                            End Select
                        Next j
                        emptyLine = emptyLine + 1
                    End If
                Next Order
            End If
        Next deliveryDate
    Next SoldTo
End Sub
Sub prepaMailFranco()
    
    If Not functionVariables = "activated" Then Variables
    If Not functionEstablish_listFranco = "activated" Then Establish_listFranco
    
    For Each Client In listFranco
        Call mail(Get_Contact_Of(Client), "DANONE - Rappel Franco commande", Mail_Franco_Of(Client))
    Next Client
    
End Sub
