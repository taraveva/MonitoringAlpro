Attribute VB_Name = "Monitoring__Frequence"
Sub Monitoring_frequence()
    
    Dim emptyLine As Integer
    emptyLine = firstRowMonitoring
    Dim arrayf(500, 500) As Variant
    'Déclaration des Dictionnaires
    Dim listDoublon As New Scripting.Dictionary
    Dim listOrderBook As New Scripting.Dictionary
    Dim listOrder As New Scripting.Dictionary
    Dim listOrderDetails As New Scripting.Dictionary
        
    For Each ligneDeCommande In commandesAllTime.Keys()
        
        deliveryDate = CDate(sheetExtract.Cells(ligneDeCommande, columnRequestedDeliveryDate_SAP).Value)
        SoldTo = sheetExtract.Cells(ligneDeCommande, columnSoldTo_SAP).Value
        Order = sheetExtract.Cells(ligneDeCommande, columnOrder_SAP).Value
        deliveryDateWeek = WorksheetFunction.WeekNum(deliveryDate, vbMonday)
             
        'Vérification que la semaine existe dans l'orderBook
        If Not listOrderBook.Exists(deliveryDateWeek) Then
            Set listOrder = New Scripting.Dictionary
            Set listOrderDetails = New Scripting.Dictionary
            listOrderDetails.Add Order, deliveryDate
            listOrder.Add SoldTo, listOrderDetails
            listOrderBook.Add deliveryDateWeek, listOrder
        Else
            'Vérification si une commande du même soldTo est présente dans la semaine
            If Not listOrderBook(deliveryDateWeek).Exists(SoldTo) Then
                'Si non alors on l'ajoute à l'orderBook
                Set listOrderDetails = New Scripting.Dictionary
                listOrderDetails.Add Order, deliveryDate
                listOrderBook(deliveryDateWeek).Add SoldTo, listOrderDetails
            Else
                'Si le SoldTo existe on vérifie si l'order existe
                If Not listOrderBook(deliveryDateWeek)(SoldTo).Exists(Order) Then
                    listOrderBook(deliveryDateWeek)(SoldTo).Add Order, deliveryDate
                End If
            End If
        End If
    Next ligneDeCommande
            
    For Each Week In listOrderBook
        For Each Client In listOrderBook(Week)
            
            If listOrderBook(Week)(Client).Count > 1 Then
                Set listDeliveryDates = New Scripting.Dictionary
                frequence = False
                
                For Each Order In listOrderBook(Week)(Client)
                    deliveryDate = listOrderBook(Week)(Client)(Order)
                    If Not listDeliveryDates.Exists(deliveryDate) Then
                        listDeliveryDates.Add deliveryDate, Order
                        If listDeliveryDates.Count > 1 Then
                            frequence = True
                        End If
                    End If
                Next Order
                If frequence = True Then
                    ordersArray = listOrderBook(Week)(Client).Keys
                    For num = LBound(ordersArray) To UBound(ordersArray)
                        Copy_OrderLine_from_SAP sheetFrequence, emptyLine, ordersArray(num), False
                        emptyLine = emptyLine + 1
                    Next num
                End If
            End If
        
        Next Client
    Next Week

End Sub
Sub prepaMailFrequence()
    If Not functionVariables = "activated" Then Variables
    If Not functionEstablish_listFrequence = "activated" Then Establish_listFrequence
    'Constitution d'un corps de mail par SoldTo
    For Each Client In listFrequence
        Call mail(Get_Contact_Of(Client), "DANONE - Problème Fréquence de livraison", Mail_Frequence_Of(Client))
    Next Client
End Sub

