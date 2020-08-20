Attribute VB_Name = "Monitoring__Ruptures"
Sub Monitoring_ruptures()

    Dim emptyLine As Integer
    emptyLine = firstRowMonitoring
            
    'Check de chaque ligne de commande
    For Each ligneDeCommande In commandesDuJour.Keys()
    
        Product = sheetExtract.Cells(ligneDeCommande, columnMaterial_SAP).Value
        SoldTo = sheetExtract.Cells(ligneDeCommande, columnSoldTo_SAP).Value
        deliveryDate = CDate(sheetExtract.Cells(ligneDeCommande, columnRequestedDeliveryDate_SAP).Value)
        
        
        Set FoundRupture = sheetDMS.Range("B:B").Find(What:=Product)
        
        If Not FoundRupture Is Nothing Then
            ruptureDate = CDate(sheetDMS.Cells(FoundRupture.row, columnRAN).Value)
            If deliveryDate < ruptureDate Then
                For j = 1 To 15
                    sheetRuptures.Cells(emptyLine, j + 1).Value = sheetExtract.Cells(ligneDeCommande, j).Value
                Next j
                emptyLine = emptyLine + 1
            End If
        End If

    Next ligneDeCommande
       
End Sub
Sub prepaMailRuptures()
         
    If Not functionVariables = "activated" Then Variables
    If Not functionEstablish_listRuptures = "activated" Then Establish_listRuptures
    
    For Each Client In listRuptures
       Call mail(Get_Contact_Of(Client), "DANONE - Alerte rupture", Mail_Ruptures_Of(Client))
    Next
    
End Sub
