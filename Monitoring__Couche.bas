Attribute VB_Name = "Monitoring__Couche"
Sub Monitoring_couche()
    
    'Premi�re ligne du tableau
    Dim emptyLine As Integer
    emptyLine = firstRowMonitoring
            
    'Check de chaque ligne de commande respectant date de cr�ation = date du jour
    For Each ligneDeCommande In commandesDuJour.Keys()
    
        Product = CLng(sheetExtract.Cells(ligneDeCommande, columnMaterial_SAP).Value)
        SoldTo = CLng(sheetExtract.Cells(ligneDeCommande, columnSoldTo_SAP).Value)
        orderedQty = CLng(sheetExtract.Cells(ligneDeCommande, columnOrderQty_SAP).Value)
        
        'V�rification le client n'est pas Auchan ou Leclerc
        If Not listExceptionsClient.Exists(SoldTo) Then
            
            'Cherche les caract�ristiques du produit concern� dans la BDDproduit
            nbCaissesCouche = Get_CoucheCriteria(Product)
            nbCaissesPalette = Get_PaletteCriteria(Product)
            
            'V�rification si commande respecte la couche ou la palette
            If (orderedQty Mod nbCaissesCouche) <> 0 Then
                If (orderedQty Mod nbCaissesPalette) <> 0 Then
                    'Si �a ne respecte pas les conditions alors copie de la ligne de commande concern�e dans le tableau de monitoring
                    For j = 1 To 15
                        sheetCouche.Cells(emptyLine, j + 1).Value = sheetExtract.Cells(ligneDeCommande, j).Value
                    Next j
                    emptyLine = emptyLine + 1
                End If
            End If
        End If
    
    Next ligneDeCommande
       
End Sub

Sub prepaMailCouche()

    If Not functionVariables = "activated" Then Variables
    If Not functionEstablish_listCouche = "activated" Then Establish_listCouche

    For Each Client In listCouche
        Call mail(Get_Contact_Of(Client), "DANONE - Rappel commande � la couche", Mail_Couche_Of(Client))
    Next Client
    
End Sub
