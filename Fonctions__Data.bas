Attribute VB_Name = "Fonctions__Data"
Public Function Get_Contact_Of(ByVal SoldTo As Long) As String 'passer le SoldTo en argument via son SoldTo
     
     'Recherche de la ligne correspondante dans la BDDClients
     Set FoundCell = BDDClients.Range("A:A").Find(What:=SoldTo)
     clientLine = FoundCell.row
     
     'Récupération de son contact
     Get_Contact_Of = CStr(BDDClients.Cells(clientLine, columnContactAppro).Value)
     
End Function
Public Function Get_OrderedQty_Of(ByVal Order As Long, ByVal Product As Long) As String
    Dim i As Integer
    For i = 2 To lastRowExportSAP
        If sheetExtract.Cells(i, columnOrder_SAP).Value = Order And sheetExtract.Cells(i, columnMaterial_SAP).Value = Product Then
            Get_OrderedQty_Of = CStr(sheetExtract.Cells(i, columnOrderQty_SAP).Value)
            Exit Function
        End If
    Next i
End Function
Public Function Get_CoucheCriteria(ByVal Produit As Long) As Long
     
     'Recherche de la ligne correspondante dans la BDDProduit
     Set FoundCell = BDDProduits.Range("A:A").Find(What:=Produit)
     productLine = FoundCell.row
     
     'Récupération du critère de couche
     Get_CoucheCriteria = BDDProduits.Cells(productLine, columnNbCaissesCouche).Value
          
End Function
Public Function Get_PaletteCriteria(ByVal Produit As Long) As Long
     
     'Recherche de la ligne correspondante dans la BDDProduit
     Set FoundCell = BDDProduits.Range("A:A").Find(What:=Produit)
     productLine = FoundCell.row
     
     'Récupération du critère de couche
     Get_PaletteCriteria = BDDProduits.Cells(productLine, columnNbCaissesPalette).Value
          
End Function
Public Function Get_EAN_Of(ByVal Produit As Long) As String
     
     'Recherche de la ligne correspondante dans la BDDProduit
     Set FoundCell = BDDProduits.Range("A:A").Find(What:=Produit)
     productLine = FoundCell.row
     
     'Récupération du critère de couche
     Get_EAN_Of = BDDProduits.Cells(productLine, columnEAN).Value
          
End Function
Public Function Get_Libelle_Of(ByVal Produit As Long) As String
     
     'Recherche de la ligne correspondante dans la BDDProduit
     Set FoundCell = BDDProduits.Range("A:A").Find(What:=Produit)
     productLine = FoundCell.row
     
     'Récupération du critère de couche
     Get_Libelle_Of = BDDProduits.Cells(productLine, columnLibelle).Value
          
End Function
Public Function Get_PO_Of(ByVal Order As Long) As String
     
     'Recherche de la ligne correspondante dans la BDDProduit
     Set FoundCell = sheetExtract.Range("A:A").Find(What:=Order)
     orderLine = FoundCell.row
     
     'Récupération du critère de couche
     Get_PO_Of = sheetExtract.Cells(orderLine, columnPO_SAP).Value
          
End Function
Public Function Get_Entrepot_Of(ByVal SoldTo As Long) As String
     
     'Recherche de la ligne correspondante dans la BDDProduit
     Set FoundCell = BDDClients.Range("A:A").Find(What:=SoldTo)
     SoldToLine = FoundCell.row
     
     'Récupération du critère de couche
     Get_Entrepot_Of = BDDClients.Cells(SoldToLine, columnEntrepot).Value
          
End Function
Public Function Get_DeliveryDate_Of(ByVal Order As Long) As String
     
     'Recherche de la ligne correspondante dans la BDDProduit
     Set FoundCell = sheetExtract.Range("A:A").Find(What:=Order)
     orderLine = FoundCell.row
     
     'Récupération du critère de couche
     Get_DeliveryDate_Of = sheetExtract.Cells(orderLine, columnRequestedDeliveryDate_SAP).Value
          
End Function
Public Function Get_RAN_Of(ByVal Produit As Long) As String
    
    Set FoundRupture = sheetDMS.Range("B:B").Find(What:=Produit)
    productRupture = FoundRupture.row
    Get_RAN_Of = sheetDMS.Cells(productRupture, columnRAN).Value
    
End Function
Public Function Get_deliveryHour_Of(ByVal SoldTo As Long) As String
     
     'Recherche de la ligne correspondante dans la BDDProduit
     Set FoundCell = BDDClients.Range("A:A").Find(What:=SoldTo)
     SoldToLine = FoundCell.row
     
     'Récupération de l'heure de livraison
     Get_deliveryHour_Of = BDDClients.Cells(SoldToLine, columnHourStart).Value
          
End Function
Public Function Get_preparationDate_Of(ByVal Order As Long) As Date
    'Recherche de la ligne correspondante dans la BDDProduit
     Set FoundCell = sheetExtract.Range("A:A").Find(What:=Order)
     orderLine = FoundCell.row
     
     'Récupération du critère de couche
     Get_preparationDate_Of = sheetExtract.Cells(orderLine, columnMaterialAvaibilityDate_SAP).Value
End Function
