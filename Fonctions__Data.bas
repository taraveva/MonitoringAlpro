Attribute VB_Name = "Fonctions__Data"
Public Function Get_Contact_Of(ByVal Client As Long) As String 'passer le client en argument via son SoldTo
     
     'Recherche de la ligne correspondante dans la BDDClient
     Set FoundCell = BDDClients.Range("A:A").Find(What:=Client)
     clientLine = FoundCell.row
     
     'Récupération de son contact
     Get_Contact_Of = CStr(BDDClients.Cells(clientLine, columnContactAppro).Value)
     
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
Public Function Get_RAN_Of(ByVal Produit As Long) As String
    
    Set FoundRupture = sheetDMS.Range("B:B").Find(What:=Produit)
    productRupture = FoundRupture.row
    Get_RAN_Of = sheetDMS.Cells(productRupture, columnRAN).Value
    
End Function
