Attribute VB_Name = "Variables_globales"
'Declaration des deux workbook utilisé: l'outils et l'extract SAP
Public pilotage As Workbook
Public exportSAP As Workbook

'----------Declaration de constantes------------------------------------------------------------------------
'Declaration des colonnes de l'extract SAP
Public Const columnOrder_SAP As Integer = 1
Public Const columnSoldTo_SAP As Integer = 3
Public Const columnPO_SAP As Integer = 4
Public Const columnMaterial_SAP As Integer = 6
Public Const columnMaterialDescription_SAP As Integer = 7
Public Const columnDelivBlock_SAP As Integer = 9
Public Const columnCreatedOn_SAP As Integer = 10
Public Const columnMaterialAvaibilityDate_SAP As Integer = 11
Public Const columnRequestedDeliveryDate_SAP As Integer = 12
Public Const columnOrderQty_SAP As Integer = 13

'Declaration des colonnes des tableaux de monitoring
Public Const columnOrder_Monitoring As Integer = 2
Public Const columnSoldToName_Monitoring As Integer = 3
Public Const columnSoldTo_Monitoring As Integer = 4
Public Const columnPO_Monitoring As Integer = 5
Public Const columnMaterial_Monitoring As Integer = 7
Public Const columnMaterialDescription_Monitoring As Integer = 8
Public Const columnDelivBlock_Monitoring As Integer = 10
Public Const columnCreatedOn_Monitoring As Integer = 11
Public Const columnMaterialAvaibilityDate_Monitoring As Integer = 12
Public Const columnRequestedDeliveryDate_Monitoring As Integer = 13
Public Const columnOrderQty_Monitoring As Integer = 14

'Declaration des colonnes de la BDDProduit
Public Const columnLibelle As Integer = 2
Public Const columnNbCaissesCouche As Integer = 3
Public Const columnNbCaissesPalette As Integer = 5
Public Const columnEAN As Integer = 6

'Declaration des colonnes de la BDDClient
Public Const columnContactAppro As Integer = 3
Public Const columnEntrepot As Integer = 2
Public Const columnHourStart As Integer = 5

'Déclaration des colonnes du DMS
Public Const columnRAN As Integer = 8

'Declaration de la première ligne des tableaux de monitoring
Public Const firstRowMonitoring As Integer = 9
'-----------------------------------------------------------------------------------------------------

'Déclaration de la dernière ligne de l'exract SAP
Public lastRowExportSAP As Integer

Public sheetPilotage As Worksheet
Public sheetSuiviQuotidien As Worksheet
Public sheetSuiviLongTerme As Worksheet
Public sheetDMS As Worksheet
Public sheetRuptures As Worksheet
Public sheetCouche As Worksheet
Public sheetFrequence As Worksheet
Public sheetFranco As Worksheet
Public sheetSchema As Worksheet
Public sheetValidation As Worksheet
Public BDDProduits As Worksheet
Public BDDClients As Worksheet
Public Archives As Worksheet
Public sheetExtract As Worksheet

Public listExceptionsClient As New Scripting.Dictionary 'Liste des clients Auchan et Leclerc dispensés du monitoring à la couche
Public listExceptionsProduit As New Scripting.Dictionary 'Liste des produits promo Monitoring display
Public sheetsMonitoring As New Scripting.Dictionary
Public commandesDuJour As New Scripting.Dictionary 'Liste des commandes dans l'extract SAP qui n'ont pas de client ou produit en erreur et du jour
Public commandesAllTime As New Scripting.Dictionary 'Liste des commandes dans l'extract SAP qui n'ont pas de client ou produit en erreur sans contrainte de date de création
Public listRuptures As New Scripting.Dictionary 'Liste des SoldTo nécessitant un mail du monitoring Ruptures
Public listCouche As New Scripting.Dictionary 'Liste des SoldTo nécessitant un mail du monitoring Couche
Public listFrequence As New Scripting.Dictionary 'Liste des SoldTo nécessitant un mail du monitoring Frequence
Public listFranco As New Scripting.Dictionary 'Liste des SoldTo nécessitant un mail du monitoring Franco
Public listSchema As New Scripting.Dictionary 'Liste des SoldTo nécessitant un mail du monitoring Schema
Public listValidation As New Scripting.Dictionary

Public functionVariables As String
Public functionEstablish_listCouche As String
Public functionEstablish_listSchema As String
Public functionEstablish_listFranco As String
Public functionEstablish_listFrequence As String
Public functionEstablish_listRuptures As String
Public functionEstablish_listValidation As String

Public endLineRuptures As Long
Public endLineCouche As Long
Public endLineFrequence As Long
Public endLineFranco As Long
Public endLineSchema As Long
Public addLineFranco As Boolean
Public addLineCouche As Boolean
Public addLineRuptures As Boolean
Public addLineFrequence As Boolean
Public multi As Boolean
Sub Variables()

    'Récupération du nom du fichier dans le dossier "controle commandes" crée à la racine de C:
    ChDir "C:\Controle Commandes"
    Filename = Dir("*.xls")
    Filename = "C:\Controle Commandes\" + Filename
    
    Dim App As New Excel.Application
    
    'Déclaration des différents workbooks
    Set exportSAP = App.Workbooks.Open(Filename:=Filename, UpdateLinks:=True, ReadOnly:=True)
    Set sheetExtract = exportSAP.Worksheets("Feuil1")
    Set sheetPilotage = Worksheets("Pilotage")
    'Set sheetSuiviQuotidien = Worksheets("Suivi quotidien")
    'Set sheetSuiviLongTerme = Worksheets("Suivi long terme KPI")
    Set sheetDMS = Worksheets("DMS")
    Set sheetRuptures = Worksheets("Monitoring ruptures")
    Set sheetCouche = Worksheets("Monitoring à la couche")
    Set sheetFrequence = Worksheets("Fréquence de livraison")
    Set sheetFranco = Worksheets("Franco")
    Set sheetSchema = Worksheets("Schéma")
    Set sheetValidation = Worksheets("Validation")
    Set BDDProduits = Worksheets("BDD Produits")
    Set BDDClients = Worksheets("BDD Clients")
    Set Archives = Worksheets("Archives")
    
    'Nombre de commandes extraites de SAP
    lastRowExportSAP = sheetExtract.Cells(Rows.Count, 2).End(xlUp).row
    addLineFrequence = False
    addLineFranco = False
    addLineRuptures = False
    addLineCouche = False
    addLineSchema = False
    multi = False
    
    'Déclaration des listes d'exceptions clients et produits
    soldTo_Exclus = Array(150060444, 150060418, 150060442, 150060419, 150060443, 150052663, 150024794, 150048687, 150061849, 150061850, 150024454, 150060895, 150060520, 150060519, 150060531, 150060504, 150060898, 150061254, 150060860, 150060861, 150061888, 150061889, 150061901)
    produits_Display = Array(144811, 144815, 158262)
    
    sheets_Monitoring = Array(sheetRuptures, sheetCouche, sheetFrequence, sheetFranco, sheetSchema, sheetValidation)
    
    'Remplissage des dictionnaires ExceptionsClient, Exceptionsproduits
    For i = 0 To UBound(soldTo_Exclus)
        listExceptionsClient.Add soldTo_Exclus(i), i
    Next i
    For i = 0 To UBound(produits_Display)
        listExceptionsProduit.Add produits_Display(i), i
    Next i
    For i = 0 To UBound(sheets_Monitoring)
        If i = 0 Then sheetsMonitoring.Add sheets_Monitoring(i), "ruptures"
        If i = 1 Then sheetsMonitoring.Add sheets_Monitoring(i), "couche"
        If i = 2 Then sheetsMonitoring.Add sheets_Monitoring(i), "frequence"
        If i = 3 Then sheetsMonitoring.Add sheets_Monitoring(i), "franco"
        If i = 4 Then sheetsMonitoring.Add sheets_Monitoring(i), "schema"
        If i = 5 Then sheetsMonitoring.Add sheets_Monitoring(i), "validation"
    Next i
    
    functionEstablish_listFranco = "not activated"
    functionEstablish_listSchema = "not activated"
    functionEstablish_listCouche = "not activated"
    functionEstablish_listFrequence = "not activated"
    functionEstablish_listRuptures = "not activated"
    functionEstablish_listValidation = "not activated"
    functionVariables = "activated"
End Sub
