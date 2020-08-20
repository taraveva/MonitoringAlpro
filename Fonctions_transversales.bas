Attribute VB_Name = "Fonctions_transversales"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function Copy_OrderLine_from_SAP(destinationSheet As Worksheet, destinationLine As Integer, ByVal Order As Long, shortVersion As Boolean)
    
    Set FoundCell = sheetExtract.Range("A:A").Find(What:=Order)
    orderLine = FoundCell.row
    If Not shortVersion = True Then
        For j = 1 To 15
            destinationSheet.Cells(destinationLine, j + 1).Value = sheetExtract.Cells(orderLine, j).Value
        Next j
    Else
        For j = 1 To 4
            destinationSheet.Cells(destinationLine, j + 1).Value = sheetExtract.Cells(orderLine, j).Value
        Next j
        For j = 9 To 16
            destinationSheet.Cells(destinationLine, j + 1).Value = sheetExtract.Cells(orderLine, j).Value
        Next j
    End If
    
End Function
Public Function Reset_Error()
    Range("T6:AF8").ClearContents
End Function
Sub Reset_ActiveSheet()

    Dim tableRange As String
    startRow = firstRowMonitoring
    lastRow = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).row
    
    If Not startRow = lastRow Then
        Let tableRange = "B" & startRow & ":P" & lastRow
        Range(tableRange).Select
        Selection.Delete Shift:=xlUp
    Else
        Let tableRange = "B" & startRow & ":P" & startRow
        Range(tableRange).Select
        Selection.ClearContents
    End If
    
End Sub
Public Function Reset_MonitoringSheets()
    
    Dim tableRange As String
    
    For Each monitoring_Sheet In sheetsMonitoring
        
        monitoring_Sheet.Activate
        lastRowMonitoring = monitoring_Sheet.Cells(Rows.Count, 2).End(xlUp).row
        If Not firstRowMonitoring = lastRowMonitoring Then
            Let tableRange = "B" & firstRowMonitoring & ":P" & lastRowMonitoring
            Range(tableRange).Select
            Selection.Delete Shift:=xlUp
        Else
            Let tableRange = "B" & firstRowMonitoring & ":P" & firstRowMonitoring
            Range(tableRange).Select
            Selection.ClearContents
        End If
    
    Next monitoring_Sheet
    
End Function
Sub CheckExtract()
    
    Dim ErrorClient As Integer
    ErrorClient = 0
    Dim ErrorProduct As Integer
    ErrorProduct = 0
    Dim ErrorZtext As Integer
    ErrorZtext = 0
    Dim nowDate As Date
    nowDate = Date

    Dim listErrorClient As New Scripting.Dictionary
    Dim listErrorProduct As New Scripting.Dictionary
    Dim listZText As New Scripting.Dictionary
    Set commandesDuJour = New Scripting.Dictionary
    
    'Pour chaque ligne de l'extract SAP
    For i = 2 To lastRowExportSAP
        
        Order = sheetExtract.Cells(i, 1).Value
        SoldTo = CLng(sheetExtract.Cells(i, 3).Value)
        Product = sheetExtract.Cells(i, 6).Value
        creationDate = sheetExtract.Cells(i, 10).Value
        lineInError = False
        
        'Vérification si la ligne de commande n'est pas en Ztext
        If Not Product = "" Then
        
            'Recherche du client dans la BDDClients
            Set FoundClient = BDDClients.Range("A:A").Find(What:=SoldTo)
            'Recherche du produit dans la BDDProduits
            Set FoundProduct = BDDProduits.Range("A:A").Find(What:=Product)
            
            'Si le client n'a pas été trouvé
            If FoundClient Is Nothing Then
                'Ajout à la liste d'erreurs clients (si il n'y est pas déjà)
                If Not listErrorClient.Exists(SoldTo) Then
                    listErrorClient.Add SoldTo, i
                    sheetPilotage.Cells(6, ErrorClient + 20).Value = SoldTo
                    ErrorClient = ErrorClient + 1
                    lineInError = True
                End If
                lineInError = True
            End If
            
            'Si le produit n'a pas été trouvé
            If FoundProduct Is Nothing Then
                'Ajout à la liste d'erreurs produits (si il n'y est pas déjà)
                If Not listErrorProduct.Exists(Product) Then
                    listErrorProduct.Add Product, i
                    sheetPilotage.Cells(7, ErrorProduct + 20).Value = Product
                    ErrorProduct = ErrorProduct + 1
                    lineInError = True
                End If
                lineInError = True
            End If
        Else
            If Not listZText.Exists(Order) Then
                listZText.Add Order, i
                sheetPilotage.Cells(8, ErrorZtext + 20).Value = Order
                ErrorZtext = ErrorZtext + 1
                lineInError = True
            End If
            lineInError = True
        End If
        
        If Not lineInError Then
            If creationDate = nowDate Then
                commandesDuJour.Add i, 1
            End If
            commandesAllTime.Add i, 1
        End If
    Next i
    
End Sub
Sub Monitoring()

    Reset_MonitoringSheets
    Monitoring_ruptures
    Monitoring_couche
    Monitoring_frequence
    Monitoring_franco
    Monitoring_schema
    Monitoring_validation
    sheetPilotage.Activate
    
End Sub
Sub Envoi_Mail()
    
   If Not functionVariables = "activated" Then Variables

    prepaMail_Monitoring
    prepaMailValidation
    Suivi_Quotidien
    Reset_MonitoringSheets
    sheetPilotage.Activate
    
End Sub
Sub Clean_Extract()
    
    If Not functionVariables = "activated" Then Variables
    Reset_Error
    CheckExtract
    
End Sub
Sub Vider_Dossier()
    
    Const dossier_exportSAP As String = "C:\Controle Commandes\"
    Dim fichier As String
    fichier = Dir(dossier_exportSAP)
    Do While fichier <> ""
        Kill dossier_exportSAP & fichier
        fichier = Dir
    Loop
    
End Sub
Sub Suivi_Quotidien()
    
    If Not functionVariables = "activated" Then Variables
    
    Dim tableRange As String

    startRow = 9
    LastRowSchema = sheetSchema.Cells(Rows.Count, 2).End(xlUp).row
    LastRowFranco = sheetFranco.Cells(Rows.Count, 2).End(xlUp).row
    LastRowRuptures = sheetRuptures.Cells(Rows.Count, 2).End(xlUp).row
    LastRowCouche = sheetCouche.Cells(Rows.Count, 2).End(xlUp).row
    LastRowFrequence = sheetFrequence.Cells(Rows.Count, 2).End(xlUp).row
    endLineSchema = sheetSchema.Cells(Rows.Count, 18).End(xlUp).row
    endLineFranco = sheetFranco.Cells(Rows.Count, 18).End(xlUp).row
    endLineRuptures = sheetRuptures.Cells(Rows.Count, 18).End(xlUp).row
    endLineCouche = sheetCouche.Cells(Rows.Count, 18).End(xlUp).row
    endLineFrequence = sheetFrequence.Cells(Rows.Count, 18).End(xlUp).row
    
    If Not CStr(sheetSchema.Cells(9, 22).Value) = "" Then endLineSchema = endLineSchema + 1
    If Not CStr(sheetFranco.Cells(9, 22).Value) = "" Then endLineFranco = endLineFranco + 1
    If Not CStr(sheetRuptures.Cells(9, 22).Value) = "" Then endLineRuptures = endLineRuptures + 1
    If Not CStr(sheetCouche.Cells(9, 22).Value) = "" Then endLineCouche = endLineCouche + 1
    If Not CStr(sheetFrequence.Cells(9, 22).Value) = "" Then endLineFrequence = endLineFrequence + 1
    
    Let rangeSchema = "B" & startRow & ":P" & LastRowSchema
    Let rangeFranco = "B" & startRow & ":P" & LastRowFranco
    Let rangeRuptures = "B" & startRow & ":P" & LastRowRuptures
    Let rangeCouche = "B" & startRow & ":P" & LastRowCouche
    Let rangeFrequence = "B" & startRow & ":P" & LastRowFrequence
    
    Let RangeArraySchema = "T" & endLineSchema & ":AH" & endLineSchema
    Let RangeArrayFranco = "T" & endLineFranco & ":AH" & endLineFranco
    Let RangeArrayFrequence = "T" & endLineFrequence & ":AH" & endLineFrequence
    Let RangeArrayCouche = "T" & endLineCouche & ":AH" & endLineCouche
    Let RangeArrayRuptures = "T" & endLineRuptures & ":AH" & endLineRuptures
    
    sheetSchema.Range(rangeSchema).Copy
    sheetSchema.Paste Destination:=sheetSchema.Range(RangeArraySchema)
    
    sheetFranco.Range(rangeFranco).Copy
    sheetFranco.Paste Destination:=sheetFranco.Range(RangeArrayFranco)
    
    sheetFrequence.Range(rangeFrequence).Copy
    sheetFrequence.Paste Destination:=sheetFrequence.Range(RangeArrayFrequence)
    
    sheetCouche.Range(rangeCouche).Copy
    sheetCouche.Paste Destination:=sheetCouche.Range(RangeArrayCouche)
    
    sheetRuptures.Range(rangeRuptures).Copy
    sheetRuptures.Paste Destination:=sheetRuptures.Range(RangeArrayRuptures)
    
End Sub
Sub Suivi_LongTerme()

End Sub
Sub Archivage()
    
    If Not functionVariables = "activated" Then Variables
    On Error Resume Next
    
    Dim dateNow As Date
    dateNow = Date
    
    Dim endArchivesRange As String
    Dim tableRange As String
    
    For Each monitoring_Sheet In sheetsMonitoring
        
        monitoring_Sheet.Activate
        firstRowSuivi = 9
        lastRowSuivi = monitoring_Sheet.Cells(Rows.Count, 22).End(xlUp).row
        
        For i = firstRowSuivi To lastRowSuivi
            
            contestationReussie = CStr(monitoring_Sheet.Cells(i, 18).Value)
            contestationRefusee = CStr(monitoring_Sheet.Cells(i, 19).Value)
            datePreparation = Get_preparationDate_Of(monitoring_Sheet.Cells(i, 20).Value)
            endOfArchives = Archives.Cells(Rows.Count, 2).End(xlUp).row + 1
            endArchivesRange = "B" & endOfArchives & ":" & "P" & endOfArchives
            lineRange = "T" & i & ":" & "AH" & i
            FullLineRange = "R" & i & ":" & "AH" & i
            If Not monitoring_Sheet.Cells(i, 20) = 0 Then
                If Not contestationReussie = "" Then
                    If monitoring_Sheet.Name = "Monitoring ruptures" Then Archives.Cells(endOfArchives, 17).Value = "rupture": Archives.Cells(endOfArchives, 18).Value = "oui"
                    If monitoring_Sheet.Name = "Monitoring à la couche" Then Archives.Cells(endOfArchives, 17).Value = "couche": Archives.Cells(endOfArchives, 18).Value = "oui"
                    If monitoring_Sheet.Name = "Fréquence de livraison" Then Archives.Cells(endOfArchives, 17).Value = "frequence": Archives.Cells(endOfArchives, 18).Value = "oui"
                    If monitoring_Sheet.Name = "Franco" Then Archives.Cells(endOfArchives, 17).Value = "franco": Archives.Cells(endOfArchives, 18).Value = "oui"
                    If monitoring_Sheet.Name = "Schéma" Then Archives.Cells(endOfArchives, 17).Value = "schema": Archives.Cells(endOfArchives, 18).Value = "oui"
    
                    monitoring_Sheet.Range(lineRange).Copy
                    Archives.Paste Destination:=Archives.Range(endArchivesRange)
                    monitoring_Sheet.Range(FullLineRange).ClearContents
                    endOfArchives = endOfArchives + 1
                Else
                    If Not contestationRefusee = "" Then
                        If monitoring_Sheet.Name = "Monitoring ruptures" Then Archives.Cells(endOfArchives, 17).Value = "rupture": Archives.Cells(endOfArchives, 18).Value = "non"
                        If monitoring_Sheet.Name = "Monitoring à la couche" Then Archives.Cells(endOfArchives, 17).Value = "couche": Archives.Cells(endOfArchives, 18).Value = "non"
                        If monitoring_Sheet.Name = "Fréquence de livraison" Then Archives.Cells(endOfArchives, 17).Value = "frequence": Archives.Cells(endOfArchives, 18).Value = "non"
                        If monitoring_Sheet.Name = "Franco" Then Archives.Cells(endOfArchives, 17).Value = "franco": Archives.Cells(endOfArchives, 18).Value = "non"
                        If monitoring_Sheet.Name = "Schéma" Then Archives.Cells(endOfArchives, 17).Value = "schema": Archives.Cells(endOfArchives, 18).Value = "non"
                        monitoring_Sheet.Range(lineRange).Copy
                        Archives.Paste Destination:=Archives.Range(endArchivesRange)
                        monitoring_Sheet.Range(FullLineRange).ClearContents
                        endOfArchives = endOfArchives + 1
                    Else
                        If datePreparation < dateNow Then
                            If monitoring_Sheet.Name = "Monitoring ruptures" Then Archives.Cells(endOfArchives, 17).Value = "rupture": Archives.Cells(endOfArchives, 18).Value = "hors délai"
                            If monitoring_Sheet.Name = "Monitoring à la couche" Then Archives.Cells(endOfArchives, 17).Value = "couche": Archives.Cells(endOfArchives, 18).Value = "hors délai"
                            If monitoring_Sheet.Name = "Fréquence de livraison" Then Archives.Cells(endOfArchives, 17).Value = "frequence": Archives.Cells(endOfArchives, 18).Value = "hors délai"
                            If monitoring_Sheet.Name = "Franco" Then Archives.Cells(endOfArchives, 17).Value = "franco": Archives.Cells(endOfArchives, 18).Value = "hors délai"
                            If monitoring_Sheet.Name = "Schéma" Then Archives.Cells(endOfArchives, 17).Value = "schema": Archives.Cells(endOfArchives, 18).Value = "hors délai"
                            monitoring_Sheet.Range(lineRange).Copy
                            Archives.Paste Destination:=Archives.Range(endArchivesRange)
                            monitoring_Sheet.Range(FullLineRange).ClearContents
                            endOfArchives = endOfArchives + 1
                        End If
                    End If
                End If
            End If
        Next i
        endOfSuivi = lastRowSuivi
        For i = firstRowSuivi To lastRowSuivi
            If i = firstRowSuivi And monitoring_Sheet.Cells(i, 21).Value = "" And firstRowSuivi = lastRowSuivi Then Exit For
            If monitoring_Sheet.Cells(i, 21).Value = "" And i <= endOfSuivi Then
                lineRange = "R" & i & ":AH" & i
                j = i
                decalage = 0
                If Not i = endOfSuivi Then
                    Do
                        j = j + 1
                        Line = j
                        decalage = decalage + 1
                    Loop Until monitoring_Sheet.Cells(j, 21).Value <> ""
                Else
                    decalage = 1
                    Line = j
                End If

                lineRangeUnder = "R" & (Line) & ":AH" & (endOfSuivi)
                If decalage = 1 Then lastSuiviLine = "R" & (endOfSuivi - decalage) & ":AH" & (endOfSuivi)
                If decalage > 1 Then lastSuiviLine = "R" & (endOfSuivi - decalage + 1) & ":AH" & (endOfSuivi)
                If Not i = endOfSuivi Then monitoring_Sheet.Range(lineRangeUnder).Copy
                If Not i = endOfSuivi Then monitoring_Sheet.Paste Destination:=monitoring_Sheet.Range(lineRange)
                If Not i = endOfSuivi Then monitoring_Sheet.Range(lastSuiviLine).ClearContents
                endOfSuivi = endOfSuivi - decalage

            End If
        Next i

        Dim RangeTab As String
        RangeTab = "R8:AH" & endOfSuivi
        If monitoring_Sheet.Name = "Monitoring ruptures" Then sheetRuptures.ListObjects("TableauSuiviRuptures").Resize Range(RangeTab)
        If monitoring_Sheet.Name = "Monitoring à la couche" Then sheetCouche.ListObjects("TableauSuiviCouche").Resize Range(RangeTab)
        If monitoring_Sheet.Name = "Fréquence de livraison" Then sheetFrequence.ListObjects("TableauSuiviFrequence").Resize Range(RangeTab)
        If monitoring_Sheet.Name = "Franco" Then sheetFranco.ListObjects("TableauSuiviFranco").Resize Range(RangeTab)
        If monitoring_Sheet.Name = "Schéma" Then sheetSchema.ListObjects("TableauSuiviSchema").Resize Range(RangeTab)
    Next

End Sub

Sub prepaMail_Monitoring()

    Dim reason As String
    Dim listClient As New Scripting.Dictionary
    
    For Each monitoring_Sheet In sheetsMonitoring
        If Not monitoring_Sheet Is sheetValidation Then
            
            lastRowMonitoring = monitoring_Sheet.Cells(Rows.Count, 2).End(xlUp).row
            
            For ligneDeCommande = firstRowMonitoring To lastRowMonitoring
                SoldTo = CLng(monitoring_Sheet.Cells(ligneDeCommande, columnSoldTo_Monitoring).Value)
                If Not SoldTo = 0 Then
                    reason = sheetsMonitoring(monitoring_Sheet)
                    If Not listClient.Exists(SoldTo) Then
                        listClient.Add SoldTo, reason
                    Else
                        If Not reason = listClient.Item(SoldTo) Then
                            reason = "multi"
                            listClient.Item(SoldTo) = reason
                        End If
                    End If
                End If
            Next ligneDeCommande
        End If
    Next monitoring_Sheet
    
    For Each Client In listClient
        
        If listClient(Client) = "multi" Then
            multi = True
            Call mail(Get_Contact_Of(Client), "DANONE - Commande(s) bloquée(s) pour non respect des CGV", Mail_Monitoring_Of(Client))
        Else
            multi = False
            If listClient(Client) = "ruptures" Then
                Call mail(Get_Contact_Of(Client), "DANONE - Alerte rupture", Mail_Ruptures_Of(Client))
            ElseIf listClient(Client) = "couche" Then
                Call mail(Get_Contact_Of(Client), "DANONE - Rappel commande à la couche", Mail_Couche_Of(Client))
            ElseIf listClient(Client) = "frequence" Then
                Call mail(Get_Contact_Of(Client), "DANONE - Problème fréquence commandes", Mail_Frequence_Of(Client))
            ElseIf listClient(Client) = "franco" Then
                Call mail(Get_Contact_Of(Client), "DANONE - Rappel Franco commande", Mail_Franco_Of(Client))
            ElseIf listClient(Client) = "schema" Then
                Call mail(Get_Contact_Of(Client), "DANONE - Rappel Schéma de livraison", Mail_Schema_Of(Client))
            ElseIf listClient(Client) = "validation" Then
                Call mail(Get_Contact_Of(Client), "DANONE - Prise de rendez-vous livraison", Mail_Validation_Of(Client))
            End If
        End If
    
    Next Client
    
End Sub
