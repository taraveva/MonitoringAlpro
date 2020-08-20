Attribute VB_Name = "Fonctions__Mailing"
Public Function corpsMail_Franco(commandes() As Variant) As String
    'Un bullet point par Commande qui respecte pas le franco
    For i = 0 To UBound(commandes)
        PO = Get_PO_Of(commandes(i))
        dateDeLivraison = Get_DeliveryDate_Of(commandes(i))
        corpsMail_Franco = corpsMail_Franco + "<li>&nbsp;<strong>" + CStr(PO) + "</strong>&nbsp;: commande en livraison le<em> " + dateDeLivraison + "</em></li>"
    Next i
    corpsMail_Franco = corpsMail_Franco + "</ul>"
End Function
Public Function Mail_Franco_Of(ByVal Client As Long) As String
    
    If Not functionEstablish_listFranco = "activated" Then Establish_listFranco
    
    Dim commandes() As Variant
    commandes = listFranco(Client).Keys
    
    'Début du Mail
    Introduction = "<p>Bonjour,&nbsp;</p><p>Nous avons détecté que la/les commande(s) suivante(s) ne respectai(en)t pas le franco de 100 caisses&nbsp;:</p><ul>"
    IntroductionMulti = "<p><u>Franco:</u></p><ul>"
    'Corps du mail
    Corps = corpsMail_Franco(commandes)
    'Fin de Mail
    Conclusion = "<p><strong>Par cons&eacute;quent, votre commande est bloqu&eacute;e dans notre syst&egrave;me.</strong></p><p>Merci de nous faire un retour dans un d&eacute;lai de 24h, sur la quantit&eacute; souhait&eacute;e respectant le franco.</p><p>Cordialement,</p><p>Le Service Client Danone <br />01 49 48 56 41</p>"
    
    If Not multi = True Then
        Mail_Franco_Of = Introduction + Corps + Conclusion
    Else
        Mail_Franco_Of = IntroductionMulti + Corps
    End If
    
End Function
Public Function Establish_listFranco()

    If Not functionVariables = "activated" Then Variables
    
    lastRowMonitoring = sheetFranco.Cells(Rows.Count, 2).End(xlUp).row
    Dim orderList As Scripting.Dictionary
    
    'Pour chaque ligne de commande ne respectant pas le franco
    For ligneDeCommande = firstRowMonitoring To lastRowMonitoring
        
        Order = CLng(sheetFranco.Cells(ligneDeCommande, columnOrder_Monitoring).Value)
        PurchaseOrder = CStr(sheetFranco.Cells(ligneDeCommande, columnPO_Monitoring).Value)
        orderedQuantity = CInt(sheetFranco.Cells(ligneDeCommande, columnOrderQty_Monitoring).Value)
        SoldTo = CLng(sheetFranco.Cells(ligneDeCommande, columnSoldTo_Monitoring).Value)
        If Not SoldTo = 0 Then
            'Création de la table associant SoldTo et Order
            If Not listFranco.Exists(SoldTo) Then
                Set orderList = New Scripting.Dictionary
                orderList.Add Order, orderedQuantity
                listFranco.Add SoldTo, orderList
            Else
                listFranco(SoldTo).Add Order, deliveryDate
            End If
        End If

    Next ligneDeCommande
    
    functionEstablish_listFranco = "activated"
    
End Function
Public Function corpsMail_Schema(commandes() As Variant) As String
        
    For i = 0 To UBound(commandes)
        PO = Get_PO_Of(commandes(i))
        dateDeLivraison = Get_DeliveryDate_Of(commandes(i))
        corpsMail_Schema = corpsMail_Schema + "<li>&nbsp;Commande <strong>" + CStr(PO) + "</strong>&nbsp; en livraison le<em> " + dateDeLivraison + "</em></li>"
    Next i
    corpsMail_Schema = corpsMail_Schema + "</ul>"
End Function
Public Function Mail_Schema_Of(ByVal Client As Long) As String
    
    If Not functionEstablish_listSchema = "activated" Then Establish_listSchema
    
    Dim commandes() As Variant
    commandes = listSchema(Client).Keys
    
    'Début du Mail
    Introduction = "<p>Bonjour,&nbsp;</p><p>Nous avons détecté que la/les commande(s) suivante(s) ne respectai(en)t pas le schéma de livraison&nbsp;:</p><ul>"
    IntroductionMulti = "<p><u>Schéma de livraison non respecté:</u></p><ul>"
    'Corps du mail
    Corps = corpsMail_Schema(commandes)
    'Fin de Mail
    Conclusion = "<p><strong>Par cons&eacute;quent, votre commande est bloqu&eacute;e dans notre syst&egrave;me.</strong></p><p>Merci de nous faire un retour dans un d&eacute;lai de 24h, sur une nouvelle date de livraison.</p><p>Cordialement,</p><p>Le Service Client Danone <br />01 49 48 56 41</p>"
    
    If Not multi = True Then
        Mail_Schema_Of = Introduction + Corps + Conclusion
    Else
        Mail_Schema_Of = IntroductionMulti + Corps
    End If
    
End Function
Public Function Establish_listSchema()
    
    If Not functionVariables = "activated" Then Variables
    
    lastRowMonitoring = sheetSchema.Cells(Rows.Count, 2).End(xlUp).row
    Dim orderList As Scripting.Dictionary
        
    For ligneDeCommande = firstRowMonitoring To lastRowMonitoring
        
        Order = CLng(sheetSchema.Cells(ligneDeCommande, columnOrder_Monitoring).Value)
        SoldTo = CLng(sheetSchema.Cells(ligneDeCommande, columnSoldTo_Monitoring).Value)
        deliveryDate = CDate(sheetSchema.Cells(ligneDeCommande, columnRequestedDeliveryDate_Monitoring).Value)
        If Not SoldTo = 0 Then
            'Création de la table associant SoldTo et Order
            If Not listSchema.Exists(SoldTo) Then
                Set orderList = New Scripting.Dictionary
                orderList.Add Order, deliveryDate
                listSchema.Add SoldTo, orderList
            Else
                listSchema(SoldTo).Add Order, deliveryDate
            End If
        End If

    Next ligneDeCommande

    functionEstablish_listSchema = "activated"
    
End Function
Public Function corpsMail_Couche(commandes() As Variant) As String
    
    For i = 0 To UBound(commandes)
        PO = Get_PO_Of(commandes(i, 0))
        corpsMail_Couche = corpsMail_Couche + "<li>Commande " + PO + "</li><ul>"
        For j = 0 To UBound(commandes(i, 1))
            ean13 = Get_EAN_Of(commandes(i, 1)(j))
            libelle = Get_Libelle_Of(commandes(i, 1)(j))
            qteCommande = Get_OrderedQty_Of(commandes(i, 0), commandes(i, 1)(j))
            Product = commandes(i, 1)(j)
            corpsMail_Couche = corpsMail_Couche + "<li>&nbsp;<strong>Produit " + CStr(Product) + "</strong>&nbsp;: " + ean13 + " - " + libelle + " - <em>" + qteCommande + " caisse(s)</em></li>"
        Next j
        corpsMail_Couche = corpsMail_Couche + "</ul>"
    Next i
    corpsMail_Couche = corpsMail_Couche + "</ul>"
End Function
Public Function Mail_Couche_Of(ByVal Client As Long) As String
    
    If Not functionEstablish_listCouche = "activated" Then Establish_listCouche
    
    Dim commandesCondensed() As Variant
    Dim produits() As Variant
    commandesCondensed = listCouche(Client).Keys
    ReDim commandes(UBound(commandesCondensed), 1) As Variant
    j = 0
    
    For i = 0 To UBound(commandesCondensed)
        produits = listCouche(Client)(commandesCondensed(i)).Keys
        commandes(j, 0) = commandesCondensed(i)
        commandes(j, 1) = produits
        j = j + 1
    Next i
    'Début du Mail
    IntroductionClassique = "<p>Bonjour,&nbsp;</p><p>Sur la/les commande(s) suivante(s), un ou plusieurs produits ne respectent pas la commande &agrave; la couche ou &agrave; la palette&nbsp;:</p><ul>"
    IntroductionDisplay = "<p>Bonjour,&nbsp;</p><p>Nous avons bien reçu votre commande <strong>" + CStr(PO) + "</strong>, cependant les quantités d'un ou plusieurs produits ne respectent pas le critère d'un multiple de 4:</p><ul>"
    IntroductionMulti = "<p><u>Critère de couche non respecté:</u></p><ul>"
    'Corps du mail
    Corps = corpsMail_Couche(commandes)
    'Fin de Mail
    ConclusionClassique = "<p><strong><em>Par cons&eacute;quent, votre commande est bloqu&eacute;e dans notre syst&egrave;me.</em></strong></p><p>Merci de nous faire un retour dans un d&eacute;lai de 24h, sur la quantit&eacute; souhait&eacute;e respectant le crit&egrave;re de couche ou de palettes.</p><p>Cordialement,</p><p>Le Service Client Danone <br />01 49 48 56 41</p>"
    ConclusionDisplay = "<p><strong><em>Par cons&eacute;quent, votre commande est bloqu&eacute;e dans notre syst&egrave;me.</em></strong></p><p>Merci de revenir vers nous avec la quantité corrigée dans un délai de 24h.</p><p>Cordialement,</p><p>Le Service Client Danone <br />01 49 48 56 41</p>"

    If Not multi = True Then
            Mail_Couche_Of = IntroductionClassique + Corps + ConclusionClassique
        Else
            Mail_Couche_Of = IntroductionMulti + Corps
        End If
    
End Function
Public Function Establish_listCouche()
    
    If Not functionVariables = "activated" Then Variables
    
    lastRowMonitoring = sheetCouche.Cells(Rows.Count, 2).End(xlUp).row
    
    Dim orderList As Scripting.Dictionary
    Dim productList As Scripting.Dictionary
        
    For ligneDeCommande = firstRowMonitoring To lastRowMonitoring
        
        Product = CLng(sheetCouche.Cells(ligneDeCommande, columnMaterial_Monitoring).Value)
        Order = CLng(sheetCouche.Cells(ligneDeCommande, columnOrder_Monitoring).Value)
        orderedQty = CInt(sheetCouche.Cells(ligneDeCommande, columnOrderQty_Monitoring).Value)
        SoldTo = CLng(sheetCouche.Cells(ligneDeCommande, columnSoldTo_Monitoring).Value)
        If Not SoldTo = 0 Then
            If Not listCouche.Exists(SoldTo) Then
                Set productList = New Scripting.Dictionary
                Set orderList = New Scripting.Dictionary
                productList.Add Product, orderedQty
                orderList.Add Order, productList
                listCouche.Add SoldTo, orderList
            Else
                If Not listCouche(SoldTo).Exists(Order) Then
                    Set productList = New Scripting.Dictionary
                    productList.Add Product, orderedQty
                    listCouche(SoldTo).Add Order, productList
                Else
                    listCouche(SoldTo)(Order).Add Product, orderedQty
                End If
            End If
        End If
    Next ligneDeCommande

    functionEstablish_listCouche = "activated"
    
End Function
Public Function corpsMail_Ruptures(commandes() As Variant) As String
    
    For i = 0 To UBound(commandes)
        PO = Get_PO_Of(commandes(i, 0))
        corpsMail_Ruptures = corpsMail_Ruptures + "<li><strong>Commande " + PO + "</li></strong><ul>"
        For j = 0 To UBound(commandes(i, 1))
            dateRAN = Get_RAN_Of(commandes(i, 1)(j))
            ean13 = Get_EAN_Of(commandes(i, 1)(j))
            libelle = Get_Libelle_Of(commandes(i, 1)(j))
            qteCommande = Get_OrderedQty(commandes(i, 1)(j))
            corpsMail_Ruptures = corpsMail_Ruptures + "<li>&nbsp;<em><strong> Produit " + CStr(Product) + "</strong>&nbsp;: " + ean13 + " - " + libelle + " - " + qteCommande + " caisse(s)</em> || RAN en livraison le " + dateRAN + "</li>"
        Next j
        corpsMail_Ruptures = corpsMail_Ruptures + "</ul>"
    Next i
    corpsMail_Ruptures = corpsMail_Ruptures + "</ul>"
End Function
Public Function Mail_Ruptures_Of(ByVal Client As Long) As String
    
    If Not functionEstablish_listRuptures = "activated" Then Establish_listRuptures
    
    Dim commandesCondensed() As Variant
    Dim produits() As Variant
    commandesCondensed = listRuptures(Client).Keys
    ReDim commandes(UBound(commandesCondensed), 1) As Variant
    j = 0
    
    For i = 0 To UBound(commandesCondensed)
        produits = listCouche(Client)(commandesCondensed(i)).Keys
        commandes(j, 0) = commandesCondensed(i)
        commandes(j, 1) = produits
        j = j + 1
    Next i
    
    'Début du Mail
    Introduction = "<p>Bonjour,&nbsp;</p><p>Une ou plusieurs de vos commandes comportent un ou plusieurs produits en rupture&nbsp;:</p><ul>"
    IntroductionMulti = "<p><u>Alerte ruptures:</u></p><ul>"
    'Corps du mail
    Corps = corpsMail_Ruptures(commandes)
    'Fin de Mail
    Conclusion = "<p><strong><em>Par cons&eacute;quent, votre commande est bloqu&eacute;e dans notre syst&egrave;me.</em></strong></p><p>Merci de nous faire un retour dans un d&eacute;lai de 24h pour que nous puissons arranger une autre date de livraison.</p><p>Cordialement,</p><p>Le Service Client Danone <br />01 49 48 56 41</p>"
   If Not multi = True Then
        Mail_Ruptures_Of = Introduction + Corps + Conclusion
    Else
        Mail_Ruptures_Of = IntroductionMulti + Corps
    End If
    
End Function
Public Function Establish_listRuptures()
    
    If Not functionVariables = "activated" Then Variables
    
    lastRowMonitoring = sheetRuptures.Cells(Rows.Count, 2).End(xlUp).row
    
    Dim orderList As Scripting.Dictionary
    Dim productList As Scripting.Dictionary
        
    For ligneDeCommande = firstRowMonitoring To lastRowMonitoring
        
        Product = CLng(sheetRuptures.Cells(ligneDeCommande, columnMaterial_Monitoring).Value)
        Order = CLng(sheetRuptures.Cells(ligneDeCommande, columnOrder_Monitoring).Value)
        orderedQty = CInt(sheetRuptures.Cells(ligneDeCommande, columnOrderQty_Monitoring).Value)
        SoldTo = CLng(sheetRuptures.Cells(ligneDeCommande, columnSoldTo_Monitoring).Value)
        If Not SoldTo = 0 Then
            If Not listRuptures.Exists(SoldTo) Then
                Set productList = New Scripting.Dictionary
                Set orderList = New Scripting.Dictionary
                productList.Add Product, orderedQty
                orderList.Add Order, productList
                listRuptures.Add SoldTo, orderList
            Else
                If Not listRuptures(SoldTo).Exists(Order) Then
                    Set productList = New Scripting.Dictionary
                    productList.Add Product, orderedQty
                    listRuptures(SoldTo).Add Order, productList
                Else
                    listRuptures(SoldTo)(Order).Add Product, orderedQty
                End If
            End If
        End If
        
    Next ligneDeCommande

    functionEstablish_listRuptures = "activated"
    
End Function
Public Function corpsMail_Validation(commandes() As Variant) As String
    'Un bullet point par Commande qui respecte pas le franco
    For i = 0 To UBound(commandes)
        PO = Get_PO_Of(commandes(i))
        dateDeLivraison = Get_DeliveryDate_Of(commandes(i))
        corpsMail_Validation = corpsMail_Validation + "<li>&nbsp;<strong>" + CStr(PO) + "</strong>&nbsp;: commande en livraison le<em> " + dateDeLivraison + "</em></li>"
    Next i
    corpsMail_Validation = corpsMail_Validation + "</ul>"
End Function
Public Function Mail_Validation_Of(ByVal Client As Long) As String
    
    If Not functionEstablish_listValidation = "activated" Then Establish_listValidation
    
    Dim commandes() As Variant
    commandes = listValidation(Client).Keys
    HourOfDelivery = Get_deliveryHour_Of(Client)
    'Début du Mail
    Introduction = "<p>Bonjour,&nbsp;</p><p>La/Les commande(s) suivante(s) ont été validées dans nos systèmes:&nbsp;:</p><ul>"
    IntroductionMulti = "<p><u>Commandes validées:</u></p><ul>"
    'Corps du mail
    Corps = corpsMail_Validation(commandes)
    'Fin de Mail
    Conclusion = "<p>Le transporteur STEF livrera ces/cette commande(s) à <strong>" + HourOfDelivery + ". </strong></p><p>Cordialement,</p><p>Le Service Client Danone <br />01 49 48 56 41</p>"
    
    If Not multi = True Then
        Mail_Validation_Of = Introduction + Corps + Conclusion
    Else
        Mail_Validation_Of = IntroductionMulti + Corps
    End If
    
End Function
Public Function Establish_listValidation()

    If Not functionVariables = "activated" Then Variables
    
    lastRowMonitoring = sheetValidation.Cells(Rows.Count, 2).End(xlUp).row
    Dim orderList As Scripting.Dictionary
    
    'Pour chaque ligne de commande ne respectant pas le franco
    For ligneDeCommande = firstRowMonitoring To lastRowMonitoring
        
        Order = CLng(sheetValidation.Cells(ligneDeCommande, columnOrder_Monitoring).Value)
        PurchaseOrder = CStr(sheetValidation.Cells(ligneDeCommande, columnPO_Monitoring).Value)
        orderedQuantity = CInt(sheetValidation.Cells(ligneDeCommande, columnOrderQty_Monitoring).Value)
        SoldTo = CLng(sheetValidation.Cells(ligneDeCommande, columnSoldTo_Monitoring).Value)
        If Not SoldTo = 0 Then
            'Création de la table associant SoldTo et Order
            If Not listValidation.Exists(SoldTo) Then
                Set orderList = New Scripting.Dictionary
                orderList.Add Order, orderedQuantity
                listValidation.Add SoldTo, orderList
            Else
                listValidation(SoldTo).Add Order, deliveryDate
            End If
        End If

    Next ligneDeCommande
    
    functionEstablish_listValidation = "activated"
    
End Function
Public Function corpsMail_Frequence(commandes() As Variant) As String
    Dim dateDuJour As Date
    dateDuJour = Date
    
    For i = 0 To UBound(commandes)
        corpsMail_Frequence = corpsMail_Frequence + "<li><strong>Semaine " + CStr(commandes(i, 0)) + "</li></strong><ul>"
        For j = 0 To UBound(commandes(i, 1))
            deliveryDate = Get_DeliveryDate_Of(commandes(i, 1)(j))
            PO = Get_PO_Of(commandes(i, 1)(j))
            prepDate = Get_preparationDate_Of(commandes(i, 1)(j))
            If prepDate <= dateDuJour Then
                corpsMail_Frequence = corpsMail_Frequence + "<li>La commande<em> " + CStr(PO) + " </em>en livraison le<em> " + CStr(deliveryDate) + "</em> --> Commande validée</li>"
            Else
                corpsMail_Frequence = corpsMail_Frequence + "<li>La commande<em> " + CStr(PO) + " </em>en livraison le<em> " + CStr(deliveryDate) + "</em> --> <font color=""red"">Commande bloquée</font></li>"
            End If
        Next j
        corpsMail_Frequence = corpsMail_Frequence + "</ul>"
    Next i
    corpsMail_Frequence = corpsMail_Frequence + "</ul>"
End Function
Public Function Mail_Frequence_Of(ByVal Client As Long) As String
    
    If Not functionEstablish_listFrequence = "activated" Then Establish_listFrequence
    
    Dim weekCondensed() As Variant
    Dim commandes() As Variant
    weekCondensed = listFrequence(Client).Keys
    ReDim weeks(UBound(weekCondensed), 1) As Variant
    j = 0
    
    For i = 0 To UBound(weekCondensed)
        commandes = listFrequence(Client)(weekCondensed(i)).Keys
        weeks(j, 0) = weekCondensed(i)
        weeks(j, 1) = commandes
        j = j + 1
    Next i
    
    'Début du Mail
    Introduction = "<p>Bonjour,</p><p>Les conditions g&eacute;n&eacute;rales de ventes des Boissons Alpro et Provamel pr&eacute;voient qu&rsquo;une seule livraison ait lieu par semaine.</p><p>Or nous constatons que sur votre entrep&ocirc;t de<strong> " + CStr(Get_Entrepot_Of(Client)) + "</strong>, nous avons:</p>"
    IntroductionMulti = "<p><u>Fréquence de livraison non respectée:</u></p><ul>"
    'Corps du mail
    Corps = corpsMail_Frequence(weeks)
    'Fin de Mail
    Conclusion = "<p>Nous ne pouvons vous livrer plusieurs fois cons&eacute;cutives. Par cons&eacute;quent, vos commandes sont bloqu&eacute;es.</p><p>Merci de revenir vers nous avec une seule date de livraison par semaine dans un d&eacute;lai de 24h.</p><p>Cordialement,</p><p>Le Service Client Danone <br />01 49 48 56 41</p>"
    If Not multi = True Then
        Mail_Frequence_Of = Introduction + Corps + Conclusion
    Else
        Mail_Frequence_Of = IntroductionMulti + Corps
    End If
    
End Function
Public Function Establish_listFrequence()
    
    If Not functionVariables = "activated" Then Variables
    
    lastRowMonitoring = sheetFrequence.Cells(Rows.Count, 2).End(xlUp).row
    
    Dim weekList As Scripting.Dictionary
    Dim orderList As Scripting.Dictionary
        
    For ligneDeCommande = firstRowMonitoring To lastRowMonitoring
        
        deliveryDate = CDate(sheetFrequence.Cells(ligneDeCommande, columnRequestedDeliveryDate_Monitoring).Value)
        SoldTo = CLng(sheetFrequence.Cells(ligneDeCommande, columnSoldTo_Monitoring).Value)
        Order = CStr(sheetFrequence.Cells(ligneDeCommande, columnOrder_Monitoring).Value)
        deliveryDateWeek = CStr(WorksheetFunction.WeekNum(deliveryDate, vbMonday))
        If Not SoldTo = 0 Then
            If Not listFrequence.Exists(SoldTo) Then
                Set orderList = New Scripting.Dictionary
                Set weekList = New Scripting.Dictionary
                orderList.Add Order, deliveryDate
                weekList.Add deliveryDateWeek, orderList
                listFrequence.Add SoldTo, weekList
            Else
                If Not listFrequence(SoldTo).Exists(deliveryDateWeek) Then
                    Set orderList = New Scripting.Dictionary
                    orderList.Add Order, deliveryDate
                    listFrequence(SoldTo).Add deliveryDateWeek, orderList
                Else
                    listFrequence(SoldTo)(deliveryDateWeek).Add Order, deliveryDate
                End If
            End If
        End If
        
    Next ligneDeCommande

    functionEstablish_listFrequence = "activated"
    
End Function
Sub mail(destinataire, sujet, contenuMail)
    
    Dim mail As Variant
    
    Set mail = CreateObject("Outlook.Application")
    
    With mail.CreateItem(olMailItem)
        .To = destinataire
        .Subject = sujet
        .HTMLBody = contenuMail
        .Display
    End With

End Sub
Public Function Mail_Monitoring_Of(ByVal Client As Long) As String
    If Not functionEstablish_listRuptures = "activated" Then Establish_listRuptures
    If Not functionEstablish_listCouche = "activated" Then Establish_listCouche
    If Not functionEstablish_listFrequence = "activated" Then Establish_listFrequence
    If Not functionEstablish_listFranco = "activated" Then Establish_listFranco
    If Not functionEstablish_listSchema = "activated" Then Establish_listSchema
    If Not functionEstablish_listValidation = "activated" Then Establish_listValidation
    
    Introduction = "<p>Bonjour,</p>"
    Corps = "<p>Nous avons noté que une ou plusieurs de vos commandes ne respectent pas les CGV:</p>"
    If listRuptures.Exists(Client) Then Corps = Corps + Mail_Ruptures_Of(Client)
    If listCouche.Exists(Client) Then Corps = Corps + Mail_Couche_Of(Client)
    If listFranco.Exists(Client) Then Corps = Corps + Mail_Franco_Of(Client)
    If listFrequence.Exists(Client) Then Corps = Corps + Mail_Frequence_Of(Client)
    If listSchema.Exists(Client) Then Corps = Corps + Mail_Schema_Of(Client)
    
    Conclusion = "</ul><p><strong><em>Par cons&eacute;quent, votre commande est bloqu&eacute;e dans notre syst&egrave;me.</em></strong></p><p>Merci de nous faire un retour dans un d&eacute;lai de 24h pour que nous puissons arranger une autre date de livraison.</p><p>Cordialement,</p><p>Le Service Client Danone <br />01 49 48 56 41</p>"
    Mail_Monitoring_Of = Introduction + Corps + Conclusion
    
End Function
