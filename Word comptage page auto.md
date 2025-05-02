# Word comptage page auto

## **📌 Étape 1 : Ouvrir l’éditeur VBA**

1. **Ouvrir le document Word.**
2. **Appuier sur `ALT + F11`** pour ouvrir l'éditeur VBA.
3. **Dans l’éditeur VBA :**
    
    
    - Aller dans `Insertion` &gt; `Module`.
    - Un nouveau module (`Module1`) apparaît dans le panneau de gauche.

---

## **📌 Étape 2 : Ajouter le code VBA**

1. **Copier tout le code suivant** :

```vbscript
Sub MettreAJourTableauPages()
    Dim signets As Variant
    Dim i As Integer
    Dim nbPages As Integer
    Dim signetDebut As String, signetFin As String, signetResultat As String

    ' Liste des sections avec leurs signets
    signets = Array( _
        Array("DebutPageGarde", "FinPageGarde", "ResultatPageGarde"), _
        Array("DebutNC", "FinNC", "ResultatNC"), _
        Array("DebutNCP", "FinNCP", "ResultatNCP"), _
        Array("DebutRemarque", "FinRemarque", "ResultatRemarque"), _
        Array("DebutRemarqueP", "FinRemarqueP", "ResultatRemarqueP"), _
        Array("DebutAxe", "FinAxe", "ResultatAxe"), _
        Array("DebutSynthese", "FinSynthese", "ResultatSynthese"), _
        Array("DebutPersonne", "FinPersonne", "ResultatPersonne"), _
        Array("DebutDoc", "FinDoc", "ResultatDoc") _
    )

    ' Boucle pour traiter chaque section
    For i = LBound(signets) To UBound(signets)
        signetDebut = signets(i)(0)
        signetFin = signets(i)(1)
        signetResultat = signets(i)(2)

        ' Compter le nombre de pages entre les deux signets
        nbPages = PagesEntreSignets(signetDebut, signetFin)

        ' Insérer le résultat à l'emplacement du signet de résultat
        If nbPages <> -1 Then
            InsererTexteDansSignet signetResultat, CStr(nbPages)
        End If
    Next i

    ' Calculer le total des pages du document
    Dim totalPages As Integer
    totalPages = ActiveDocument.ComputeStatistics(wdStatisticPages)
    InsererTexteDansSignet "ResultatTotal", CStr(totalPages) & " pages"

    MsgBox "Mise à jour du tableau terminée !", vbInformation, "Succès"
End Sub

Function PagesEntreSignets(signetDebut As String, signetFin As String) As Integer
    Dim debutPage As Integer, finPage As Integer

    ' Vérifier si les signets existent
    If Not ActiveDocument.Bookmarks.Exists(signetDebut) Or Not ActiveDocument.Bookmarks.Exists(signetFin) Then
        PagesEntreSignets = -1
        Exit Function
    End If

    ' Récupérer les numéros de page
    debutPage = ActiveDocument.Bookmarks(signetDebut).Range.Information(wdActiveEndPageNumber)
    finPage = ActiveDocument.Bookmarks(signetFin).Range.Information(wdActiveEndPageNumber)

    ' Retourner le nombre de pages couvertes
    PagesEntreSignets = finPage - debutPage + 1
End Function

Sub InsererTexteDansSignet(nomSignet As String, texte As String)
    Dim r As Range

    ' Vérifier si le signet existe
    If ActiveDocument.Bookmarks.Exists(nomSignet) Then
        Set r = ActiveDocument.Bookmarks(nomSignet).Range
        r.Text = texte
        ' Réinsérer le signet après remplacement du texte
        ActiveDocument.Bookmarks.Add nomSignet, r
    End If
End Sub

```

- **Coller le code dans `Module1`.**
- **Enregistrer (`CTRL + S`).**

## **📌 Étape 3 : Ajouter les signets dans Word**

Avant d’exécuter la macro, il faut insérer des signets dans le document pour identifier les zones à analyser.

1. **Ouvrir le document Word.**
2. **Placer le curseur avant le début d’une section** (ex. : Page de garde).
3. **Ajouter un signet :**
    
    
    - Aller dans `Insertion` &gt; `Lien` &gt; `Signet`.
    - Donner un nom précis (ex. : `DebutPageGarde`).
    - Cliquer sur `Ajouter`.
4. **Placer le curseur à la fin de cette section** et ajoute un signet `FinPageGarde`.
5. **Dans la cellule du tableau où doit apparaître le nombre de pages, insèrer un signet `ResultatPageGarde`.**
6. **Répèter pour chaque section :**
    
    
    - **Non-Conformité(s)** : `DebutNC`, `FinNC`, `ResultatNC`
    - **Non-Conformité(s) du rapport précédent** : `DebutNCP`, `FinNCP`, `ResultatNCP`
    - **Remarque(s)** : `DebutRemarque`, `FinRemarque`, `ResultatRemarque`
    - **Remarques(s) du rapport précédent** : `DebutRemarqueP`, `FinRemarqueP`, `ResultatRemarqueP`
    - **Axe(s) d’amélioration** : `DebutAxe`, `FinAxe`, `ResultatAxe`
    - **Synthèse de l'équipe d'audit** : `DebutSynthese`, `FinSynthese`, `ResultatSynthese`
    - **Personne(s) rencontrée(s)** : `DebutPersonne`, `FinPersonne`, `ResultatPersonne`
    - **Documents et exemples examinés** : `DebutDoc`, `FinDoc`, `ResultatDoc`
    - **Total des pages** : `ResultatTotal`

⚠️ **Les noms des signets doivent être exactement les mêmes que ceux du code VBA.**

**[![image.png](https://bsk.hackox.synology.me/uploads/images/gallery/2025-05/scaled-1680-/scnO2W7B8flE1kms-image.png)](https://bsk.hackox.synology.me/uploads/images/gallery/2025-05/scnO2W7B8flE1kms-image.png)**

**[![image.png](https://bsk.hackox.synology.me/uploads/images/gallery/2025-05/scaled-1680-/4V7lRCz6CmeChr2F-image.png)](https://bsk.hackox.synology.me/uploads/images/gallery/2025-05/4V7lRCz6CmeChr2F-image.png)**

---

## **📌 Étape 4 : Exécuter la macro**

1. **Retourner dans l'éditeur VBA (`ALT + F11`).**
2. **Sélectionner la macro `MettreAJourTableauPages`.**
3. **Cliquer sur "Exécuter" (`F5`).**
4. **Les résultats s’insèrent automatiquement dans le tableau à l’emplacement des signets.** 🎉

---

## **📊 Exemple de tableau après exécution**

<div class="overflow-x-auto contain-inline-size" id="bkmrk-composition-du-rappo"><table data-end="5425" data-start="4793"><thead data-end="4856" data-start="4793"><tr data-end="4856" data-start="4793"><th data-end="4835" data-start="4793">**Composition du rapport d'audit**</th><th data-end="4856" data-start="4835">**Nbre de pages**</th></tr></thead><tbody data-end="5425" data-start="4908"><tr data-end="4957" data-start="4908"><td>Page de garde</td><td>**1**</td></tr><tr data-end="5008" data-start="4958"><td>Non-Conformité(s)</td><td>**3**</td></tr><tr data-end="5066" data-start="5009"><td>Non-Conformité(s) du rapport précédent</td><td>**2**</td></tr><tr data-end="5117" data-start="5067"><td>Remarque(s)</td><td>**2**</td></tr><tr data-end="5170" data-start="5118"><td>Remarques(s) du rapport précédent</td><td>**1**</td></tr><tr data-end="5221" data-start="5171"><td>Axe(s) d’amélioration</td><td>**1**</td></tr><tr data-end="5272" data-start="5222"><td>Synthèse de l'équipe d'audit</td><td>**2**</td></tr><tr data-end="5323" data-start="5273"><td>Personne(s) rencontrée(s)</td><td>**2**</td></tr><tr data-end="5374" data-start="5324"><td>Documents et exemples examinés</td><td>**6**</td></tr><tr data-end="5425" data-start="5375"><td>**TOTAL**</td><td>**20 pages**</td></tr></tbody></table>

</div>---

## **✅ Avantages de cette méthode**

✔ **Automatisé** : Mets à jour tout le tableau en un clic.  
✔ **Fiable** : Utilise les signets pour localiser les sections.  
✔ **Flexible** : Facile à adapter pour d’autres rapports.
