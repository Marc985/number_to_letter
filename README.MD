# README : Script VBA pour convertir un nombre en lettres

## Objectif

Ce script VBA pour Excel permet de convertir un nombre écrit dans une cellule en sa version textuelle en français. Lorsqu'un nombre est saisi dans une cellule, sa représentation en lettres est automatiquement affichée dans la cellule adjacente.

---

## 1. Ajouter le script VBA

### Étapes :

1. **Ouvrir l'éditeur VBA** :
   - Ouvrez Excel.
   - Appuyez sur `Alt + F11` pour ouvrir l'éditeur VBA.

2. **Ajouter un module pour le code VBA** :
   - Allez dans `Insertion > Module`.
   - Collez le code suivant dans le module.

### Code VBA à ajouter dans le module :

```vba
Function NombreEnTexte(ByVal Nombre As Double) As String
    Dim Unites As Variant, Dizaines As Variant
    Dim PartieEntiere As Long, PartieDecimale As Long
    Dim Texte As String, PartieDecimaleTexte As String
    
    ' Définir les tableaux pour les unités et les dizaines
    Unites = Array("", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf", "dix", _
                   "onze", "douze", "treize", "quatorze", "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf")
    Dizaines = Array("", "", "vingt", "trente", "quarante", "cinquante", "soixante", "soixante-dix", "quatre-vingt", "quatre-vingt-dix")
    
    ' Séparer la partie entière et la partie décimale
    PartieEntiere = Int(Nombre)
    PartieDecimale = Round((Nombre - PartieEntiere) * 100)
    
    ' Convertir la partie entière
    Texte = ConvertirCentaines(PartieEntiere, Unites, Dizaines)
    
    ' Ajouter la partie décimale (si elle existe)
    If PartieDecimale > 0 Then
        PartieDecimaleTexte = ConvertirCentaines(PartieDecimale, Unites, Dizaines)
        Texte = Texte & " virgule " & PartieDecimaleTexte
    End If
    
    NombreEnTexte = Texte
End Function

Function ConvertirCentaines(ByVal Nombre As Long, Unites As Variant, Dizaines As Variant) As String
    Dim Centaines As Long, Reste As Long
    Dim Texte As String
    
    If Nombre = 0 Then
        ConvertirCentaines = "zéro"
        Exit Function
    End If
    
    ' Gérer les centaines
    Centaines = Int(Nombre / 100)
    Reste = Nombre Mod 100
    If Centaines > 0 Then
        If Centaines > 1 Then
            Texte = Unites(Centaines) & " cent"
        Else
            Texte = "cent"
        End If
        If Reste = 0 Then
            Texte = Texte & "s" ' Ajouter le "s" si c'est un nombre rond
        End If
    End If
    
    ' Ajouter les dizaines et unités
    If Reste > 0 Then
        If Centaines > 0 Then Texte = Texte & " "
        Texte = Texte & ConvertirDizaines(Reste, Unites, Dizaines)
    End If
    
    ConvertirCentaines = Texte
End Function

Function ConvertirDizaines(ByVal Nombre As Long, Unites As Variant, Dizaines As Variant) As String
    Dim Texte As String
    If Nombre < 20 Then
        Texte = Unites(Nombre)
    Else
        Dim Dizaine As Long, Unite As Long
        Dizaine = Int(Nombre / 10)
        Unite = Nombre Mod 10
        Texte = Dizaines(Dizaine)
        If Dizaine = 7 Or Dizaine = 9 Then
            Texte = Dizaines(Dizaine - 1) & "-" & Unites(10 + Unite)
        ElseIf Unite > 0 Then
            If Dizaine = 8 Then
                Texte = Texte & "-" & Unites(Unite) ' pas d


3. ** Fonctionnement
        Entrée : Lorsque vous entrez un nombre dans la colonne A (par exemple, dans la cellule A1), le script écrit automatiquement sa version en lettres dans la cellule de droite (B1).
        Sortie : Si vous supprimez le contenu de A1, la cellule B1 sera également effacée.
        Le script vérifie que la valeur saisie est un nombre pour éviter des erreurs de conversion