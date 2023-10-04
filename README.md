|*Doc. #*|*Rédacteur*|*Création*|*Mise à jour*|
|:---:|:---:|---:|:---|
|***0024***|*Loïc Drouet*|_Mercredi 04 octobre 2023_|_Mercredi 04 octobre 2023_|


<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=true} -->

<!-- code_chunk_output -->

1. [_[VBA]_ Afficher un retour à la ligne dans la cellule à partir du code VBA](#_vba_-afficher-un-retour-à-la-ligne-dans-la-cellule-à-partir-du-code-vba)
2. [_[VBA]_ Continuer le code VBA sur une nouvelle ligne](#_vba_-continuer-le-code-vba-sur-une-nouvelle-ligne)
3. [_[EXCEL-VBA]_ Déterminer et identifier l'index de couleur d'arrière-plan des cellules : obtenir les codes HEX, RGB et DEC des couleurs par les fonctions VBA concaténées](#_excel-vba_-déterminer-et-identifier-lindex-de-couleur-darrière-plan-des-cellules--obtenir-les-codes-hex-rgb-et-dec-des-couleurs-par-les-fonctions-vba-concaténées)
4. [_[EXCEL-VBA]_ Identifier les commentaires en couvrant les indicateurs de commentaire (petits triangles rouges en haut à droite dans la cellule) par des triangles bleus plus gros (formes)](#_excel-vba_-identifier-les-commentaires-en-couvrant-les-indicateurs-de-commentaire-petits-triangles-rouges-en-haut-à-droite-dans-la-cellule-par-des-triangles-bleus-plus-gros-formes)
    1. [Usage](#usage)
    2. [Création](#création)

<!-- /code_chunk_output -->



# Excel & VBA Tips {ignore=true}

!!! info

    ___[EXCEL]___ : astuce Excel
    ___[VBA]___ : astuce VBA
    ___[EXCEL-VBA]___ : astuce Excel et VBA

## _[VBA]_ Afficher un retour à la ligne dans la cellule à partir du code VBA

Grâce au caratère **Chr(10)** :

```
    texteAAfficherDansLaCellule = "Bonjour" & Chr(10) & "et" & Chr(10) & "Bienvenue"
```

Résultat dans la cellule :

```
    Bonjour
    et
    Bienvenue
```

_Source : [Retour à la ligne dans une cellule](https://www.developpez.net/forums/d491345/logiciels/microsoft-office/excel/macros-vba-excel/retour-ligne-cellule/)_

## _[VBA]_ Continuer le code VBA sur une nouvelle ligne


Grâce à **" _"** en fin de ligne (juste "Espace Underscore")

Une ligne de code sur **une** ligne :

```
    texteAAfficherDansLaCellule = "Bonjour" & Chr(10) & "et" & Chr(10) & "Bienvenue"
```

Une ligne de code sur **plusieurs** lignes :

```
    texteAAfficherDansLaCellule = "Bonjour" _
                                  & Chr(10) _
                                  & "et" _
                                  & Chr(10) _
                                  & "Bienvenue"
```

Résultat dans la cellule :

```
    Bonjour
    et
    Bienvenue
```

_Source : [Continuer le code VBA sur une nouvelle ligne](https://excel-malin.com/vba-astuces/continuer-code-vba-sur-nouvelle-ligne/)_

## _[EXCEL-VBA]_ Déterminer et identifier l'index de couleur d'arrière-plan des cellules : obtenir les codes HEX, RGB et DEC des couleurs par les fonctions VBA concaténées 

Dans la cellule Excel qui fait référence à l'autre cellule colorée (Cellule) :
```
=CONCATENER(getHexRgbColorsCodes(Cellule);getDecColorsCodes(Cellule))
```

Dans l'**éditeur VBA**, créer un **Module1** dans les **Modules** du projet VBA (**VBAProject**), ajouter **les 2 fonctions suivantes** et enregistrer :
* **getHexRgbColorsCodes**
* **getDecColorsCodes**

```
Function getHexRgbColorsCodes(FCell As Range) As String
    
    'Code HEX
    Dim hexColor As String
    hexColor = CStr(FCell.Interior.Color)
    hexColor = Right("000000" & Hex(hexColor), 6)

    'Code RGB
    Dim rgbColor As Long
    Dim R As Long, G As Long, B As Long
    rgbColor = FCell.Interior.Color
    R = rgbColor Mod 256
    G = (rgbColor \ 256) Mod 256
    B = (rgbColor \ 65536) Mod 256

    ' _ (Espace Underscore) : pour continuer le code VBA sur une nouvelle ligne
    'Chr(10): pour revenir à la ligne dans une cellule

    getHexRgbColorsCodes = "HEX " & Right(hexColor, 2) & Mid(hexColor, 3, 2) & Left(hexColor, 2) _
                            & Chr(10) _
                            & "RGB " & R & " " & G & " " & B _
                            & Chr(10) _

End Function
```

```
Function getDecColorsCodes(FCell As Range, Optional Opt As Integer = 0) As String

    'Code DEC
    Dim decColor As Long
    Dim R As Long, G As Long, B As Long
    decColor = FCell.Interior.Color
    R = decColor Mod 256
    G = (decColor \ 256) Mod 256
    B = (decColor \ 65536) Mod 256
    Select Case Opt
        Case 1
            getDecColorsCodes = R
        Case 2
            getDecColorsCodes = G
        Case 3
            getDecColorsCodes = B
        Case Else
            getDecColorsCodes = "DEC " & decColor
    End Select
End Function
```

![Image_0001](images/0001.png)
![Image_0002](images/0002.png)
![Image_0003](images/0003.png)

_Source : [Déterminer et identifier l'index de couleur d'arrière-plan des cellules dans Excel](https://fr.extendoffice.com/documents/excel/4546-excel-determine-color-of-cell.html)_

_Fichiers : [getColorsCodesFromACell](/getColorsCodesFromACell/)_

## _[EXCEL-VBA]_ Identifier les commentaires en couvrant les indicateurs de commentaire (petits triangles rouges en haut à droite dans la cellule) par des triangles bleus plus gros (formes)


### Usage

1. Visualiser les indicateurs de commentaires

![Image_0004](images/0004.png)

2. Lire les commentaires

![Image_0005](images/0005.png)
![Image_0006](images/0006.png)
![Image_0007](images/0007.png)

3. Appeler les macros : <kbd>ALT</kbd> + <kbd>F8</kbd> et double-cliquer sur `AddBigBlueTriangleOnCommentIndicator`

![Image_0008](images/0008.png)

4. Résutat :

![Image_0009](images/0009.png)

5. Appeler les macros : <kbd>ALT</kbd> + <kbd>F8</kbd> et double-cliquer sur `RemoveBigBlueTriangleOnCommentIndicator`

![Image_0010](images/0010.png)

4. Résutat :

![Image_0011](images/0011.png)

### Création

Dans l'**éditeur VBA**, créer un **Module1** dans les **Modules** du projet VBA (**VBAProject**), ajouter **les 2 procédures (Sub)** suivantes et enregistrer :
* **AddBigBlueTriangleOnCommentIndicator**
* **AddBigBlueTriangleOnCommentIndicator**

```
Sub AddBigBlueTriangleOnCommentIndicator()

    Dim pWs As Worksheet
    Dim pComment As Comment
    Dim pRng As Range
    Dim pShape As Shape
    Set pWs = Application.ActiveSheet
    wShp = 20
    hShp = 10
    For Each pComment In pWs.Comments
        Set pRng = pComment.Parent
        Set pShape = pWs.Shapes.AddShape(msoShapeRightTriangle, pRng.Offset(0, 1).Left - wShp, pRng.Top, wShp, hShp)
        With pShape
            .Flip msoFlipVertical
            .Flip msoFlipHorizontal
            .Fill.ForeColor.SchemeColor = 12
            .Fill.Visible = msoTrue
            .Fill.Solid
            .Line.Visible = msoFalse
        End With
    Next
    
End Sub
```

```
Sub RemoveBigBlueTriangleOnCommentIndicator()

    Dim pWs As Worksheet
    Dim pShape As Shape
    Set pWs = Application.ActiveSheet
    For Each pShape In pWs.Shapes
        If Not pShape.TopLeftCell.Comment Is Nothing Then
            If pShape.AutoShapeType = msoShapeRightTriangle Then
                pShape.Delete
            End If
        End If
    Next
    
End Sub
```

_Source : [Déterminer et identifier l'index de couleur d'arrière-plan des cellules dans Excel](https://fr.extendoffice.com/documents/excel/4546-excel-determine-color-of-cell.html)_

_Fichiers : [getColorsCodesFromACell](/identifyCommentIndicators/)_


