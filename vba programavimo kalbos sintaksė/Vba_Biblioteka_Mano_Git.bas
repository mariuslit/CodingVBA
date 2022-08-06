Attribute VB_Name = "Vba_Biblioteka_Mano_Git"
' VBA mano kodai, metodai, funkcijos, klases

' VBA metodai (Sub) ir funkcijos (Function) randasi modulyje "Modules"
Sub VBA_Sintakse()

' MESSAGE, CONSOLE
'    MsgBox ("MANO VBA BIBLIOTEKA GitHub'e")
    Debug.Print "Spausdinimas i konsole"

' KINTAMØJØ DEKLARAVIMAS
    Dim xtx As String ' Integer, Diuble, Boolean, Variant
    txt = "Bezdalius"

' FOR
    For i = 5 To 7 Step 2
        Debug.Print (i)
    Next
    ' Step nebutinas

' IF
    If txt = "Ne Bezdalius" Then

    End If
    ' End If - jei daugiau nei vien eilute

' GO TO NAUDOJIMAS
    GoTo Ex1
Ex1:
    GoTo Ex2
Ex2:

' KLAIDU GAUDYMAS
        ' 1
        On Error Resume Next

        ' 2
        On Error GoTo EX
EX:     Exit Sub

End Sub ' Sub VbaManoGitBiblioteka()

' VBA klases randasi modulyje "Class Modules"
Sub VBA_Klase()

    ' Class Modules: "Komponentas"

    ' matomumas Public jei noriu pasiekti visame projekte
    
    ' Klases kintamieji:
    'Public projektas As String
    'Public kas_atrinko As String
    'Public rowsCount As String
    'Public gamintojas As String
    'Public kodas_pavadinimas As String
    'Public aprasymas_pastabos As String
    'Public kiekis As Integer
    'Public zenklas As String
    'Public likutisPries As Integer
    'Public likutisPo As Integer
    'Public likPries_Kiekis_likPo As String
    'Public num As Integer

' OBJEKTO DEKLARAVIMAS
    Dim o As Komponentas
    Set o = New Komponentas
    
    ' Set
    o.gamintojas = "Phoenix"
    
    ' Get
    MsgBox (o.gamintojas)

End Sub

