Attribute VB_Name = "Vba_Biblioteka_Mano_Git"
Sub VbaManoGitBiblioteka()

' MESSAGE, KONSOL�
'    MsgBox ("''MANO VBA BIBLIOTEKA GitHub'e''")
    Debug.Print "SPAUSDINIMAS � KONSOL�"

' KINTAM�J� DEKLARAVIMAS
    Dim xtx As String '(Integer, Diuble, Boolean, Variant)
    txt = "ne o�ys"

' IF
    If txt = "o�ys" Then

' FOR
    For i = 5 To 8
        Debug.Print i
    Next


' KLAID� GAUDYMAS
        ' 1
        On Error Resume Next

        ' 2
        On Error GoTo EX
EX:     Exit Sub

    End If

' GO TO NAUDOJIMAS
    GoTo EH
EH:

    ' KLAS�S
    ' Class Modules: "Komponentas"
    
    ' Klas�s kintamieji:
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

End Sub ' Private Sub VbaManoGitBiblioteka()
