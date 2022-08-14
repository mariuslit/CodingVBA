Attribute VB_Name = "VBA_Sintakse_GitHub"
'
'
'
Sub vba1_dim_if_for()

MsgBox ("Mano VBA biblioteka GitHub'e")
Debug.Print Chr(10) & "Spausdinimas i konsole"


Dim txt As String ' Integer, Diuble, Boolean, Variant, Object, Currency, Date
    txt = "Bezdalius"


If txt = "Ne Bezdalius" Then GoTo exit_if
    ' End If - nebutinas jei atliekama tik vienas veiksmas
    ' if pilna sintakse:
    'If n <= 1 Then
    '    f2 = 1
    'ElseIf n = 3 Then
    '    f2 = 3
    'Else: ' kam reikalingas : ???
    '    f2 = n * f2(n - 1)
    'End If
exit_if:


For i = 5 To 7 Step 2
    Debug.Print (i)
    Next
    ' Step nebutinas
    ' For Each ziureti -> vba7_Gudrybes
exit_for:

End Sub ' Sub VbaManoGitBiblioteka()
'
'
'
Sub vba2_GoTo_ErrorHanding_Beep()

    GoTo ex1
    '''
ex1:
    
    
    On Error GoTo ex2 ' /Resume Next /Exit Sub
    '''
ex2:
    
    Beep ' sisteminis pyptelejimas

End Sub ' Sub VbaManoGitBiblioteka()
'

'
Sub vba3_Array()

' deklaravimas
    Dim arr1(3) As Integer ' 4 nariu masyvas index: [0,1,2,3]
    Dim arr2(1 To 100) As Integer ' 100 nariu masyvas

    arrLenght = (UBound(arr1) - LBound(arr1) + 1) ' masyvo dydis
    Debug.Print "arrLenght: " & arrLenght

' kaip deklaruoti su kintamu dydziu (konstanta)
    Dim arr3() As Integer
    Dim kint As Integer
    kint = 100
    ReDim arr3(1 To kint) As Integer

' iteracijos
    Dim arr5 As Variant
        arr5 = Array("pirmas arr5 elementas", 4, 1, 1, " kaþkas", 1, 1, 13, 1, 2, "paskutinis arr5 elementas")
    
    For Each Item In arr5
        Debug.Print ("For Each Item In arr5: " & Item)
    Next Item
    
    For i = 0 To (UBound(arr5) - LBound(arr5))
        Debug.Print "For i=0 To (UBound(arr5) - LBound(arr5)), kai i=" & i & " " & arr5(i)
    Next i

End Sub
'
'
'
Sub vba3_Array_2D()

' 2D masyvai
    Dim arr2d(1 To 3, 1 To 2) As Integer ' 2D array 3x2
    Dim arr2d2 As Variant
        arr2d2 = [{"A","B";"1","2";"++","--"}] '2D array example 3x2

    arr2dLenght = (UBound(arr2d, 1) - LBound(arr2d, 1) + 1) ' masyvo dydis
    Debug.Print "arr2dLenght: " & arr2dLenght

    For Each Item In arr2d2
        Debug.Print ("For Each Item In arr2d: " & Item)
    Next Item

End Sub
'
'
'
Sub vba4_Math()

    ' Rnd - atsitiktiniai skaiciai nuo 0 iki 1 intervale
    ' Rasti atsitiktini skaiciu nuo 0 iki 100
    n = Round(Rnd() * 100, 0)
    Debug.Print ("Rnd - atsitiktiniai skaiciai nuo 0 iki 1 intervale " & n)


End Sub
'
'
'
Sub vba4_STRING()

oel = Chr(10)
eol2 = vbNewLine
Debug.Print ("Mano VBA bibl" & Eol & "ioteka GitHub'e")


End Sub
'
'
'
Sub vba5_Excel_Selection()

    ' KOORDINATES
    'Range ("A1")       x,y
    '[A1]               x,y
    'Cells (3,2)        y,x         naudojans Cells butina naudoti Worksheets
    'Offset(1,5).Select y,x

    ' Select, Activate
    Range("B3").Select
    Worksheets("Sheet1").Cells(3, 2).Select
    Range("A1:C2,E1,E5,E20,B2").Select
    Range("D4").Activate

    Selection.End(xlToRight).Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select

    Rows(3).Select
    Columns(3).Select
    
    ' Special Cells
    Selection.SpecialCells(xlCellTypeVisible).Select
    Range("A:A").SpecialCells(xlCellTypeBlanks).Select


    ' Workbook, Sheet, Cell
    Worksheets("Sheet1").Cells(ActiveCell.Row, 2).Select
    Worksheets("Sheet1").Cells(ActiveCell.Row, 2).Activate
    Workbooks("Book1.xls").Worksheets("Sheet1").Activate

    ActiveCell.CurrentRegion.Select
    Range("A1").End(xlDown).Offset(1, 0).Select
    Worksheets("Sheet1").Cells(10, 2).Activate
    
    ' Table selection
    Range("Table1[#All]").Select
    Range("Table1").Select
    Range("Table1[[#Headers],[Column1]]").Select

End Sub
'
'
'
Sub vba5_Excel_Formating()

    Columns("C:C").ColumnWidth = 2

    ' Celes langelio formatavimas
    ActiveCell.Select
    ActiveCell.Interior.ColorIndex = 5
    ActiveCell.Font.Color = vbWhite

    ActiveWindow.Zoom = 100
    ActiveWindow.View = xlNormalView
    ActiveWindow.LargeScroll Down:=-5

End Sub
'
'
'
Sub vba5_Excel_CRUD_Copy_Read_Update_Delete()

    ' CRUD
    GoTo mmm

    ' Cpoy
    Selection.Copy
    Selection.Offset(0, 1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    ' Read
    ' eiluciu / stulpeliu skaicius
    x = Rows.Count
    xx = ActiveCell.Row
    xxx = ActiveCell.Column
    xxxx = Selection.Rows.Count
    xxxxx = Worksheets("Sheet1").Cells(1, 1)
    fileNameShort = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)

    ' Update
    Workbooks("Book1").Worksheets("Sheet1").Range("B3").Value = 12
    ThisWorkbook.Sheets("Book1").Range("A1").Value = 55
    Range("B:C").EntireColumn.AutoFit
    Range("C2").Value = "penki"
    Range("A2", "B4").Value = 998
    ' iraso reiksme (Data+Laikas) i paskutine eilute stuplelyje skaiciuojans nuo A1
    Range("A1").Offset(Range("A1", Range("A1").End(xlDown)).Rows.Count, 0).Value = Now

    ' Delete
    Selection.EntireRow.Delete
    Selection.EntireColumn.Delete
    Selection.ClearContents
    Range("B1:C10").ClearContents

    ' Find
    ' Ctrl+F "9985"
    Worksheets("Sheet1").Cells.Find(What:="9985", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate

mmm:
               
End Sub
'
'
'
Sub vba5_Excel_Workbook_Files()

    Dim file As Object
    Dim fldr As Object
    Dim mfileName, mfileName2 As String
    
    mfileName = "C:\CodingVBA\test.xlsx"
    mfileName2 = "C:\CodingVBA\test2.xlsx"

    Workbooks.Open ("C:\CodingVBA\test2.xlsx")
    Workbooks.Open fileName:="C:\CodingVBA\test.xlsx", ReadOnly:=True
    Workbooks.Open fileName:="W:\GAMYBOS PLANAS\GAMYBOS PLANAS.xlsx", UpdateLinks:=0, ReadOnly:=True
    Workbooks.Open fileName:="Y:\CodingVBA\test.xlsm", Password:="xxx", ReadOnly:=False

    ' pavercia celiu irasus hyperlinkais, reikia nurodyti tikslu adresa
    Windows("test.xlsx").Activate
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=mfileName2, TextToDisplay:=mfileName2

    ' atidaro nurodyto langelio hyperlinkà (aktyvuoja hyperlinkà)
    ActiveCell.Hyperlinks.Item(1).Follow

    ' SHEET
    Worksheets.Add

    ActiveWorkbook.Save

    ActiveWindow.Close ' uzdaro darbaknyge
    ' uzdaro darbaknyge neissaugodamas
    ActiveWindow.Close False

End Sub
'
'
'
Sub vba5_Excel_Print()

' set printer
ActiveSheet.PageSetup.PrintArea = "APSK_PLOKSTE[[Data_]:[Pastabos_]]"

' spausdinti lapà
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
End Sub
'
'
'
Sub vba7_vba_gudrybes_LoopThroughAllWorksheets()
 
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        'list the names
        Debug.Print ws.Name
    Next ws

End Sub
'
'
'
Sub vba7_vba_gudrybes_LoopThroughFilesInFolder()

    Dim FSO As Object
    Dim file As Object
    Dim fldr As Object
    Set FSO = CreateObject("scripting.FileSystemObject")
    Set fldr = FSO.getfolder("C:\Users\User\Desktop")
    For Each file In fldr.Files
    '    If Right(File.Name, 4) = ".mp3" Then
        'add to listbox
        Debug.Print file.Name
        '???Me.ListBox1.AddItem File.Name
        'End If
    Next file

End Sub
'
'
'
Sub vba7_vba_gudrybes_FolderValuesToArr2()

    Dim selectedCells As Range
    Dim arr() As Range
    Dim i As Integer

    Set selectedCells = Selection.SpecialCells(xlCellTypeVisible)
    ReDim arr(1 To selectedCells.Count) As Range

    
    ' sukelti pazymetus duomenis i
    i = 1
    For Each Item In selectedCells
        Set arr(i) = Item
        i = i + 1
    Next Item

    ' spausdinti masyva
    For Each Item In arr
        Debug.Print Item.Value
    Next Item
    
End Sub
'
'
'
Sub vba8_vba_gudrybes_BackEnd()

' Greitas kodo vykdymas
    ' At the Beginning of your code = False
    ' At the End of your code = True
    
    ' Disable/Enable ScreenUpdating
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    
    ' Disable/Enable Calculation
    Application.Calculation = False
    Application.Calculation = True
    ' kuo skiriasi?
    Application.Calculation = xlManual
    Application.Calculation = xlAutomatic
    
    ' Disable/Enable Events
    Application.EnableEvents = False
    Application.EnableEvents = True
    

' Pauze 0:00:02 = 2 sec
    Application.Wait (Now + TimeValue("0:00:02"))
        ' Sleep Butina deklaracija pries funkcija
        '     Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
        '     Sleep (1000)                                 ' 1000 = 1 sekunde

End Sub
'
'
' VBA klases randasi modulyje "Class Modules"
Sub vba9_Class()

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

    ' OBJEKTAS
    Dim o As komponentas
    Set o = New komponentas
    
    ' Set
    o.gamintojas = "Phoenix"
    
    ' Get
    MsgBox (o.gamintojas)

End Sub
'
'
' rekursine funkcija faktorialui skaiciuoti
Function vba7_vba_gudrybes_Funkcija_Rekursija(n As Single) As Single
    If n <= 1 Then
        f2 = 1
    Else:
        f2 = n * f2(n - 1)
    End If
End Function
'
'
' Tekstinio Failo sukurimas, irasas ir istrynimas
Public FSO As New FileSystemObject
'
'
'
Sub vba99_ManoMetodai_CreateFile()

    Dim txtstr As TextStream
    Dim fileName As String
    Dim FileContent  As String
    Dim file As Object
    
    'File to be created
    fileName = "C:\CodingVBA\File.txt"
    
    'Creating a file and writing content to it
    FileContent = InputBox("Enter the File Content")
    If Len(FileContent) > 0 Then
        Set txtstr = FSO.CreateTextFile(fileName, True, True)
        txtstr.Write FileContent
        txtstr.Close
    End If
    
    ' Reading from the file that we have just created
    If FSO.FileExists(fileName) Then
        Set file = FSO.GetFile(fileName)
        Set txtstr = file.OpenAsTextStream(ForReading, TristateUseDefault)
        MsgBox txtstr.ReadAll
        txtstr.Close
      
        ' Finally Deleting the File
        file.Delete (True)
    End If
    
End Sub
'
'
' [button]
Sub vba99_ManoMetodai_FindByClipboarValue()

    Dim dataObj As Object
    Dim rr As String
    Dim txt As String
    
    rr = ActiveCell.Address
    ' Error handler
    On Error GoTo Skip
'    [a1] = rr
    
    ' Set up the data object
    Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    ' Get the data from the clipboard
    dataObj.GetFromClipboard
    
    ' Get the clipboard contents
    txt = dataObj.GetText(1)
    
    Cells.Find(What:="*" & txt & "*", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate

Skip:

    If ActiveCell.Address = rr Then
        Beep
        MsgBox ("NERASTA")
    End If
    
    ActiveWindow.SmallScroll Down:=-100

End Sub
'
'
'


