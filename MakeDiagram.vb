Sub MakeDiagram()

i=2
Dim daymax As Date
Dim Daymin As Date
Dim stt As Date
Dim fin As Date
Dim genre As String
Dim genrerange As Range
Dim genrecells As Range
Dim genrecolor As Long

Do Until Worksheets("ToDoリスト").Cells(i,1)=""

Worksheets("チャート").Cellls(i,1)=Worksheets("ToDoリスト").Cells(i,1)

i=i+1

Loop

i=2
j=2

With Worksheets("ToDoリスト").Application.WorksheetFunction

daymax=.Max(Range("B:C"))
daymin=.Max(Range("B:C"))
i=i+1

End With

Do Until daymin=daymax

Worksheets("チャート").Cells(1,i).Value=daymin

daymin=daymin+1
i=i+1

Loop

Worksheets("チャート").Cells(1,i).Value=aymin
daymin=daymin+1
Worksheets("チャート").Cells(1,i).Value=daymin

i=2

Do Until Worksheets("ToDoリスト").Cells(i,1)=""

    j=2

    genre=Worksheets("ToDoリスト").Cells(i,4)
    Set genrerange=Worksheets("設定").Range("A:A").Find(what:=genre)
    genrecolor=Worksheets("設定").Cells(Genrerange.Row,1).Interior.Color
    Do Until Worksheets("ToDoリスト").Cells(i,3)=Worksheets("チャート").Cells(1,j)-1

        Worksheets("チャート").Cells(i,j).Interior.Color=genrecolor
        
        j=j+1
    
    Loop
j=2

Do Until Worksheets("ToDoリスト").Cells(i,2)=Worksheets("チャート").Cells(i,j)

    Worksheets("チャート").Cells(i,j).Interior.ColorIndex=0
    j=j+1

Loop

i=i+1

Loop

End Sub