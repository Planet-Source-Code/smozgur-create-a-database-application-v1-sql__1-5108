Attribute VB_Name = "DBModule"
Global DBMain As Database
Global RecSet As Recordset

Public Sub CreateNewDB()
    'Goto the Creating Database Applications for the Beginners in Planet
End Sub

Public Sub Main()
    Stop
    'Please ReadThis

'!!!WARNING!!!
'First you have to add Microsoft DAO Object Library
'using by References Dialog Box
'If not VB doesnot know the database elements
    
    'Last code of mine was about quering tables
    'And i used a little bit SQL
    'In this small code i will try to explain SQLText string
    'How is it created and working.
    
    'If you ready Press F5 to continue
    
    'Enjoy again!
    
    'This series will continue under this name
    '!Lets Create Database Applications-V?
    
    'Suat
    
    Load Form1
    Form1.Show
End Sub
Public Sub QueryData(reqText As String)
    'first you need to learn little SQL
    'SQL is Structured Query Language
    'it is like this:
    
    'SELECT some fields FROM databases WHERE somehing you choose
    
    'Of course it is not real SQL exactly
    'But you have to use this first time
    
    'I always create my database queries with SQL
    'It is more fast and easy
    'I will teach you how you use other way to querying
    'But this time SQL is the better way
    
    'Let's say the user first hit 'w' key and continue
    'But we run this code from the Change event of box
    'So this code will query the database for each key
    'This means for the first key it bring to me
    'the words begin with 'w'
    
    'Then for the second word lets say 'wo'
    'it will bring to me the words begin with 'wo'
    'And continue...
       
    'The reqText is the word which looking for
    'Then it is not important how many characters if it is
    
 
    'Dimensioning string
Dim SQLText As String
    'Create our SQL Text
    Form1.Label1 = "SELECT * : means select all fields in table and fixed for this sample."
    Form1.Label2 = "FROM DBWords : means get fields from specified table and fixed for this sample."
    Form1.Label3 = "WHERE Left(dbWord," & Len(reqText) & ")='" & reqText & "';"
    SQLText = "SELECT *"
    SQLText = SQLText + " FROM DBWords"
    SQLText = SQLText + " WHERE Left(dbWord," & Len(reqText) & ")='" & reqText & "';"
    Form1.Label4 = SQLText
    
    'And Create a Recordset object with this SQLText
    Set RecSet = DBMain.OpenRecordset(SQLText)
    'Clear the list box
    
    Form1.lstWords.Clear
    If RecSet.RecordCount = 0 Then
        Form1.lstWords.AddItem "No item was found"
        Form1.txtExp.Text = ""
        Form1.Label5 = "You typed " & Len(reqText) & ", but no words begin this letters."
        Exit Sub
    End If
    
    'Do this, i will explain why later
    RecSet.MoveLast: RecSet.MoveFirst
    'fill the list box
    Do Until RecSet.EOF
        Form1.lstWords.AddItem RecSet.Fields(0)
        RecSet.MoveNext
    Loop
    'if there is just one word you see in the list
    'then pretend like user click a list box item
    'and bring the sigle item description which able to be selected
    If Form1.lstWords.ListCount = 1 Then
        RecSet.MoveFirst
        Form1.txtExp.Text = RecSet.Fields(1)
        Form1.txtWord.Text = Form1.lstWords.List(0)
        Form1.txtWord.SelLength = Len(Form1.txtWord.Text)
    Else
    'if not dont fill the list box
    'because there are lots of words
        Form1.txtExp.Text = ""
    End If
    
    If Form1.lstWords.ListCount = 1 Then
        Form1.Label5 = "You typed enought letter for the words, then there is no another word begin this letters."
    ElseIf reqText = "" Then
        Form1.Label5 = "You didnot enter any letter to requery, then it filled the list with all words in table unless looking for their beginner letters."
    ElseIf Form1.lstWords.ListCount > 1 Then
        Form1.Label5 = "You typed " & Len(reqText) & " letter(s) then query is looking for the first " & Len(reqText) & " letters in the words and filled listbox with " & Form1.lstWords.ListCount & " words."
    End If


    'ok
    'This is the second code of mine
    'I see my code's visitor
    'But i am not really sure if you need my codes
    'If it is i will create bigger database application
    'here
    'Please tell me your idea.
    'Bye for today
End Sub

Public Sub StoreData()
    'Goto the Creating Database Applications for the Beginners in Planet
End Sub
