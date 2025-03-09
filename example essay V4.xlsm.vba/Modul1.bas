Option Explicit
Const start_x As Integer = 9
Const Text_y As Integer = 5
Const Spalten As Integer = 7   ' note: Spalten = y Ende...
Const XML_Out_X As Integer = 17
Const wrong As String = "xxxxxxxxxxxxxxxxxxxxxx"
Const Question_type As String = "<!-- xx Essay xx -->"
Const New_Question As String = "<!-- xx NewQuestion xx -->"
Const Anz_Link_Fields As Integer = 2          ' Linkfelder für die Zufallsgenerierung im Tabellenblatt Rnd_Matrix
Sub test()
Dim teststr As String
    teststr = export_plain_html("", "", "", 0)
    Debug.Print (teststr)
End Sub
Function export_plain_html(ByVal q_text As String, ByVal feed_text As String, ByVal grader_text As String, ByVal number As Integer) As String
' diese Funktion ist für einen externen Aufruf zur Erstellung von HTML Klausuren
' Funktion export_plain kann Fragentext übergeben werden
' Reihenfolge: Frage#Feedback#GraderInfo#Punkte#Zeilenlänge
Dim question_text As String
    Call Get_Question_Text(q_text, feed_text, grader_text)
    question_text = q_text + "#" + feed_text + "#" + Trim(Str(Worksheets("Questions").Cells(8, 2).Value)) + "#" + Trim(Str(Worksheets("Gen_output").Cells(12, 2).Value))
    export_plain_html = Question_type + vbLf + question_text
End Function
Private Sub Get_Question_Text(ByRef q_text As String, ByRef feed_text As String, ByRef grader_text As String)
    With Worksheets("Questions")
        If q_text = "" Then
            q_text = .Cells(2, 2).Value
        End If
        If feed_text = "" Then
            feed_text = .Cells(4, 2).Value
        End If
        If grader_text = "" Then
            grader_text = .Cells(6, 2).Value
        End If
    End With
End Sub
Private Function Essay_Frage_generieren(ByVal q_text As String, ByVal feed_text As String, ByVal grader_text As String, ByVal number As Integer) As String
Dim question_string As String
Dim x As Integer
Dim n As String
Dim numb As String
    With Worksheets("Form_Sheet")
        x = 1
        question_string = ""
        While .Cells(x, 1).Value <> ""
            question_string = question_string + .Cells(x, 1).Value + vbLf
            x = x + 1
        Wend
    End With
    Call Get_Question_Text(q_text, feed_text, grader_text)
    With Worksheets("Questions")
        question_string = Replace(question_string, "XXXQ_TextXXX", q_text)
        question_string = Replace(question_string, "XXXQ_FeedbackXXX", feed_text)
        question_string = Replace(question_string, "XXXQ_Grader_InfoXXX", grader_text)
        question_string = Replace(question_string, "XXXQ_GradeXXX", Trim(Str(.Cells(8, 2).Value)))
    End With
    With Worksheets("Gen_output")
        numb = ""
        If number > 0 Then
            numb = " - " + Trim(Str(number))
        End If
        question_string = Replace(question_string, "XXXQ_NameXXX", .Cells(4, 2).Value + numb)
        question_string = Replace(question_string, "XXXQ_LinesXXX", Trim(Str(.Cells(12, 2).Value)))
        If .Cells(13, 2).Value = 0 Then
            n = ""
        Else
            n = Trim(Str(.Cells(13, 2).Value))
        End If
        question_string = Replace(question_string, "XXXQ_MinWXXX", n)
        If .Cells(14, 2).Value = 0 Then
            n = ""
        Else
            n = Trim(Str(.Cells(14, 2).Value))
        End If
        question_string = Replace(question_string, "XXXQ_MaxWXXX", n)
    End With
    Essay_Frage_generieren = question_string
End Function


Sub Generate_Moodle_XML()
Dim Filepath As String
Dim Text_File As Integer
Dim x As Integer
Dim number As Integer
Dim Question_HTML As String
Dim question_string As String
Dim max_substitute As Integer
Dim substitute As Integer
    Randomize
    With Worksheets("Gen_output")
        Filepath = .Cells(5, 2) + .Cells(6, 2)
        max_substitute = .Cells(10, 2).Value
    End With
    Text_File = FreeFile
    Open Filepath For Output As Text_File
    Print #Text_File, vbLf + "<quiz>" + vbLf
    For number = 1 To Worksheets("Gen_output").Cells(8, 2).Value
        question_string = Essay_Frage_generieren("", "", "", number)
        substitute = Int(Rnd() * (max_substitute)) + 1
        Call Print_Replace(Text_File, substitute, question_string)
    Next number
    Print #Text_File, vbLf + "</quiz>"
    Close Text_File
    
End Sub

Private Sub Print_Replace(ByVal Text_File As Integer, ByVal Repl As Integer, ByVal print_text As String)
Dim start_Zeichen As Long
Dim laenge As Integer
Dim ant_nr As Integer
Dim antwort As String
Dim platzhalter As String
Dim Alt_Name As Boolean
Dim ar_name As Variant
' Funktion schreibt in Datei, falls replace aktiviert werden Imagelinks / Beschriftungen angepasst
    If Repl > 0 Then
        start_Zeichen = 1
        start_Zeichen = InStr(start_Zeichen, print_text, "##")
        While start_Zeichen > 0
            laenge = InStr(start_Zeichen, print_text, ";") - start_Zeichen - 2
            platzhalter = Right(Left(print_text, start_Zeichen + 1 + laenge), laenge)
            Alt_Name = (platzhalter = "DatName")
            If IsNumeric(platzhalter) Then
                ant_nr = CInt(platzhalter) + Anz_Link_Fields
            Else
                If Alt_Name Then
                    platzhalter = "ImgLink"
                End If
                ant_nr = get_ant_nr(platzhalter)
                If ant_nr = 0 Then
                    MsgBox ("Replace Term " + platzhalter + " not found! Replacement aborted!")
                    Exit Sub
                End If
            End If
            antwort = Worksheets("Rnd_Matrix").Cells(1 + ant_nr, 1 + Repl).Value
            If Alt_Name Then
                ar_name = Split(antwort, "/")
                antwort = ar_name(UBound(ar_name))
                platzhalter = "DatName"
            End If
            print_text = Replace(print_text, "##" + platzhalter + ";", antwort, 1, 1)
            start_Zeichen = InStr(start_Zeichen, print_text, "##")
        Wend
    End If
    Call Clean_HTML(print_text)
    Print #Text_File, print_text
End Sub

Private Function get_ant_nr(ByVal Suchstring As String) As Integer
Dim i As Integer
    With Worksheets("Rnd_Matrix")
        i = 1
        While .Cells(1 + i, 1).Value <> ""
            If .Cells(1 + i, 1).Value = Suchstring Then
                get_ant_nr = i
                Exit Function
            End If
            i = i + 1
        Wend
        get_ant_nr = 0
    End With
End Function

Private Sub Clean_HTML(ByRef To_Be_Cleaned As String)
' Funktion ersetzt symbole durch HTML Code - neu drin seit Version 3.1.2 - sollte in die anderen Makros noch integriert werden!
' Aufruf steht in sub print replace direkt vor dem Schreiben
Dim non_html As Variant
Dim is_html As Variant
Dim anz As Integer
Dim i As Integer
    non_html = Array("", "§", "°", "²", "³", "µ", "Ä", "Ö", "Ü", "ä", "ö", "ü", "ß")
    is_html = Array("", "&sect;", "&deg;", "&sup2;", "&sup3;", "&micro;", "&Auml;", "&Ouml;", "&Uuml;", "&auml;", "&ouml;", "&uuml;", "&szlig;")
    For i = 1 To UBound(non_html)
        To_Be_Cleaned = Replace(To_Be_Cleaned, non_html(i), is_html(i))
    Next i
End Sub
