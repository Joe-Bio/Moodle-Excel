Option Explicit
Const Text_y As Integer = 5
Const Start_x As Integer = 6
Const start_MCS_x As Integer = 13
Const Question_type As String = "<!--xx MR xx -->"
Const Anz_Link_Fields As Integer = 2          ' Linkfelder für die Zufallsgenerierung im Tabellenblatt Rnd_Matrix
Sub test()
Dim teststr As String
    teststr = export_plain_html(2)
End Sub
Function export_plain_html(ByVal Ziehen As Integer) As String
' diese Funktion ist für einen externen Aufruf zur Erstellung von HTML Klausuren
' ziehen: 0: nicht ziehen; 1: ziehen, nicht zurücksetzen; 2: ziehen und zurücksetzen
'
Dim gezogen, gezogen0, gez_prozent As Variant
Dim i As Integer
Dim x As Integer
Dim bonus_comment As Boolean

    If Ziehen > 0 Then
        Call init(Ziehen, gezogen, gezogen0, gez_prozent, bonus_comment)
        Call select_questions(gezogen, gezogen0, gez_prozent)
    End If
    Call Prozente_Anpassen(gez_prozent)
    export_plain_html = Question_type + Code_Generieren(gezogen, gez_prozent) + extra_Generieren(gezogen, True)
End Function
Private Sub select_questions(ByRef gezogen As Variant, ByVal gezogen0 As Variant, ByRef prozent As Variant)
Dim Category As Integer
Dim Anz_category As Variant
Dim i As Variant
Dim x As Variant
Dim alles_falsch As Integer
Dim invalid_count As Integer
Dim valid As Boolean
    valid = False
    gezogen = gezogen0
    Anz_category = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    alles_falsch = 1
    invalid_count = 0
    While Not (valid)
        For i = 1 To 9
            Anz_category(i) = Worksheets("Gen_output").Cells(19, i + 1).Value
        Next i
        x = Start_x
        While Worksheets("questions").Cells(x, Text_y).Value <> ""
            Worksheets("questions").Cells(x, 1).Value = ""
            x = x + 1
        Wend
        For Category = 1 To 5
            While Anz_category(Category) > 0
                Call Draw_Text(Category)
                Anz_category(Category) = Anz_category(Category) - 1
            Wend
        Next Category
        valid = Validity_Check()
        If Not (valid) Then
            invalid_count = invalid_count + 1
            If invalid_count > 20 Then
                MsgBox ("Too many wrong questions, question generation not possible. Try to increase number of correct questions or include last always - none of others is correct.")
                Exit Sub
            End If
        End If
    Wend
    x = Start_x
    i = 0
    While Worksheets("questions").Cells(x, Text_y).Value <> ""
        If Worksheets("questions").Cells(x, 1).Value = "x" Then
            i = i + 1
            If Worksheets("questions").Cells(x, 3).Value = Worksheets("Gen_output").Cells(25, 2).Value Then
                gezogen(0) = x
                prozent(0) = Worksheets("questions").Cells(x, 4).Value
                i = i - 1
            Else
                gezogen(i) = x
                prozent(i) = Worksheets("questions").Cells(x, 4).Value
            End If
            If prozent(i) = 1 Then
                alles_falsch = -1
            End If
            Worksheets("questions").Cells(x, 2).Value = Worksheets("questions").Cells(x, 2).Value + 1
        End If
        x = x + 1
    Wend
    If Worksheets("Gen_output").Cells(23, 2).Value = 1 Then
        i = i + 1
        gezogen(i) = x
        prozent(i) = alles_falsch
    End If
End Sub

Private Sub Draw_Text(ByVal Category As Integer)
Dim x As Integer
Dim i As Integer
Dim selected As Boolean
Dim Min As Integer
Dim numb_count As Integer
    x = Start_x
    Call Anzahl_cat(numb_count, Min, Category)
    i = Int(Rnd * numb_count)
    selected = False
    While Not (selected)
        If (Worksheets("questions").Cells(x, 1).Value <> "x") And (Worksheets("questions").Cells(x, 2).Value = Min) And (Worksheets("questions").Cells(x, 3).Value = Category) Then
            If i = 0 Then
                selected = True
                Worksheets("questions").Cells(x, 1).Value = "x"
            Else
                i = i - 1
            End If
        End If
        x = x + 1
    Wend
End Sub
Private Sub Anzahl_cat(ByRef Anzahl As Integer, ByRef Min As Integer, ByVal Category As Integer)
Dim x As Integer
    Anzahl = 0
    Min = 100
    x = Start_x
    While Worksheets("questions").Cells(x, Text_y).Value <> ""
        If (Worksheets("questions").Cells(x, 3).Value = Category) And (Worksheets("questions").Cells(x, 1).Value <> "x") And (Min > Worksheets("questions").Cells(x, 2).Value) Then
            Min = Worksheets("questions").Cells(x, 2).Value
        End If
        x = x + 1
    Wend
    x = Start_x
    While Worksheets("questions").Cells(x, Text_y).Value <> ""
        If (Worksheets("questions").Cells(x, 3).Value = Category) And (Worksheets("questions").Cells(x, 1).Value <> "x") And (Min = Worksheets("questions").Cells(x, 2).Value) Then
            Anzahl = Anzahl + 1
        End If
        x = x + 1
    Wend
End Sub
Private Function Validity_Check() As Boolean
Dim x As Integer
Dim i As Variant
Dim proz As Variant
    If Worksheets("Gen_output").Cells(23, 2).Value = 1 Then
        Validity_Check = True
    Else
        x = Start_x
        i = 0
        proz = 0
        While Worksheets("questions").Cells(x, Text_y).Value <> ""
            If Worksheets("questions").Cells(x, 1).Value = "x" Then
                i = i + 1
                proz = proz + Worksheets("questions").Cells(x, 4).Value
            End If
            x = x + 1
        Wend
        Validity_Check = (i <> Abs(proz))
    End If
End Function


Private Sub Prozente_Anpassen(ByRef gez_prozent As Variant)
Dim summe_richtig, summe_falsch As Integer
Dim i, anz_fragen As Integer
Dim proz_plus, proz_minus As Variant
    anz_fragen = Worksheets("Gen_output").Cells(17, 2).Value
    For i = 1 To anz_fragen
        If gez_prozent(i) > 0 Then
            summe_richtig = summe_richtig + 1
        Else
            summe_falsch = summe_falsch + 1
        End If
    Next i
    proz_plus = Round(1 / summe_richtig * 100, 0) / 100
    proz_minus = -1 * Round(1 / summe_falsch * 100, 0) / 100
    If proz_minus > (-1 * Worksheets("Gen_output").Cells(20, 2).Value) Then
        proz_minus = -1 * Worksheets("Gen_output").Cells(20, 2).Value
    End If
    For i = 1 To anz_fragen
        If gez_prozent(i) > 0 Then
            gez_prozent(i) = proz_plus
        Else
            gez_prozent(i) = proz_minus
        End If
    Next i
End Sub

Private Function Code_Generieren(ByVal gezogen As Variant, ByVal gez_prozent As Variant) As String
Dim i, Anzahl, letzte_Immer As Integer
Dim prozente As String
Dim Extra As Integer
    Code_Generieren = "<p>" + Worksheets("questions").Cells(3, 3).Value + "</p>" + vbLf + "<p>"
    Code_Generieren = Code_Generieren + "{" + Trim(Str(Worksheets("Gen_output").Cells(22, 2).Value)) + ":" + Worksheets("Gen_output").Cells(21, 2).Value + ":"
    Anzahl = Worksheets("Gen_output").Cells(17, 2).Value
    letzte_Immer = Worksheets("Gen_output").Cells(23, 2).Value
    If Worksheets("Gen_output").Cells(25, 2).Value > 0 Then
        Extra = 1
    Else
        Extra = 0
    End If
    For i = 1 To Anzahl - Extra
        prozente = "%" + Trim(Str(Round(gez_prozent(i) * 100, 0))) + "%"
        If (i <= Anzahl - letzte_Immer - Extra) Then
            Code_Generieren = Code_Generieren + prozente + " " + Worksheets("questions").Cells(gezogen(i), 5).Value
        Else
            If letzte_Immer = 1 Then
                Code_Generieren = Code_Generieren + prozente + " " + Worksheets("Gen_output").Cells(23, 4).Value
            End If
        End If
        If (Worksheets("questions").Cells(gezogen(i), Text_y + 1).Value <> "") And (Worksheets("Gen_output").Cells(12, 2).Value) Then
            Code_Generieren = Code_Generieren + "#" + Worksheets("questions").Cells(gezogen(i), 6).Value
        End If
                
        If i < Anzahl - Extra Then
            Code_Generieren = Code_Generieren + " ~"
        End If
    Next i
    
    Code_Generieren = Code_Generieren + "}</p>"
End Function
Private Function extra_Generieren(ByVal gezogen As Variant, ByVal No_Comment As Boolean)
Dim i, Anzahl As Integer
Dim Yes_No As String
Dim x As Integer
Dim comment_text As String
    If Worksheets("Gen_output").Cells(25, 2).Value > 0 Then
        extra_Generieren = vbLf + "<p>" + Worksheets("questions").Cells(4, 3).Value + vbLf + Worksheets("questions").Cells(gezogen(0), Text_y).Value
        comment_text = Worksheets("questions").Cells(gezogen(0), Text_y + 1).Value
        If comment_text <> "" Then
            comment_text = "#" + comment_text
        End If
        If No_Comment Then comment_text = ""
        If Worksheets("questions").Cells(gezogen(0), 4).Value < 0 Then
            Yes_No = "Yes" + comment_text + " ~=No" + comment_text + "}"
        Else
            Yes_No = "=Yes" + comment_text + " ~No" + comment_text + "}"
        End If
        extra_Generieren = extra_Generieren + "</p><p>Do you agree? {1:MC:" + Yes_No + "</p>"
    Else
        extra_Generieren = ""
    End If
End Function


Private Sub init(ByVal Ziehen_Typ As Integer, ByRef gezogen As Variant, ByRef gezogen0 As Variant, ByRef gez_prozent As Variant, ByRef No_Bonus_comment As Boolean)
Dim x As Integer
    x = Start_x
    If Ziehen_Typ = 2 Then
        Randomize
        While Worksheets("questions").Cells(x, Text_y).Value <> ""
            Worksheets("questions").Cells(x, 2).Value = 0
            Worksheets("questions").Cells(x, 1).Value = ""
            x = x + 1
        Wend
    End If
    
    gezogen = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    gez_prozent = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    gezogen0 = gezogen
    
    If Worksheets("Gen_output").Cells(25, 2).Value > 0 Then
        No_Bonus_comment = True
    Else
        No_Bonus_comment = False
    End If
    
End Sub

Sub Export_Moodle_XLM()
Dim Filepath As String
Dim Text_File As Integer
Dim x As Integer
Dim Question_String As String
Dim Single_Question As String
Dim max_substitute As Integer
Dim substitute As Integer
Dim gezogen, gezogen0, gez_prozent As Variant
Dim i As Integer
Dim No_Bonus_comment As Boolean

    Randomize
    Call init(2, gezogen, gezogen0, gez_prozent, No_Bonus_comment)
    With Worksheets("Gen_output")
        max_substitute = .Cells(14, 2).Value
        Filepath = .Cells(4, 2).Value + .Cells(5, 2).Value
    End With
    Text_File = FreeFile
    Open Filepath For Output As Text_File
    Call Print_Replace(Text_File, max_substitute, vbLf + "<quiz>" + vbLf)
    For i = 1 To Worksheets("Gen_output").Cells(16, 2).Value
        substitute = Int(Rnd() * (max_substitute)) + 1
        Call select_questions(gezogen, gezogen0, gez_prozent)
        Call Prozente_Anpassen(gez_prozent)
        Single_Question = Code_Generieren(gezogen, gez_prozent) + extra_Generieren(gezogen, No_Bonus_comment)
        Question_String = XML_Header(i) + Single_Question + XML_End(i)
        Call Print_Replace(Text_File, substitute, Question_String)
    Next i
    Call Print_Replace(Text_File, substitute, vbLf + "</quiz>")
    Close Text_File
End Sub
Private Function XML_Header(number As Integer) As String
    XML_Header = vbLf + "<question type=""cloze"">" + vbLf + "<name><text>" + _
        Worksheets("Gen_output").Cells(6, 2).Value + " - " + Trim(Str(number)) + "</text></name>" + vbLf + "<questiontext format=""html"">" + vbLf + "<text><![CDATA[" + vbLf
End Function
Private Function XML_End(number As Integer) As String
Dim XML_End_Text As String
Dim n_tags As Integer
Dim int_tag_text As String
Dim x As Integer
    XML_End_Text = vbLf + "]]></text></questiontext>" + vbLf + "<IDNumber>" + Trim(Str(number)) + "</IDNumber>"
    
    For x = 8 To 10
        int_tag_text = ""       ' tags lassen wir hier erst mal weg...
        XML_End_Text = XML_End_Text + Gen_XML(x, int_tag_text)
    Next x
    XML_End = XML_End_Text + vbLf + "</question>"
End Function
Private Function Get_Tags(ByVal x As Integer, ByRef n_tags As Integer) As String
Dim i As Integer
Dim tag_str As String
    i = 1
    tag_str = ""
    n_tags = 0
    With Worksheets("Gen_output")
        While .Cells(x + i, 3).Value = "tag"
            If .Cells(x + i, 2).Value = 1 Then
                tag_str = tag_str + " <" + .Cells(x + i, 1).Value + "/> "
            End If
            n_tags = n_tags + 1
            i = i + 1
        Wend
    End With
    Get_Tags = tag_str
End Function
Private Function Gen_XML(ByVal x As Integer, ByVal int_tag_text As String) As String
Dim xml_str As String
Dim FieldN As String
    xml_str = ""
    With Worksheets("Gen_output")
        If .Cells(x, 2).Value <> "" Then
            FieldN = .Cells(x, 1).Value
            If .Cells(x, 3).Value = "html" Then
                xml_str = vbLf + "<" + FieldN + " format=""html""> <text><![CDATA[<p>"
                xml_str = xml_str + Trim_Hyp(.Cells(x, 2).Value)
                xml_str = xml_str + "</p>]]></text> " + int_tag_text + " </" + FieldN + ">"
            Else
                xml_str = vbLf + "<" + FieldN + ">" + Trim_Hyp(.Cells(x, 2).Value) + "</" + FieldN + ">"
            End If
        End If
    End With
    Gen_XML = xml_str
End Function
Private Function Trim_Hyp(ByVal Zahl_Str As String) As String
    If Left(Zahl_Str, 1) = "'" Then
        Zahl_Str = Right(Zahl_Str, Len(Zahl_Str) - 1)
    End If
    Trim_Hyp = Zahl_Str
End Function
Sub Check_Questions()
Dim x As Integer
Dim correct As Boolean
Dim response As Boolean
Dim is_string As String
Dim resp As Integer
    x = Start_x
    With Worksheets("questions")
        While .Cells(x, 5).Value <> ""
            correct = (.Cells(x, 4).Value = 1)
            resp = MsgBox("Is this statement correct: " + .Cells(x, 5).Value, vbYesNoCancel)
            If resp = vbCancel Then
                Exit Sub
            End If
            response = (resp = vbYes)
            If (response <> correct) Then
                is_string = "Correct"
                If Not (correct) Then
                    is_string = "Wrong"
                End If
                resp = MsgBox("Attention! The statement: " + .Cells(x, 5).Value + " is considered as " + is_string + " with the following comment: " + _
                    .Cells(x, 6).Value + ". Do you think this has to be changed? ", vbYesNoCancel, "****RESPONSES DIFFER****")
                If resp = vbCancel Then
                    Exit Sub
                End If
                If resp = vbYes Then
                    If (MsgBox("Response considered as " + is_string + " will be changed. Are you sure?", vbYesNo, "RESPONSE WILL BE CHANGED") = vbYes) Then
                        .Cells(x, 4).Value = -1 * .Cells(x, 4).Value
                    End If
                End If
            End If
            x = x + 1
        Wend
    End With
    
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
            print_text = replace(print_text, "##" + platzhalter + ";", antwort, 1, 1)
            start_Zeichen = InStr(start_Zeichen, print_text, "##")
        Wend
    End If
    Call Clean_HTML(print_text)
    Print #Text_File, print_text
End Sub

Private Function get_ant_nr(ByVal suchstring As String) As Integer
Dim i As Integer
    With Worksheets("Rnd_Matrix")
        i = 1
        While .Cells(1 + i, 1).Value <> ""
            If .Cells(1 + i, 1).Value = suchstring Then
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
        To_Be_Cleaned = replace(To_Be_Cleaned, non_html(i), is_html(i))
    Next i
End Sub

