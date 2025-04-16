Option Explicit
Const Text_y As Integer = 5
Const Start_x As Integer = 6
Const Question_type As String = "<!--xx SC xx -->"
Const Anz_Link_Fields As Integer = 2          ' Linkfelder für die Zufallsgenerierung im Tabellenblatt Rnd_Matrix
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
Private Sub select_questions(ByRef gezogen As Variant, ByVal gezogen0 As Variant, ByRef gez_wahr As Variant, _
        ByRef gez_falsch As Variant, ByVal pos_quest As Variant, ByVal category As Integer)
' im Prinzip fertig, muss noch getestet werden!!!!!
Dim Wahre_Version As Boolean
Dim extra As Boolean
Dim extra_value As Variant
Dim y_multiple As Integer
Dim remain As Integer
Dim q_anz As Integer
Dim Ende As Integer
Dim x As Integer
    gezogen = gezogen0
    x = pos_quest(category) + 1
    extra_value = Worksheets("Gen_output").Cells(20, 2).Value
    extra = (extra_value > 0)
    q_anz = Worksheets("Gen_output").Cells(17, 2).Value
    ' draw correct or wrong version
    Wahre_Version = (gez_wahr(category) >= (Rnd * (gez_wahr(category) + gez_falsch(category))))
    If Wahre_Version Then
        gez_wahr(category) = gez_wahr(category) - 1
        y_multiple = 5
    Else
        gez_falsch(category) = gez_falsch(category) - 1
        y_multiple = 6
    End If
    
    ' "korrekte" antwort ziehen
    gezogen(0) = 0
    If extra Then
        Ende = 1
    Else
        Ende = 0
    End If
    q_anz = q_anz - 1
    If (Rnd < extra_value) Then     ' extra antwort als richtig
        gezogen(1) = 1000
        gezogen(0) = 1
        extra = False
        Ende = 0
    Else                                                ' Antwort ziehen
        Call Draw_Text(gezogen, y_multiple + Round((5.5 - y_multiple) * 2, 0), category, x)
    End If
    
    While q_anz > Ende
        Call Draw_Text(gezogen, y_multiple, category, x)        ' falsche ziehen
        q_anz = q_anz - 1
    Wend
    If extra Then          ' letzte hinzufügen
        gezogen(0) = gezogen(0) + 1
        gezogen(gezogen(0)) = 1000
    End If
    If Wahre_Version Then
        gezogen(0) = gezogen(0) * (-1)
    End If
End Sub

Private Sub Draw_Text(ByRef gezogen As Variant, ByVal y_current As Integer, ByVal q_no As Integer, ByVal start_cat_x As Integer)

Dim category As Integer
Dim x As Integer
Dim selected As Boolean
Dim Min As Integer
Dim i As Integer
Dim numb_count As Integer
    x = start_cat_x
    category = Worksheets("questions").Cells(x, 3).Value
    Call Anzahl_cat(numb_count, Min, start_cat_x, y_current, category)
    i = Int(Rnd * numb_count)
    selected = False
    With Worksheets("questions")
        While Not (selected)
            If (.Cells(x, 1).Value <> "x") And (.Cells(x, 2).Value = Min) And (.Cells(x, 3).Value = category) And (.Cells(x, y_current).Value <> "") Then
                If i = 0 Then
                    selected = True
                    .Cells(x, 1).Value = "x"
                    .Cells(x, 2).Value = .Cells(x, 2).Value + 1
                    gezogen(0) = gezogen(0) + 1
                    gezogen(gezogen(0)) = x
                Else
                    i = i - 1
                End If
            End If
            x = x + 1
        Wend
    End With
End Sub
Private Sub Anzahl_cat(ByRef Anzahl As Integer, ByRef Min As Integer, ByVal start_cat_x As Integer, _
        ByVal Current_y As Integer, ByVal category As Integer)
Dim x As Integer
    Anzahl = 0
    Min = 100
    x = start_cat_x
    With Worksheets("questions")
        While (.Cells(x, 3).Value = category)
            If (.Cells(x, 1).Value <> "x") And (.Cells(x, Current_y).Value <> "") And (Min > .Cells(x, 2).Value) Then
                Min = .Cells(x, 2).Value
            End If
            x = x + 1
        Wend
        x = start_cat_x
        While (.Cells(x, 3).Value = category)
            If (.Cells(x, 1).Value <> "x") And (Min = .Cells(x, 2).Value) And (.Cells(x, Current_y).Value <> "") Then
                Anzahl = Anzahl + 1
            End If
            x = x + 1
        Wend
    End With
End Sub

Private Function Code_Generieren(ByVal gezogen As Variant, ByVal x_pos As Integer) As String
Dim i, Anzahl, letzte_Immer As Integer
Dim typ As String
Dim bold_start, bold_end As String
Dim y As Integer
Dim points As String
Dim extra As Integer
Dim include_comment As Boolean
Dim last_chosen As String
Dim comment As String
    typ = Trim(Worksheets("questions").Cells(x_pos, 4).Value)
    If typ = "" Then
        typ = "MCVS"
    End If
    include_comment = Worksheets("Gen_output").Cells(12, 2).Value
    If gezogen(0) < 0 Then
        y = 6
    Else
        y = 5
    End If
    If Worksheets("Gen_output").Cells(13, 2) Then
        bold_start = "<b>"
        bold_end = "</b>"
    Else
        bold_start = ""
        bold_end = ""
    End If
    last_chosen = Worksheets("Gen_output").Cells(20, y - 1).Value
    points = ""
    If Val(typ) = 0 Then
        points = Trim(Str(Worksheets("Gen_output").Cells(19, 2).Value)) + ":"
    End If
    Code_Generieren = "<p>" + bold_start + Worksheets("questions").Cells(x_pos, y).Value + bold_end + "</p>" + vbLf + "<p>"
    Code_Generieren = Code_Generieren + "{" + points + typ + ":"
    For i = 1 To Abs(gezogen(0))
        If i = 1 Then
            Code_Generieren = Code_Generieren + "="
        Else
            Code_Generieren = Code_Generieren + " ~"
        End If
        If gezogen(i) <> 1000 Then
            Code_Generieren = Code_Generieren + Worksheets("questions").Cells(gezogen(i), y).Value
            If include_comment And Worksheets("questions").Cells(gezogen(i), 7).Value <> "" Then
                Code_Generieren = Code_Generieren + "#" + Worksheets("questions").Cells(gezogen(i), 7).Value
            End If
        Else
            Code_Generieren = Code_Generieren + last_chosen
        End If
        If i = 1 Then
            y = y + (5.5 - y) * 2
        End If
    Next i
    
    Code_Generieren = Code_Generieren + "}</p>" + vbLf
End Function


Private Sub init(ByVal Ziehen_Typ As Integer, ByRef gez_wahr As Variant, ByRef gez_falsch As Variant, ByRef gezogen As Variant, ByRef gezogen0 As Variant, _
            ByRef stat_wahr As Variant, ByRef stat_falsch As Variant, ByRef pos_quest As Variant, ByRef valid_questions As Variant)
' Logik gez_wahr, gez_falsch:
' wird "Wahre Antwort" gezogen, so ist ein Statment richtig, und die restlichen falsch.
' damit das also erfüllbar ist muss es mindestens ein Wahres statement geben und mindestens no_q - 1 falsche statements.
' gibt es mehr falsche als wahre statements, so sollte öfters eine WAhre Antwort gezogen werden als eine falsche, um die
' Möglichkeiten optimal auszunutzen. gez_wahr ist dann also größer als gez_falsch - und das ist genau umgekehrt als die anzahl der Statements (die in
' stat_wahr und stat_falsch abgelegt sind.
' für valide Antworten muss noch beachtet werden, dass ein Statement auch falsch und richtig enthalten kann - dann erhöht sich die Anzahl der notwendigen statements entsprechend.
Dim x As Integer
Dim i As Integer
Dim val_q As Integer
Dim no_q As Integer     ' Anzahl statements pro Frage (abzüglich 1 - entspricht also der Anzahl der gliechartigen Statements, die eine Frage aufweisen muss)
Dim no_clz As Integer   ' anzahl der Cloze Fragen (mit jeweils mehreren SC Frage)
Dim q_add As Integer
Dim quest_anz As Variant
Dim last_always As Boolean
Dim zusatz As Integer
    x = Start_x
    If Ziehen_Typ = 2 Then
        Randomize
        While Worksheets("questions").Cells(x, 3).Value <> ""
            Worksheets("questions").Cells(x, 2).Value = 0
            Worksheets("questions").Cells(x, 1).Value = ""
            x = x + 1
        Wend
    End If
    
    gez_wahr = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    gez_falsch = gez_wahr
    gezogen0 = gez_wahr
    stat_wahr = gez_wahr
    stat_falsch = gez_wahr
    valid_questions = gez_wahr
    gezogen = gez_wahr
    pos_quest = gez_wahr
    quest_anz = gez_wahr
    i = 0
    x = Start_x
    q_add = 1
    With Worksheets("questions")
        While .Cells(x, 3).Value <> ""
            If .Cells(x, 3).Value > i Then
                i = i + 1
                pos_quest(i) = x
                x = x + 1
            End If
            If .Cells(x, 5).Value <> "" Then
                stat_wahr(i) = stat_wahr(i) + 1
            End If
            If .Cells(x, 6).Value <> "" Then
                stat_falsch(i) = stat_falsch(i) + 1
                If .Cells(x, 5).Value <> "" Then
                    q_add = 2
                End If
            End If
            x = x + 1
        Wend
    End With
    quest_anz(0) = i
    no_q = Worksheets("Gen_output").Cells(17, 2).Value - 1
    val_q = 0
    last_always = (Worksheets("Gen_output").Cells(20, 2).Value > 0)
    For i = 1 To quest_anz(0)           ' Voraussetzungen abklappern, dass Frage möglich ist
        If (stat_wahr(i) + stat_falsch(i)) >= no_q Then
            If Not (Not (last_always) And ((stat_wahr(i) = 0) Or (stat_falsch(i) = 0))) Then
                If (((stat_wahr(i) + q_add) >= no_q) Or ((stat_falsch(i) + q_add) >= no_q)) Then
                    If Not (Worksheets("Gen_output").Cells(21, 2).Value And ((stat_wahr(i) = 0) Or (stat_falsch(i) + q_add < no_q))) Then
                        val_q = val_q + 1
                        valid_questions(val_q) = i
                    End If
                End If
            End If
        End If
    Next i
    valid_questions(0) = val_q
    If val_q = 0 Then
        Exit Sub
    End If
    no_clz = Worksheets("Gen_output").Cells(16, 2).Value
    For i = 1 To valid_questions(0)
        If Worksheets("Gen_output").Cells(21, 2).Value Then
            gez_wahr(valid_questions(i)) = no_clz
        Else
            gez_wahr(valid_questions(i)) = Round(no_clz * (stat_falsch(valid_questions(i)) / (stat_wahr(valid_questions(i)) + stat_falsch(valid_questions(i)))), 0)
            gez_falsch(valid_questions(i)) = no_clz - gez_wahr(valid_questions(i))
        End If
        If last_always And (stat_wahr(valid_questions(i)) < (no_q - 1)) Then
            gez_falsch(valid_questions(i)) = no_clz
            gez_wahr(valid_questions(i)) = 0
        End If
        If last_always And (stat_falsch(valid_questions(i)) < (no_q - 1)) Then
            gez_wahr(valid_questions(i)) = no_clz
            gez_falsch(valid_questions(i)) = 0
        End If
        
    Next i
        
End Sub

Sub Export_Moodle_XLM()
Dim Filepath As String
Dim Text_File As Integer
Dim x As Integer
Dim Question_String As String
Dim Single_Question As String
Dim max_substitute As Integer
Dim substitute As Integer
Dim gez_wahr, gez_falsch, gezogen0, quest_anz, quest_wahr, quest_falsch, valid_quest, gezogen, pos_quest As Variant
Dim i As Integer
Dim n As Integer
Dim No_Bonus_comment As Boolean

    Randomize
    Call init(2, gez_wahr, gez_falsch, gezogen, gezogen0, quest_wahr, quest_falsch, pos_quest, valid_quest)
    If valid_quest(0) = 0 Then
        MsgBox ("No valid questions can be generated. Please check number of positive and negative responses and your requested number of statemtens per question.")
        Exit Sub
    End If
    With Worksheets("Gen_output")
        max_substitute = .Cells(14, 2).Value
        Filepath = .Cells(4, 2).Value + .Cells(5, 2).Value
    End With
    Text_File = FreeFile
    Open Filepath For Output As Text_File
    Call Print_Replace(Text_File, max_substitute, vbLf + "<quiz>" + vbLf)
    For n = 1 To Worksheets("Gen_output").Cells(16, 2).Value
        Question_String = XML_Header(n)
        If Worksheets("questions").Cells(3, 3).Value <> "" Then
            Question_String = Question_String + Worksheets("questions").Cells(3, 3).Value + "</p>" + vbLf
        End If
        
        For i = 1 To valid_quest(0)
            substitute = Int(Rnd() * (max_substitute)) + 1
            Call select_questions(gezogen, gezogen0, gez_wahr, gez_falsch, pos_quest, valid_quest(i))
            Single_Question = Code_Generieren(gezogen, pos_quest(valid_quest(i)))
            Question_String = Question_String + Single_Question
        Next i
        Question_String = Question_String + XML_End(n)
        Call Print_Replace(Text_File, substitute, Question_String)
        
        Worksheets("questions").Range("A6:A500").ClearContents

    Next n
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
Sub Transpose_MR_to_SC()
'
' Transpose_MR_to_SC Makro
' makro transposes a question in the MR format into this SC format.
'
Dim x As Integer
Dim cat As Integer
    Range("F6:F47").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("G6").Select
    ActiveSheet.Paste
    Range("A6").Select
    cat = 0
    x = 6
    With Worksheets("questions")
        While .Cells(x, 4).Value <> ""
            If .Cells(x, 3).Value > cat Then             ' neue Zeile für Überschrift einsetzen
                cat = .Cells(x, 3).Value
                .Cells(x, 3).EntireRow.Insert shift:=xlDown
                .Cells(x, 3).Value = cat
                .Cells(x, 4).Value = "MCVS"
                .Cells(x, 5).Value = "Please insert question text here"
            Else
                If .Cells(x, 4).Value < 0 Then
                    .Cells(x, 6).Value = .Cells(x, 5).Value
                    .Cells(x, 5).Value = ""
                End If
                .Cells(x, 4).Value = ""
            End If
            x = x + 1
        Wend
    End With
End Sub
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

