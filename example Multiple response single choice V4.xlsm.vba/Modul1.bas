Option Explicit
Const Text_y As Integer = 5
Const Start_x As Integer = 6
Const start_MCS_x As Integer = 13
Const Question_type As String = "<!-- xx MR-SC xx -->"
Const Anz_Link_Fields As Integer = 2          ' Linkfelder für die Zufallsgenerierung im Tabellenblatt Rnd_Matrix
Function export_plain_html(ByVal Ziehen As Integer) As String
' diese Funktion ist für einen externen Aufruf zur Erstellung von HTML Klausuren
' ziehen: 0: nicht ziehen; 1: ziehen, nicht zurücksetzen; 2: ziehen und zurücksetzen
' cave: bei dieser Fragenart multiple response single choice ist das erstellen der Frage ohne ziehen problematisch!
Dim gezogen, gezogen0, gez_prozent As Variant
Dim i As Integer
Dim x As Integer

    If Ziehen > 0 Then
        Call init(Ziehen, gezogen, gezogen0, gez_prozent)
        Call select_questions(gezogen, gezogen0, gez_prozent)
        Call Generate_Response_Is(gezogen, gez_prozent)
        Call Generate_Response_Set
    End If
    export_plain_html = Question_type + Code_Generieren(gezogen)
End Function
Sub export_moodle_xml()
Dim gezogen, gezogen0, gez_prozent As Variant
Dim i As Integer
Dim Filepath As String
Dim Text_File As Integer
Dim x As Integer
Dim Question_String As String
Dim max_substitute As Integer
Dim substitute As Integer

    Call init(2, gezogen, gezogen0, gez_prozent)
    With Worksheets("Gen_output")
        Filepath = .Cells(4, 2) + .Cells(5, 2)
        max_substitute = .Cells(29, 2).Value
    End With
    Text_File = FreeFile
    Open Filepath For Output As Text_File
    Print #Text_File, vbLf + "<quiz>" + vbLf

    For i = 1 To Worksheets("Gen_output").Cells(20, 2).Value
        Call select_questions(gezogen, gezogen0, gez_prozent)
        Call Generate_Response_Is(gezogen, gez_prozent)
        Call Generate_Response_Set
        Question_String = XML_Header(i)
        Question_String = Question_String + Code_Generieren(gezogen)
        Question_String = Question_String + XML_End(i)
        substitute = Int(Rnd() * (max_substitute)) + 1
        Call Print_Replace(Text_File, substitute, Question_String)
    Next i
    
    Print #Text_File, vbLf + "</quiz>"
    Close Text_File


End Sub
Private Sub Generate_Response_Is(ByVal gezogen As Variant, ByVal gez_prozent As Variant)
Dim i As Integer
Dim k As Integer
Dim n, m As Integer
Dim low As Integer
Dim high As Integer
Dim gesamt As Integer
Dim anz_pro_statement As Integer
Dim numb_choices As Integer
Dim numb_statements As Integer
Dim Wuerfel_array As Variant
Dim zahl_array As Variant
Dim wuerfel_array0 As Variant
Dim Verbl_Statement As Variant
Dim anz_ziehen As Integer
Dim is_ziehen As Integer
Dim is_gezogen As Integer
Dim anz_antwort As Variant
Dim Anz_Statement As Variant
Dim valid As Boolean
    anz_antwort = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    zahl_array = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
    Anz_Statement = anz_antwort
    wuerfel_array0 = anz_antwort
    low = Worksheets("Gen_output").Cells(13, 2).Value
    high = Worksheets("Gen_output").Cells(14, 2).Value
    numb_choices = Worksheets("Gen_output").Cells(16, 2).Value
    numb_statements = Worksheets("Gen_output").Cells(21, 2).Value
    Worksheets("Response_Matrix").Range("B4:V99").ClearContents
    
' anzahl der Optionen pro Auswahl festlegen --> reihe
    gesamt = 0
    For i = 1 To numb_choices
        anz_antwort(i) = Wuerfeln(high, low)
        gesamt = gesamt + anz_antwort(i)

    Next i
' anzahl der Optionen pro statement festlegen --> spalte
    
    anz_pro_statement = WorksheetFunction.RoundUp(gesamt / numb_statements, 0)
    For i = 1 To numb_statements
        Anz_Statement(i) = anz_pro_statement
    Next i
    k = gesamt Mod numb_statements
    If k > 0 Then
        k = numb_statements - k
    End If
    Call Array_Wuerfeln(Wuerfel_array, wuerfel_array0, zahl_array, numb_statements, k)
    For i = 1 To k
        Anz_Statement(Wuerfel_array(i)) = Anz_Statement(Wuerfel_array(i)) - 1
    Next i
    

' Würfeln
' damit es klappt: erst wird geschaut, wo aufgefüllt werden muss, dann wird der rest ausgelost
    
    For i = 1 To numb_statements
        n = 0
        m = 0
        For k = 1 To numb_choices
           If anz_antwort(k) > 0 Then
                If anz_antwort(k) = (numb_statements - i + 1) Then
                    ' das sind die, die ausgefüllt werden
                    zahl_array(numb_choices - m) = k
                    m = m + 1
                Else
                    ' das sind die, die gewürfelt werden müssen
                    n = n + 1
                    zahl_array(n) = k
                End If
           End If
        Next k

        Call Array_Wuerfeln(Wuerfel_array, wuerfel_array0, zahl_array, n, Anz_Statement(i) - m)
        For k = 1 To m
            ' die nicht gewürfelten noch übertragen
            Wuerfel_array(Anz_Statement(i) - m + k) = zahl_array(numb_choices - k + 1)
        Next k
        For k = 1 To Anz_Statement(i)
            Worksheets("Response_Matrix").Cells(3 + Wuerfel_array(k), 1 + i).Value = gez_prozent(i)
            Worksheets("Response_Matrix").Cells(3 + Wuerfel_array(k), 12 + i).Value = gez_prozent(i)
            anz_antwort(Wuerfel_array(k)) = anz_antwort(Wuerfel_array(k)) - 1
        Next k

    Next i
    
End Sub

Private Function Wuerfeln(ByVal Obere_Grenze As Integer, ByVal untere_Grenze As Integer) As Integer
    Wuerfeln = Int(Rnd() * (Obere_Grenze - untere_Grenze + 1)) + untere_Grenze
End Function
Private Sub Array_Wuerfeln(ByRef Wuerfel_array As Variant, ByVal wuerfel_array0 As Variant, ByVal zahl_array As Variant, ByVal Obere_Grenze As Integer, ByVal Anzahl As Integer)

Dim gewuerfelt As Integer
Dim i As Integer
' cave: aktuelle Version maximal bis 20
    Wuerfel_array = wuerfel_array0
    For i = 1 To Anzahl
        gewuerfelt = Wuerfeln(Obere_Grenze - i + 1, 1)
        Wuerfel_array(i) = zahl_array(gewuerfelt)
        zahl_array(0) = zahl_array(gewuerfelt)
        zahl_array(gewuerfelt) = zahl_array(Obere_Grenze - i + 1)
        zahl_array(Obere_Grenze - i + 1) = zahl_array(0)
    Next i
End Sub
Private Sub Generate_Response_Set()
Dim i, k As Integer
Dim numb_choices, numb_statements As Integer
Dim korrekt As Integer
Dim anz_antwort As Integer
Dim max_inverted As Integer
Dim Tobe_inverted As Integer
Dim Wuerfel_array As Variant
Dim wuerfel_array0 As Variant
Dim zahl_array As Variant
Dim auswaehlen As Boolean
Dim identisch_count As Integer
    auswaehlen = True
    numb_choices = Worksheets("Gen_output").Cells(16, 2).Value
    numb_statements = Worksheets("Gen_output").Cells(21, 2).Value
    max_inverted = Worksheets("Gen_output").Cells(15, 2).Value
    korrekt = Wuerfeln(numb_choices, 1)
    wuerfel_array0 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    zahl_array = wuerfel_array0
    identisch_count = 0
    With Worksheets("Response_Matrix")
        While auswaehlen
            For i = 1 To numb_choices
                If i <> korrekt Then
                    anz_antwort = 0
                    For k = 1 To numb_statements
                        .Cells(3 + i, 12 + k).Value = .Cells(3 + i, 1 + k).Value
                        If .Cells(3 + i, 1 + k) <> "" Then
                            anz_antwort = anz_antwort + 1
                            zahl_array(anz_antwort) = k
                        End If
                    Next k
                    Tobe_inverted = Wuerfeln(max_inverted, 1)
                    If (Tobe_inverted > 1) And (Tobe_inverted = anz_antwort) Then
                        Tobe_inverted = Tobe_inverted - 1
                    End If
                    Call Array_Wuerfeln(Wuerfel_array, wuerfel_array0, zahl_array, anz_antwort, Tobe_inverted)
                    
                    For k = 1 To Tobe_inverted
                        .Cells(3 + i, 12 + Wuerfel_array(k)).Value = -1 * .Cells(3 + i, 12 + Wuerfel_array(k)).Value
                    Next k
                End If
            Next i
            auswaehlen = Worksheets("Response_Matrix").Cells(1, 25).Value
            If auswaehlen Then
                identisch_count = identisch_count + 1
                If identisch_count = 100 Then
                    MsgBox ("Generation of question not possible. Please use a lower number of choices and/or a higher number of max inverted and/or a lower number of max choices.")
                    Exit Sub
                End If
                
            End If
        Wend
    
    End With
End Sub
Private Function Val_Check()
    Val_Check = True
End Function
Private Sub select_questions(ByRef gezogen As Variant, ByVal gezogen0 As Variant, ByRef prozent As Variant)
Dim Category As Integer
Dim Anz_category As Variant
Dim i As Variant
Dim x As Variant
Dim valid As Boolean
    valid = False
    gezogen = gezogen0
    Anz_category = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    For i = 1 To 5
        Anz_category(i) = Worksheets("Gen_output").Cells(23, i + 1).Value
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
    x = Start_x
    i = 0
    While Worksheets("questions").Cells(x, Text_y).Value <> ""
        If Worksheets("questions").Cells(x, 1).Value = "x" Then
            i = i + 1
            gezogen(i) = x
            prozent(i) = Worksheets("questions").Cells(x, 4).Value
            Worksheets("questions").Cells(x, 2).Value = Worksheets("questions").Cells(x, 2).Value + 1
        End If
        x = x + 1
    Wend
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
        If (Worksheets("questions").Cells(x, 1).Value <> "x") And (Worksheets("questions").Cells(x, 2).Value = Min) And _
            (Worksheets("questions").Cells(x, 3).Value = Category) Then
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

Private Function Code_Generieren(ByVal gezogen As Variant) As String
Dim i, Anzahl As Integer
Dim prozente As String
Dim Extra As Integer
Dim text_correct As String
Dim text_wrong As String
Dim text_combi As String
Dim is_equal As String
Dim numb_choices As Integer
Dim numb_statements As Integer
Dim Nummerierung As Variant
    Nummerierung = Array("0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J")
    numb_choices = Worksheets("Gen_output").Cells(16, 2).Value
    numb_statements = Worksheets("Gen_output").Cells(21, 2).Value
    Code_Generieren = "<p>" + Worksheets("questions").Cells(3, 3).Value + "</p>" + vbLf
    For i = 1 To numb_statements
        Code_Generieren = Code_Generieren + "<p>" + Nummerierung(i) + ": " + Worksheets("questions").Cells(gezogen(i), 5).Value + "</p>"
    Next i
    Code_Generieren = Code_Generieren + vbLf + "<p><br>" + Worksheets("questions").Cells(4, 3).Value + "<br></p>" + vbLf
    Code_Generieren = Code_Generieren + "<p>{" + Trim(Str(Worksheets("Gen_output").Cells(26, 2).Value)) + ":" + Worksheets("Gen_output").Cells(25, 2).Value + ":"
    For i = 1 To numb_choices
        text_correct = Get_Combi(1, i, numb_statements, Nummerierung)
        text_wrong = Get_Combi(-1, i, numb_statements, Nummerierung)
        text_combi = ""
        If text_correct <> "" And text_wrong <> "" Then
            If Worksheets("Gen_output").Cells(11, 2).Value = "Deutsch" Then
                text_combi = " und "
            Else
                text_combi = " and "
            End If
        End If
        is_equal = ""
        If Worksheets("Response_Matrix").Cells(3 + i, 1).Value Then
            is_equal = "="
        End If
        Code_Generieren = Code_Generieren + is_equal + text_correct + text_combi + text_wrong + "."
        If i < numb_choices Then
            Code_Generieren = Code_Generieren + " ~"
        End If
    Next i
    Code_Generieren = Code_Generieren + "}</p>"
End Function
Private Function Get_Combi(ByVal typ As Integer, ByVal number As Integer, ByVal numb_statements As Integer, ByVal Nummerierung As Variant) As String
Dim i As Integer
Dim is_combi As Variant
Dim anz_combi As Integer
Dim P_And As Variant
Dim P_Are As Variant
Dim P_Is As Variant
Dim P_Correct As Variant
Dim P_Wrong As Variant
Dim Language As Integer
    is_combi = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    P_And = Array(" and ", " und ")
    P_Are = Array(" are", " sind")
    P_Is = Array(" is", " ist")
    P_Correct = Array(" correct", " richtig")
    P_Wrong = Array(" wrong", " falsch")
    
    Language = 0
    If Worksheets("Gen_output").Cells(11, 2).Value = "Deutsch" Then
        Language = 1
    End If
    anz_combi = 0
    For i = 1 To numb_statements
        If Worksheets("Response_Matrix").Cells(3 + number, 12 + i).Value = typ Then
            anz_combi = anz_combi + 1
            is_combi(anz_combi) = i
        End If
    Next i
    Get_Combi = ""
    For i = 1 To anz_combi
        Get_Combi = Get_Combi + Nummerierung(is_combi(i))
        Select Case i
            Case Is < (anz_combi - 1)
                Get_Combi = Get_Combi + ", "
            Case (anz_combi - 1)
                Get_Combi = Get_Combi + P_And(Language)
        End Select
    Next i
    If anz_combi > 0 Then
        If anz_combi > 1 Then
            Get_Combi = Get_Combi + P_Are(Language)
        Else
            Get_Combi = Get_Combi + P_Is(Language)
        End If
        If typ = 1 Then
            Get_Combi = Get_Combi + P_Correct(Language)
        Else
            Get_Combi = Get_Combi + P_Wrong(Language)
        End If
    End If

End Function


Private Sub init(ByVal Ziehen_Typ As Integer, ByRef gezogen As Variant, ByRef gezogen0 As Variant, ByRef gez_prozent As Variant)
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
    
End Sub


Private Function XML_Header(number As Integer) As String
    XML_Header = vbLf + "<question type=""cloze"">" + vbLf + "<name><text>" + _
        Worksheets("Gen_output").Cells(6, 2).Value + " - " + Trim(Str(number)) + "</text></name>" + vbLf + _
           "<questiontext format=""html"">" + vbLf + "<text><![CDATA[" + vbLf
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
            If .Cells(x, 9).Value = "html" Then
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


