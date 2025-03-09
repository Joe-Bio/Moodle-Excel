Option Explicit
Const start_x As Integer = 9
Const Text_y As Integer = 5
Const Spalten As Integer = 7   ' note: Spalten = y Ende...
Const wrong As String = "xxxxxxxxxxxxxxxxxxxxxx"
Const Anz_Link_Fields As Integer = 2          ' Linkfelder für die Zufallsgenerierung im Tabellenblatt Rnd_Matrix

Sub test()
Dim teststr As String
    teststr = export_plain_html(2)
    Worksheets("Gen_output").Cells(18, 1).Value = teststr
End Sub
Function export_plain_html(ByVal Ziehen As Integer) As String
' diese Funktion ist für einen externen Aufruf zur Erstellung von HTML Klausuren
' ziehen: 0: nicht ziehen; 1: ziehen, nicht zurücksetzen; 2: ziehen und zurücksetzen
' cave: bei dieser Fragenart multiple response single choice ist das erstellen der Frage ohne ziehen problematisch!
Dim gezogen, gezogen0, gez_prozent As Variant
Dim i As Integer
Dim x As Integer
Dim typ As Integer
Dim bonus_comment As Boolean
Dim question_type As String
Dim cloze_text As String
'        Call Select_Cloze(False)
'        cloze_text = XML_Header(i)
'        With Worksheets("questions")
'            cloze_text = cloze_text + "<p>" + .Cells(2, 2).Value + "</p>" + vbLf + "<p>"
'            x = start_x
'            While .Cells(x, Text_y).Value <> ""
'                If .Cells(x, 1).Value = "x" Then
'                    cloze_text = cloze_text + Cloze_Zeile_generieren(x)
'                End If
'                x = x + 1
'            Wend
'            cloze_text = cloze_text + "</p>"
'        End With
'        cloze_text = cloze_text + XML_End(i)

    question_type = "<!-- xx Assign xx " + Trim(Str(Worksheets("Gen_output").Cells(14, 2).Value)) + " xx -->"
    If Ziehen > 0 Then
        If Ziehen > 1 Then
            Worksheets("questions").Range("A9:B200").ClearContents
        End If
        Call Select_Cloze(False)
        cloze_text = ""
        With Worksheets("questions")
            cloze_text = cloze_text + "<p>" + .Cells(2, 2).Value + "</p>" + vbLf + "<p>"
            x = start_x
            While .Cells(x, Text_y).Value <> ""
                If .Cells(x, 1).Value = "x" Then
                    cloze_text = cloze_text + Cloze_Zeile_generieren(x)
                End If
                x = x + 1
            Wend
            cloze_text = cloze_text + "</p>"
        End With
    End If
    export_plain_html = question_type + cloze_text
End Function

Sub Export_Moodle_XML()
Dim Filepath As String
Dim Text_File As Integer
Dim x As Integer
Dim i, n As Integer
Dim col As Integer
Dim tot_col As Integer
Dim cloze_text As String
Dim it_start As String
Dim it_end As String
Dim anz As Integer
Dim Question_String As String
Dim max_substitute As Integer
Dim substitute As Integer
Dim Selected, selected0 As Variant
    selected0 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    With Worksheets("Gen_output")
        Filepath = .Cells(4, 2) + .Cells(5, 2)
        anz = .Cells(12, 2)
        max_substitute = .Cells(16, 2).Value
    End With
    Worksheets("questions").Range("A9:B200").ClearContents
    Text_File = FreeFile
    Open Filepath For Output As Text_File
    tot_col = Worksheets("Gen_output").Cells(23, 2).Value
    Print #Text_File, vbLf + "<quiz>" + vbLf
    For i = 1 To anz
        Selected = selected0
        Call select_assignments
        cloze_text = XML_Header(i)
        With Worksheets("questions")
            cloze_text = cloze_text + "<p>" + .Cells(2, 2).Value + "</p>" + vbLf + "<table cellpadding=""5"" border=""1"">" + vbLf + "<thead>" + vbLf
            ' header schreiben
            cloze_text = cloze_text + "<tr align=""center"">" + vbLf
            For col = 1 To tot_col
                cloze_text = cloze_text + "<th>" + Worksheets("Gen_output").Cells(25, 2).Value + "</th>" + vbLf
                cloze_text = cloze_text + "<th>" + Worksheets("Gen_output").Cells(26, 2).Value + "</th>" + vbLf
            Next col
            cloze_text = cloze_text + "</tr>" + vbLf + "</thead>" + vbLf + "<tbody>" + vbLf
            Call Read_Selected(Selected)
            col = 1
            For n = 1 To Selected(0)
                x = Selected(n)
                If col = 1 Then
                    cloze_text = cloze_text + "<tr align=""center"">" + vbLf
                End If
                If .Cells(x, 4).Value = "Text" Then
                    it_start = "<i>"
                    it_end = "</i>"
                Else
                    it_start = ""
                    it_end = ""
                End If
                cloze_text = cloze_text + "<td>" + it_start + .Cells(x, 5).Value + it_end + "</td>" + vbLf
                cloze_text = cloze_text + "<td>" + it_start + Moodle_code_generieren(x) + it_end + "</td>" + vbLf
                If col = tot_col Then
                    col = 1
                    cloze_text = cloze_text + "</tr>" + vbLf
                Else
                    col = col + 1
                End If
            Next n
            While (col <= tot_col) And (col > 1)
                cloze_text = cloze_text + "<td></td>" + vbLf + "<td></td>" + vbLf
                If col = tot_col Then
                    cloze_text = cloze_text + "<tr>" + vbLf
                    col = 0
                End If
                col = col + 1
            Wend
            cloze_text = cloze_text + "</tbody>" + vbLf + "</table>" + vbLf
        End With
        cloze_text = cloze_text + XML_End(i)
        substitute = Int(Rnd() * (max_substitute)) + 1
        Call Print_Replace(Text_File, substitute, cloze_text)
    Next i
    Print #Text_File, vbLf + "</quiz>"
    Close Text_File

End Sub
Private Sub Read_Selected(ByRef Selected As Variant)
Dim i As Integer
Dim pos As Integer
Dim Zufall As Boolean
Dim x As Integer
    x = start_x
    i = 1
    With Worksheets("questions")
        While .Cells(x, Text_y).Value <> ""
            If .Cells(x, 1).Value = "x" Then
                Selected(i) = x
                i = i + 1
            End If
            x = x + 1
        Wend
    End With
    Selected(0) = i - 1     ' das ist wichtig, den nicht einfach von der Fragenzahl zu übernehmen, da ja auch immer selektierte Textfelder (schwierigkeit 0) drin sein können!
' randomisieren
    If Worksheets("Gen_output").Cells(17, 2).Value Then
        For i = Selected(0) To 2 Step -1
            pos = Int(i * Rnd + 1)
            Selected(Selected(0) + 2) = Selected(pos)
            Selected(pos) = Selected(i)
            Selected(i) = Selected(Selected(0) + 2)
        Next i
    End If
End Sub
Private Function XML_Header(ByVal number As Integer) As String
    XML_Header = vbLf + "<question type=""cloze"">" + vbLf + "<name><text>" + _
        Worksheets("Gen_output").Cells(6, 2).Value + " - " + Trim(Str(number)) + "</text></name>" + vbLf + "<questiontext format=""html"">" + vbLf + "<text><![CDATA[" + vbLf
End Function
Private Function XML_End(ByVal number As Integer) As String
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
Private Function Moodle_code_generieren(ByVal x As Integer) As String
Dim y As Integer
Dim Fragentyp As String
Dim moodle_code As String
Dim moodle As Boolean
Dim position As Integer
With Worksheets("questions")
    Fragentyp = .Cells(5, 2).Value
    If .Cells(x, 4).Value <> "" Then
        Fragentyp = .Cells(x, 4).Value
    End If
    If Fragentyp = "" Then
        Fragentyp = "MCS"
    End If
    If Fragentyp <> "Text" Then
        If .Cells(x, Text_y + 1) <> "" Then
            moodle_code = "{1:" + Fragentyp + ":=" + .Cells(x, Text_y + 1).Value + " ~" + .Cells(x, Text_y + 2).Value + "}"
        Else
            moodle_code = "{1:" + Fragentyp + ":" + .Cells(x, Text_y + 2).Value + "}"
        End If
    Else
        moodle_code = .Cells(x, Text_y + 1).Value
    End If
    Moodle_code_generieren = moodle_code
End With
End Function

Private Sub select_assignments()
Dim Category As Integer
Dim Anz_category As Variant
Dim i As Variant
Dim x As Variant
    Anz_category = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    For i = 1 To 5
        Anz_category(i) = Worksheets("Gen_output").Cells(21, i + 1).Value
    Next i
    x = start_x
    While Worksheets("questions").Cells(x, Text_y).Value <> ""
        Worksheets("questions").Cells(x, 1).Value = ""
        x = x + 1
    Wend
    For Category = 1 To 5
        While Anz_category(Category) > 0
            Call Draw_Assignment(Category)
            Anz_category(Category) = Anz_category(Category) - 1
        Wend
    Next Category
    ' category 0 ausfüllen
    x = start_x
    With Worksheets("questions")
        While .Cells(x, Text_y).Value <> ""
            If .Cells(x, 3).Value = 0 Then
                .Cells(x, 1).Value = "x"
                .Cells(x, 2).Value = .Cells(x, 2).Value + 1
            End If
            x = x + 1
        Wend
    End With
End Sub

Private Sub Draw_Assignment(ByVal Category As Integer)
Dim x As Integer
Dim i As Integer
Dim Selected As Boolean
Dim Min As Integer
Dim numb_count As Integer
    x = start_x
    Call Anzahl_cat(numb_count, Min, Category)
    i = Int(Rnd * numb_count)
    Selected = False
    While Not (Selected)
        If (Worksheets("questions").Cells(x, 1).Value <> "x") And (Worksheets("questions").Cells(x, 2).Value = Min) And (Worksheets("questions").Cells(x, 3).Value = Category) Then
            If i = 0 Then
                Selected = True
                Worksheets("questions").Cells(x, 1).Value = "x"
                Worksheets("questions").Cells(x, 2).Value = Worksheets("questions").Cells(x, 2).Value + 1
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
    x = start_x
    While Worksheets("questions").Cells(x, Text_y).Value <> ""
        If (Worksheets("questions").Cells(x, 3).Value = Category) And (Worksheets("questions").Cells(x, 1).Value <> "x") And (Min > Worksheets("questions").Cells(x, 2).Value) Then
            Min = Worksheets("questions").Cells(x, 2).Value
        End If
        x = x + 1
    Wend
    x = start_x
    While Worksheets("questions").Cells(x, Text_y).Value <> ""
        If (Worksheets("questions").Cells(x, 3).Value = Category) And (Worksheets("questions").Cells(x, 1).Value <> "x") And (Min = Worksheets("questions").Cells(x, 2).Value) Then
            Anzahl = Anzahl + 1
        End If
        x = x + 1
    Wend
End Sub

Private Sub Increase_Count()
Dim x As Integer
With Worksheets("questions")
    x = start_x
    While .Cells(x, Text_y).Value <> ""
        If .Cells(x, 1).Value = "x" Then
            .Cells(x, 2).Value = .Cells(x, 2).Value + 1
        End If
        x = x + 1
    Wend
End With
End Sub
Private Function Generate_Selection_String() As String
Dim actual_sel As String
Dim x As Integer
With Worksheets("questions")
    x = start_x
    actual_sel = ""
    While .Cells(x, Text_y).Value <> ""
        If .Cells(x, 1).Value = "x" Then
            actual_sel = actual_sel + Trim(Str(x)) + ", "
        End If
        x = x + 1
    Wend
    Generate_Selection_String = actual_sel
End With
End Function
Private Sub Adapt_Selection(ByVal sel_string As String)
Dim x As Integer
With Worksheets("questions")
    x = start_x
    While .Cells(x, Text_y).Value <> ""
        .Cells(x, 1).Value = ""
        x = x + 1
    Wend
    While sel_string <> ""
        x = Get_Number(sel_string)
        .Cells(x, 1).Value = "x"
    Wend
End With
End Sub
Private Function Get_Number(ByRef sel_string As String) As Integer
Dim i As Integer
Dim x_string As String
Dim remain_string As String
    i = InStr(sel_string, ",")
    Get_Number = Round(Val(Left(sel_string, i)), 0)
    sel_string = Right(sel_string, Len(sel_string) - i)
    If Len(sel_string) < 2 Then sel_string = ""
End Function
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
        To_Be_Cleaned = Replace(To_Be_Cleaned, non_html(i), is_html(i))
    Next i
End Sub

