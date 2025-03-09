Option Explicit
Const start_x As Integer = 9
Const Text_y As Integer = 5
Const Spalten As Integer = 7   ' note: Spalten = y Ende...
Const XML_Out_X As Integer = 17
Const wrong As String = "xxxxxxxxxxxxxxxxxxxxxx"
Const Question_type As String = "<!-- xx SC xx -->"
Const New_Question As String = "<!-- xx NewQuestion xx -->"
Const Anz_Link_Fields As Integer = 2          ' Linkfelder für die Zufallsgenerierung im Tabellenblatt Rnd_Matrix
Sub test()
Dim teststr As String
    teststr = export_plain_html(0)
    Debug.Print (teststr)
End Sub
Function export_plain_html(ByVal Ziehen As Integer) As String
' diese Funktion ist für einen externen Aufruf zur Erstellung von HTML Klausuren
' ziehen: 0: nicht ziehen; 1: ziehen, nicht zurücksetzen; 2: ziehen und zurücksetzen
' cave: bei dieser Fragenart multiple response single choice ist das erstellen der Frage ohne ziehen problematisch!
Dim gezogen, gezogen0, gez_prozent As Variant
Dim x As Integer
Dim bonus_comment As Boolean
Dim zufall As Boolean

    If Ziehen = 2 Then
        Randomize
        x = start_x
        With Worksheets("Questions")
            While .Cells(x, 3).Value <> ""
                .Cells(x, 2).Value = ""
                x = x + 1
            Wend
        End With
    End If
    x = start_x
    If Ziehen > 0 Then
        Call Draw_Questions(False)
    End If
    export_plain_html = Generate_Question_HTML(True)
End Function

Sub Copy_HTML()
Dim oData As New DataObject
Dim Cloze_SC_Fragen As String
    Cloze_SC_Fragen = Generate_Question_HTML(False)
    With oData
        .SetText Cloze_SC_Fragen
        .PutInClipboard
    End With
End Sub
Private Function Generate_Question_HTML(ByVal Include_HTML_Separators As Boolean) As String
Dim x As Integer
Dim Cloze_SC_Fragen As String
Dim Nummer As Integer
Dim HTML_separators As String
    HTML_separators = ""
    If Include_HTML_Separators Then
        HTML_separators = Question_type + vbLf
    End If
    With Worksheets("Questions")
        Cloze_SC_Fragen = HTML_separators + "<p>" + .Cells(2, 2).Value + "</p>" + vbLf
        x = start_x
        Nummer = 1
        If Include_HTML_Separators Then
            HTML_separators = New_Question + Question_type + vbLf
        End If

        While .Cells(x, Text_y).Value <> ""
            If .Cells(x, 1).Value = "x" Then
                If Nummer > 1 Then
                    Cloze_SC_Fragen = Cloze_SC_Fragen + HTML_separators
                End If
                Cloze_SC_Fragen = Cloze_SC_Fragen + Cloze_SC_Frage_generieren(x, Nummer)
                Nummer = Nummer + 1
            End If
            x = x + 1
        Wend
    End With
    Generate_Question_HTML = Cloze_SC_Fragen
End Function
Private Function Cloze_SC_Frage_generieren(ByVal x As Integer, ByVal Nummer As Integer) As String
Dim Fragentyp As String
Dim moodle_code As String
Dim Numbering As String
Dim moodle As Boolean
Dim y As Integer
    With Worksheets("Questions")
        Fragentyp = .Cells(3, 2).Value
        If Fragentyp = "" Then
            Fragentyp = "MCVS"
        End If
        y = Text_y + 1
        If .Cells(x, y) <> "" Then
            moodle_code = "{1:" + Fragentyp + ":=" + .Cells(x, y).Value + " ~" + .Cells(x, y + 1).Value
        Else
            moodle_code = "{1:" + Fragentyp + ":" + .Cells(x, y + 1).Value
        End If
        y = y + 2
        While .Cells(x, y).Value <> ""
            moodle_code = moodle_code + " ~" + .Cells(x, y).Value
            y = y + 1
        Wend
        moodle_code = moodle_code + "}"
        If .Cells(7, 2).Value Then
            Numbering = "<b>" + Trim(Str(Nummer)) + ")</b> "
        Else
            Numbering = ""
        End If
        Cloze_SC_Frage_generieren = vbLf + "<p><br>" + Numbering + .Cells(x, Text_y).Value + "<br>" + vbLf + moodle_code + "<p>" + vbLf
    End With
End Function

Sub Draw_Questions(ByVal Just_draw As Boolean)
Dim sel_string As String
Dim Category As Integer
Dim Amount_Per_Category As Variant
    Call Init_Select(Category, Amount_Per_Category)
    While Category > 0
        While Amount_Per_Category(Category) > 0
            Call Draw_Text(Category)
            Amount_Per_Category(Category) = Amount_Per_Category(Category) - 1
        Wend
        Category = Category - 1
    Wend
    If Just_draw Then
        sel_string = InputBox("Actual selection, if required adapt (i.e. change rows; comma at the end needs to remain!).", "Adapt selection?", Generate_Selection_String())
        Call Adapt_Selection(sel_string)
    End If
    Call Increase_Count
End Sub
Private Sub Init_Select(ByRef Category As Integer, ByRef Amount_Per_Category As Variant)
Dim x As Integer
Dim i As Integer
    Randomize
    Category = 0
    x = start_x
    Amount_Per_Category = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    i = 1
    With Worksheets("Questions")
        While .Cells(5, i).Value <> ""
            Amount_Per_Category(i) = .Cells(6, i).Value
            i = i + 1
        Wend
        While .Cells(x, Text_y).Value <> ""
            If .Cells(x, 3).Value = 0 Then
                .Cells(x, 1).Value = "x"
            Else
                .Cells(x, 1).Value = ""
            End If
            If .Cells(x, 2).Value = "" Then
                .Cells(x, 2).Value = 0
            End If
            If .Cells(x, 3).Value > Category Then
                Category = .Cells(x, 3).Value
            End If
            x = x + 1
        Wend
    End With
End Sub


Private Sub Anzahl(ByRef Anzahl As Integer, ByRef Min As Integer, ByVal Category As Integer)
Dim x As Integer
    Anzahl = 0
    Min = 100
    x = start_x
    With Worksheets("Questions")
        While .Cells(x, 3).Value <> ""
            If (.Cells(x, 3).Value = Category) And (.Cells(x, 1).Value <> "x") And (Min > .Cells(x, 2).Value) Then
                Min = .Cells(x, 2).Value
            End If
            x = x + 1
        Wend
        x = start_x
        While .Cells(x, 3).Value <> ""
            If (.Cells(x, 3).Value = Category) And (.Cells(x, 1).Value <> "x") And (Min = .Cells(x, 2).Value) Then
                Anzahl = Anzahl + 1
            End If
            x = x + 1
        Wend
    End With
End Sub
Private Sub Draw_Text(ByVal Category As Integer)
Dim x As Integer
Dim i As Integer
Dim selected As Boolean
Dim Min As Integer
Dim numb_count As Integer
    x = start_x
    Call Anzahl(numb_count, Min, Category)
    i = Int(Rnd * numb_count)
    selected = False
    With Worksheets("Questions")
        While Not (selected)
            If (.Cells(x, 1).Value <> "x") And (.Cells(x, 2).Value = Min) And (.Cells(x, 3).Value = Category) Then
                If i = 0 Then
                    selected = True
                    .Cells(x, 1).Value = "x"
                Else
                    i = i - 1
                End If
            End If
            x = x + 1
        Wend
    End With
End Sub
Private Sub Increase_Count()
Dim x As Integer
    x = start_x
    With Worksheets("Questions")
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
    x = start_x
    actual_sel = ""
    With Worksheets("Questions")
        While .Cells(x, Text_y).Value <> ""
            If .Cells(x, 1).Value = "x" Then
                actual_sel = actual_sel + Trim(Str(x)) + ", "
            End If
            x = x + 1
        Wend
    End With
    Generate_Selection_String = actual_sel
End Function
Private Sub Adapt_Selection(ByVal sel_string As String)
Dim x As Integer
    x = start_x
    With Worksheets("Questions")
        While .Cells(x, Text_y).Value <> ""
            .Cells(x, 1).Value = ""
            x = x + 1
        Wend
        While sel_string <> ""
            x = Get_Number(sel_string)
            Cells(x, 1).Value = "x"
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
    x = start_x
    With Worksheets("Questions")
        While .Cells(x, 3).Value <> ""
            .Cells(x, 2).Value = ""
            x = x + 1
        Wend
    End With
    For number = 1 To Worksheets("Gen_output").Cells(8, 2).Value
        Call Draw_Questions(False)
        question_string = XML_Header(number) + Generate_Question_HTML(False) + XML_End(number)
        substitute = Int(Rnd() * (max_substitute)) + 1
        Call Print_Replace(Text_File, substitute, question_string)
    Next number
    Print #Text_File, vbLf + "</quiz>"
    Close Text_File

End Sub
Private Function XML_Header(number As Integer) As String
    XML_Header = vbLf + "<question type=""cloze"">" + vbLf + "<name><text>" + _
        Worksheets("Gen_output").Cells(4, 2).Value + " - " + Trim(Str(number)) + "</text></name>" + vbLf + "<questiontext format=""html"">" + vbLf + "<text><![CDATA[" + vbLf
End Function
Private Function XML_End(number As Integer) As String
Dim XML_End_Text As String
Dim n_tags As Integer
Dim int_tag_text As String
Dim x As Integer
    Worksheets("Gen_output").Cells(20, 2).Value = "'" + Trim(Str(number))
    XML_End_Text = vbLf + "]]></text></questiontext>"
    x = XML_Out_X
    While Worksheets("Gen_output").Cells(x, 1).Value <> ""
        If Worksheets("Gen_output").Cells(x, 1).Value = "hint" Then
            int_tag_text = Get_Tags(x, n_tags)
            XML_End_Text = XML_End_Text + Gen_XML(x, int_tag_text)
            x = x + 1 + n_tags
        Else
            int_tag_text = ""
            XML_End_Text = XML_End_Text + Gen_XML(x, int_tag_text)
            x = x + 1
        End If
    Wend
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
                If .Cells(x, 3).Value = "tag" Then
                    If .Cells(x, 2).Value = 1 Then
                        xml_str = "<" + FieldN + "/>"
                    End If
                Else
                    xml_str = vbLf + "<" + FieldN + ">" + Trim_Hyp(.Cells(x, 2).Value) + "</" + FieldN + ">"
                End If
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