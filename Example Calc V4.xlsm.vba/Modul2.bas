Option Explicit
Const start_x As Integer = 7
Const start_intro_x As Integer = 4
Const start_sol_x As Integer = 5
Const dec_del As String = "."
Const Wrong_No As String = "~999999"
Const Zellen_Verbinden As Boolean = True
Const Text_y As Integer = 3
Const Anz_Link_Fields As Integer = 2          ' Linkfelder für die Zufallsgenerierung im Tabellenblatt Rnd_Matrix
Sub test()
Dim exp_string As String
    exp_string = export_plain_html(1)
    Debug.Print (exp_string)
End Sub
Function export_plain_html(ByVal Ziehen As Integer) As String
' diese Funktion ist für einen externen Aufruf zur Erstellung von HTML Klausuren
' ziehen: 0: nicht ziehen; 1: ziehen, nicht zurücksetzen; 2: ziehen und zurücksetzen
' cave: bei dieser Fragenart multiple response single choice ist das erstellen der Frage ohne ziehen problematisch!
Dim Question_String As String
Dim Intro_Text As String
Dim Table_Text As String
Dim Text_File As Integer
Dim Solution_text As String
Dim number_set As Integer
Dim anz_number As Integer
Dim var_x_init As Variant
Dim var_x_sol As Variant
Dim Include_ID As Boolean
Dim i As Integer
Dim n_row As Integer
Dim n_col As Integer
Dim x_tab As Variant

    Call init_write(var_x_init, var_x_sol, Text_File, anz_number, Include_ID, True)
    Include_ID = False
    With Worksheets("Gen_output")
        Question_String = .Cells(29, 2).Value + vbLf
        If Ziehen = 2 Then
            .Cells(30, 2).Value = 0
        End If
        number_set = .Cells(30, 2).Value
        If number_set = anz_number Then
            number_set = 1
        Else
            number_set = number_set + 1
        End If
        .Cells(30, 2).Value = number_set
    End With
    Question_String = Question_String + "<!-- xx IdNumber xx " + Str(number_set) + " xx -->" + vbLf
    Call Transfer_Variables(var_x_init, number_set)
    Call Adopt_Format(var_x_init, var_x_sol)
    Intro_Text = Write_Intro(True)
    Solution_text = Write_Solution(number_set, True)
    
    Question_String = Question_String + Intro_Text + Table_Text + "<p><br></p>" + vbLf + Solution_text
    export_plain_html = Question_String
End Function

Sub Gen_Rand_Sets()
Dim var_x_init As Variant
Dim var_x_sol As Variant
Dim number As Integer
Dim anz_draw As Integer
Dim Draw_Random As Boolean
Dim respect_limits As Boolean
    Call init_rand(var_x_init, var_x_sol)
    respect_limits = Cells(3, 2).Value
    anz_draw = Cells(2, 2).Value
    Draw_Random = (anz_draw > 0)
    If anz_draw = 0 Then
        anz_draw = 1
    End If
    For number = 1 To anz_draw
        If Draw_Random Then
            Call Draw_Init(var_x_init)
            While NoLimits(respect_limits)
                Call Draw_Init(var_x_init)
            Wend
        End If
        Call Write_Dataset(number, var_x_init, var_x_sol)
    Next number
    
End Sub
Private Sub init_rand(ByRef var_x_init As Variant, ByRef var_x_sol As Variant)
Dim y As Integer
Dim x As Integer
Dim i As Integer
Dim n As Integer
    Range("A6:AZ200").ClearContents
    Call Clear_Evaluation
    Randomize
    var_x_init = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    var_x_sol = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    i = 1
    x = start_intro_x
    While Worksheets("Intro").Cells(x, 2).Value <> ""
        If Worksheets("Intro").Cells(x, 5).Value <> "" Then     ' es werden nur variablen berücksichtigt, die über ein lower limit verfügen - sonst bleiben die Variablen unberührt!
            var_x_init(i) = x
            i = i + 1
        End If
        x = x + 1
    Wend
    var_x_init(0) = i - 1

    i = 1
    x = start_sol_x
    While Worksheets("Solution").Cells(x, 1).Value <> ""
        If Worksheets("Solution").Cells(x, 10).Value <> "" Then
            var_x_sol(i) = x
            i = i + 1
        End If
        x = x + 1
    Wend
    var_x_sol(0) = i - 1

    Cells(start_x - 1, 1).Value = "Number"
    x = start_intro_x
    i = 0
    While Worksheets("Intro").Cells(x, 2).Value <> ""
        If Worksheets("Intro").Cells(x, 8).Value <> "" Then
            i = i + 1
            Cells(start_x - 1, 1 + i).Value = Worksheets("Intro").Cells(x, 8).Value
        End If
        x = x + 1
    Wend
    For n = 1 To var_x_sol(0)
        Cells(start_x - 1, 1 + i + n).Value = Worksheets("Solution").Cells(var_x_sol(n), 10).Value
    Next n
End Sub
Private Function NoLimits(ByVal respect_limits As Boolean) As Boolean
Dim Incl_Table As Boolean
    NoLimits = respect_limits
    If respect_limits Then
        NoLimits = Not (Worksheets("Solution").Cells(3, 2).Value)
    End If
End Function
Private Sub Write_Dataset(ByVal number As Integer, ByVal var_x_init As Variant, ByVal var_x_sol As Variant)
Dim x_dataset As Integer
Dim x As Integer
Dim i As Integer
Dim n As Integer
    x_dataset = number + start_x - 1
    Cells(x_dataset, 1).Value = number
    i = 0
    x = start_intro_x
    While Worksheets("Intro").Cells(x, 2).Value <> ""
        If Worksheets("Intro").Cells(x, 8).Value <> "" Then
            i = i + 1
            Cells(x_dataset, i + 1).Value = Worksheets("Intro").Cells(x, 4).Value
        End If
        x = x + 1
    Wend
    For n = 1 To var_x_sol(0)
        Cells(x_dataset, i + n + 1).Value = Worksheets("Solution").Cells(var_x_sol(n), 5).Value
    Next n
End Sub
Private Sub Draw_Init(ByVal var_x_init As Variant)
Dim i As Integer
Dim drawn As Variant
Dim ll As Variant
Dim ul As Variant
Dim digits As Integer
    For i = 1 To var_x_init(0)
        ll = Worksheets("Intro").Cells(var_x_init(i), 5).Value
        ul = Worksheets("Intro").Cells(var_x_init(i), 6).Value
        digits = Worksheets("Intro").Cells(var_x_init(i), 7).Value
        drawn = ll + (Rnd() * (ul - ll))
        drawn = Return_digits(drawn, digits)
        Worksheets("Intro").Cells(var_x_init(i), 4).Value = drawn
    Next i
End Sub
Private Function Return_digits(ByVal number As Variant, ByVal sig_digits As Integer) As Variant
Dim potenz As Integer
    potenz = Int(Log_10(number))
    number = number / (10 ^ potenz)
    number = 10 ^ potenz * Round(number, sig_digits - 1)
    Return_digits = number
End Function
Sub Write_Moodle()
Dim Intro_Text As String
Dim Table_Text As String
Dim XML_String As String
Dim Text_File As Integer
Dim Solution_text As String
Dim number_set As Integer
Dim anz_number As Integer
Dim var_x_init As Variant
Dim var_x_sol As Variant
Dim Include_ID As Boolean
Dim i As Integer
Dim max_substitute As Integer
Dim substitute As Integer

    Call init_write(var_x_init, var_x_sol, Text_File, anz_number, Include_ID, False)
    max_substitute = Worksheets("Gen_output").Cells(32, 2).Value
    For number_set = 1 To anz_number
        Call Transfer_Variables(var_x_init, number_set)
        Call Adopt_Format(var_x_init, var_x_sol)
        substitute = Int(Rnd() * (max_substitute)) + 1
        Intro_Text = Write_Intro(True)
        Solution_text = Write_Solution(number_set, True)
        XML_String = XML_Header(number_set) + Intro_Text + Table_Text + "<p><br></p>" + vbLf + Solution_text + XML_End(number_set)
        Call Print_Replace(Text_File, substitute, XML_String)
    Next number_set
    
    Print #Text_File, vbLf + "</quiz>"
    Close Text_File

End Sub
Private Sub init_write(ByRef var_x_init As Variant, ByRef var_x_sol As Variant, ByRef Text_File As Integer, ByRef anz_number As Integer, ByRef Include_ID As Boolean, ByVal HTML_exp As Boolean)
Dim y As Integer
Dim x As Integer
Dim i As Integer
Dim n_he As Integer
Dim FilePath As String
    Call Clean_HTML
    Call Clear_Evaluation
    var_x_init = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)    ' hier stehen die zu übertragenden Variablen
    var_x_sol = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)     ' hier stehen die zu formatierenden Zahlen
    anz_number = Worksheets("Dataset").Cells(2, 2).Value
    If anz_number = 0 Then
        anz_number = 1
    End If
    Call Get_Variables_X(var_x_init)
    
    i = 1
    x = start_sol_x
    While Worksheets("Solution").Cells(x, 1).Value <> ""
        If (Worksheets("Solution").Cells(x, 9).Value <> "") And (Worksheets("Solution").Cells(x, 1).Value <> "Hidden") Then
            var_x_sol(i) = x
            i = i + 1
        End If
        x = x + 1
    Wend
    var_x_sol(0) = i - 1
    ' init xml_string
    With Worksheets("Gen_output")
        FilePath = .Cells(5, 2).Value + .Cells(6, 2).Value
        Include_ID = .Cells(7, 2).Value
    End With
    If Not (HTML_exp) Then
        Text_File = FreeFile
        Open FilePath For Output As Text_File
        Print #Text_File, vbLf + "<quiz>" + vbLf
    End If
End Sub
Private Sub Get_Variables_X(ByRef var_x_init As Variant)
Dim i As Integer
Dim x As Integer
    i = 1
    x = start_intro_x
    While Worksheets("Intro").Cells(x, 2).Value <> ""
        If Worksheets("Intro").Cells(x, 8).Value <> "" Then
            var_x_init(i) = x
            i = i + 1
        End If
        x = x + 1
    Wend
    var_x_init(0) = i - 1
End Sub
Private Sub Transfer_Variables(ByVal var_x_init As Variant, ByVal number_set As Integer)
Dim i As Integer
    For i = 1 To var_x_init(0)      ' Datensatz übertragen
        Worksheets("Intro").Cells(var_x_init(i), 4).Value = Worksheets("Dataset").Cells(start_x + number_set - 1, 1 + i).Value
    Next i
End Sub
Private Sub Adopt_Format(ByVal var_x_init As Variant, ByVal var_x_sol As Variant)
Dim i As Integer
Dim zahl As Variant
Dim Grenze_Up As Integer
Dim Grenze_Lo As Integer
Dim Err_100 As Variant
Dim Err_50 As Variant
Dim digits As Integer
Dim potenz As Integer
Dim Resultat As String
Dim Res_str As String
Dim x As Integer
Dim y As Integer
Dim Add_Wrong As String
    With Worksheets("Gen_output")
        Grenze_Up = .Cells(8, 2).Value
        Grenze_Lo = .Cells(9, 2).Value
        Err_100 = .Cells(11, 2).Value
        Err_50 = .Cells(12, 2).Value
    End With
    x = start_intro_x
    With Worksheets("Intro")
        While .Cells(x, 2).Value <> ""
            If .Cells(x, 7).Value <> "" Then
                zahl = .Cells(x, 4).Value
                digits = .Cells(x, 7).Value
                zahl = Zahl_Digits(zahl, digits)
                .Cells(x, 3).Value = "'" + Format_Zahl(zahl, 1, Grenze_Up, Grenze_Lo)
            End If
            x = x + 1
        Wend
    End With
    With Worksheets("Solution")
        For i = 1 To var_x_sol(0)
            zahl = .Cells(var_x_sol(i), 5).Value
            potenz = Int(Log_10(zahl))
            digits = .Cells(var_x_sol(i), 9).Value
            zahl = Zahl_Digits(zahl, digits)
            .Cells(var_x_sol(i), 4).Value = "'" + Format_Zahl(zahl, 2, Grenze_Up, Grenze_Lo)
        Next i
        x = start_sol_x
        While .Cells(x, 1).Value <> ""
            If (.Cells(x, 1).Value <> "Text") And (.Cells(x, 1).Value <> "Hidden") And (.Cells(x, 1).Value <> "HTML_Comment") Then
                Resultat = Trim_Hyp(.Cells(x, 4).Value)
                If .Cells(x, 1).Value = "NM" Then
                    Res_str = Gen_Result_Tolerance(Resultat, .Cells(x, 5).Value, Err_100, Err_50)
                Else
                    Res_str = Trim_Hyp(.Cells(x, 5).Value)
                End If
                If .Cells(x, 1).Value <> "Number" Then
                    Add_Wrong = ""
                    If .Cells(x, 1).Value = "NM" Then
                        Add_Wrong = Wrong_No
                    End If
                    .Cells(x, 4).Value = " {" + Trim(Str(.Cells(x, 2).Value)) + ":" + .Cells(x, 1).Value + ":" + Res_str + Add_Wrong + "}"
                Else
                    .Cells(x, 4).Value = Res_str
                
                End If
            End If
            x = x + 1
        Wend
    End With
End Sub
Private Function Format_Zahl(ByVal zahl As Variant, ByVal Typ As Integer, ByVal p_up As Integer, ByVal p_lo As Integer) As String
Dim form_zahl As String
Dim potenz As Integer
    potenz = Int(Log_10(zahl))
    If (potenz > p_up) Or (potenz < p_lo) Then
        zahl = zahl / (10 ^ potenz)
    End If
    form_zahl = Trim(Str(zahl))
    form_zahl = Replace(form_zahl, ".", dec_del)
    If Left(form_zahl, 1) = dec_del Then
        form_zahl = "0" + form_zahl
    End If
    If Left(form_zahl, 2) = "-" + dec_del Then
        form_zahl = "-0" + Right(form_zahl, Len(form_zahl) - 1)
    End If
    If (potenz > p_up) Or (potenz < p_lo) Then
        If Typ = 1 Then ' Darstellung als Textzahl in Aufgabe
            form_zahl = form_zahl + "x10<sup>" + Trim(Str(potenz)) + "</sup>"
        Else    ' Darstellung im Moodle format
            form_zahl = form_zahl + "E" + Trim(Str(potenz))
        End If
    End If
    Format_Zahl = form_zahl
End Function
Private Function Zahl_Digits(ByVal zahl As Variant, ByVal digits As Integer) As Variant
Dim potenz As Integer
    potenz = Int(Log_10(zahl))
    If zahl <> 0 Then
        zahl = Round(zahl / (10 ^ potenz), digits - 1) * (10 ^ potenz)
    End If
    Zahl_Digits = zahl
End Function
Private Function Trim_Hyp(ByVal Zahl_Str As String) As String
    If Left(Zahl_Str, 1) = "'" Then
        Zahl_Str = Right(Zahl_Str, Len(Zahl_Str) - 1)
    End If
    Trim_Hyp = Zahl_Str
End Function
Private Function Gen_Result_Tolerance(ByVal Zahl_Str As String, ByVal zahl As Variant, ByVal Err_100 As Variant, ByVal Err_50 As Variant)
Dim Res_str As String
        Res_str = "=" + Zahl_Str + ":" + Format_Zahl(Return_digits(zahl * Err_100, 2), 2, 4, -2)
        If Err_50 > 0 Then
            Res_str = Res_str + "~%50%" + Zahl_Str + ":" + Format_Zahl(Return_digits(zahl * Err_50, 2), 2, 4, -2)
        End If
        Gen_Result_Tolerance = Res_str
End Function
Private Function Log_10(ByVal zahl As Variant) As Variant ' cave: gibt bei zahl = 0 log_10 gleich 0 aus!
    If zahl > 0 Then
        Log_10 = Log(zahl) / Log(10)
    Else
        Log_10 = 0
    End If
End Function
Private Function Cloze_Lueckentext(ByVal Lueckentext As String, ByVal Einfuegen As String, ByVal platzhalter As String) As String
Dim position As Integer
    position = InStr(Lueckentext, "xxx")
    If position > 0 Then
        Cloze_Lueckentext = Left(Lueckentext, position - 1) + Einfuegen + Right(Lueckentext, Len(Lueckentext) - position - 2)
    Else
        Cloze_Lueckentext = Lueckentext
    End If
End Function
Private Function Write_Intro(ByVal HTML_exp As Boolean) As String
Dim Intro_String As String
Dim x As Integer
    With Worksheets("Intro")
        Intro_String = "<h2>" + .Cells(1, 2).Value + "</h2>" + vbLf + "<p>"
        x = start_intro_x
        While .Cells(x, 2).Value <> ""
            If .Cells(x, 1).Value <> "Hidden" Then
                If .Cells(x, 1).Value <> "HTML_Comment" Then
                    Intro_String = Intro_String + Cloze_Lueckentext(.Cells(x, 2).Value, .Cells(x, 3).Value, "xxx") + " "
                Else
                    If HTML_exp Then
                        Intro_String = Intro_String + vbLf + .Cells(x, 2).Value + vbLf
                    End If
                End If
            End If
            x = x + 1
        Wend
    End With
    Write_Intro = Intro_String + "</p>"
End Function
Private Function Write_Solution(ByVal data_num As Integer, ByVal HTML_exp As Boolean) As String
Dim Solution_String As String
Dim x As Integer
Dim table As Boolean
Dim column_no As Integer
Dim col_count As Integer
Dim tab_sep As String
Dim tab_end As String

    x = start_sol_x
    tab_sep = ""
    tab_end = ""
    With Worksheets("Solution")
        Solution_String = "<p>" + .Cells(1, 2).Value + "</p>" + vbLf + "<p>"
        While .Cells(x, 1).Value <> ""
            If .Cells(x, 1).Value <> "Hidden" Then
                If .Cells(x, 1).Value <> "HTML_Comment" Then
                    If (.Cells(x, 1).Value = "TableStart") Or (.Cells(x, 1).Value = "TableEnd") Then
                        table = .Cells(x, 1).Value = "TableStart"
                        If table Then
                            Solution_String = Solution_String + "</p>" + vbLf + "<table border=""1"">" + vbLf + "<tbody>" + vbLf
                            If .Cells(x, Text_y).Value <> "" Then
                                column_no = .Cells(x, Text_y).Value
                            Else
                                column_no = 2
                            End If
                            col_count = 0
                            tab_sep = "</td>" + vbLf + "<td>"
                            tab_end = "</td>" + vbLf
                        Else
                            If col_count < column_no Then
                                While col_count < column_no
                                    Solution_String = Solution_String + "</td>" + vbLf + "<td>"
                                    col_count = col_count + 1
                                Wend
                            End If
                            tab_sep = ""
                            tab_end = ""
                            Solution_String = Solution_String + "</tr>" + vbLf + "</tbody>" + vbLf + "</table>" + vbLf
                        End If
                    Else
                        If table Then
                            If col_count = 0 Then
                                Solution_String = Solution_String + "<tr>" + vbLf + "<td>"
                            Else
                                If col_count >= column_no Then
                                    col_count = 0
                                    Solution_String = Solution_String + "</tr>" + vbLf + "<tr>" + vbLf
                                End If
                                Solution_String = Solution_String + "<td>"
                            End If
                        End If
                        Solution_String = Solution_String + .Cells(x, Text_y).Value + " "
                        If .Cells(x, 1).Value <> "Text" Then
                            If table Then
                                col_count = col_count + 2
                            End If
                            Solution_String = Solution_String + tab_sep + .Cells(x, 4).Value + " " + tab_end
                        Else
                            Solution_String = Solution_String + tab_end
                            col_count = col_count + 1
                        End If
                    End If
                Else
                    If HTML_exp Then
                        Solution_String = Solution_String + vbLf + .Cells(x, 3).Value + vbLf
                    End If
                End If
            End If
            x = x + 1
        Wend
        Solution_String = Solution_String + "</p>"
        If Worksheets("Gen_output").Cells(7, 2).Value Then
            Solution_String = Solution_String + vbLf + "<p><br><br><i>The following field is included for technical reasons and can be ignored.</i> {0:NM:=" + Trim(Str(data_num)) + "} </p>"
        End If
    End With
    Write_Solution = Solution_String
End Function

Private Function XML_Header(number As Integer) As String
    XML_Header = vbLf + "<question type=""cloze"">" + vbLf + "<name><text>" + _
        Worksheets("Gen_Output").Cells(4, 2).Value + " - " + Trim(Str(number)) + "</text></name>" + vbLf + "<questiontext format=""html"">" + vbLf + "<text><![CDATA[" + vbLf
End Function
Private Function XML_End(number As Integer) As String
Dim XML_End_Text As String
Dim n_tags As Integer
Dim int_tag_text As String
Dim x As Integer
    Worksheets("Gen_output").Cells(20, 2).Value = "'" + Trim(Str(number))
    XML_End_Text = vbLf + "]]></text></questiontext>"
    x = 17
'        + vbLf + "<hidden>0</hidden><idnumber>" + Trim(Str(number)) + "</idnumber>" + vbLf + "</question>"
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
Private Sub Clean_HTML()
Dim non_html As Variant
Dim is_html As Variant
Dim n_sub As Integer
Dim i As Integer
' das hier ist sauwichtig - sonst hängt es beim Moodle-import.
    non_html = Array("", "§", "°", "²", "³", "µ", "Ä", "Ö", "Ü", "ä", "ö", "ü", "ß")
    is_html = Array("", "&sect;", "&deg;", "&sup2;", "&sup3;", "&micro;", "&Auml;", "&Ouml;", "&Uuml;", "&auml;", "&ouml;", "&uuml;", "&szlig;")
    n_sub = 3
    For i = 0 To n_sub - 1
        Sheets("Gen_output").Range("B12,B17,B21,B24").Replace what:=non_html(i), replacement:=is_html(i), lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        Sheets("Intro").Range("B14,B:B").Replace what:=non_html(i), replacement:=is_html(i), lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        Worksheets("Solution").Range("B:B,C:C,E:E").Replace what:=non_html(i), replacement:=is_html(i), lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Next i
End Sub
Private Sub Clear_Evaluation()
    Dim x As Integer
    x = 8
    With Worksheets("Evaluation")
        While .Cells(x, 1).Value <> ""
            If .Cells(x, 2).Value <> "sub" Then
                .Cells(x, 5).Value = ""
            Else
                .Cells(x, 5).Value = False
            End If
            x = x + 1
        Wend
    End With
End Sub
Sub Format_Student_Sheet()
'
' Format_Student_Sheet Makro
'
' Tastenkombination: Strg+Umschalt+F
'
Dim question_no As String
Dim solution As String
Dim x, y, ys As Integer
Dim no_responses As Integer
    If Cells(4, 4).Value <> "Points" Then
        Rows("1:3").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Columns("D:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("D4").Value = "Points"
        Range("E4").Value = "Bonus Points"
        Range("F4").Value = "HTML"
        Range("A1").Value = "Result Sheet " + ActiveSheet.Name
        Range("A2").Value = "No Responses:"
        Range("A3").Value = "Scale factor:"
        Range("B3").Value = 1
        ActiveSheet.Name = "Responses"
        question_no = "Response " + (InputBox("Question number?"))
        x = 5
        Cells(4, 3).Value = "ID"
        While Cells(x, 1).Value <> ""
            Cells(x, 3).Value = Left(Cells(x, 3).Value, 5)
            x = x + 1
        Wend
        y = 7
        While Cells(4, y).Value <> ""
            While (Cells(4, y).Value <> "") And (Cells(4, y).Value <> question_no)
                Cells(4, y).EntireColumn.Delete
            Wend
            y = y + 2   ' die korrekte Antwort wird auch stehen gelassen
        Wend
        x = 5
        no_responses = Get_Response_Number(Cells(x, 7).Value)
        Cells(2, 2) = no_responses
        While Cells(x, 1).Value <> ""
            y = 9
            If Right(Cells(x, 7).Value, 1) <> ";" Then
                If Right(Cells(x, 7).Value, 1) = ":" Then
                    Cells(x, 7).Value = Cells(x, 7).Value + " "
                End If
                Cells(x, 7).Value = Cells(x, 7).Value + ";"
            End If
            solution = Cells(x, 7).Value
            While solution <> ""
                Cells(x, y).Value = Extract_Solution(solution)
                y = y + 1
            Wend
            If Right(Cells(x, 8).Value, 1) <> ";" Then
                If Right(Cells(x, 8).Value, 1) = ":" Then
                    Cells(x, 8).Value = Cells(x, 8).Value + " "
                End If
                Cells(x, 8).Value = Cells(x, 8).Value + ";"
            End If
            solution = Cells(x, 8)
            While solution <> ""
                Cells(x, y).Value = Extract_Solution(solution)
                y = y + 1
            Wend
            x = x + 1
        Wend
    End If
End Sub
Private Function Extract_Solution(ByRef solution As String) As String
Dim sol_str As String
Dim pos As Integer
    pos = InStr(1, solution, ":")
    If pos > 0 Then
        solution = Right(solution, Len(solution) - pos - 1)
    End If
    pos = InStr(1, solution, "; part ")    ' cave: in der Lösung darf "; part " nicht auftauchen!!!
    If pos = 0 Then
        pos = Len(solution)
    End If
    If pos > 0 Then
        sol_str = Left(solution, pos - 1)
        solution = Right(solution, Len(solution) - pos)
        If Len(solution) = 1 Then
            solution = ""
        End If
    Else
        sol_str = ""
        solution = ""
    End If
    Extract_Solution = sol_str
End Function
Private Function Get_Response_Number(ByVal response As String) As Integer
Dim pos_old, pos_new, resp_no As Integer
    resp_no = 1
    pos_old = 0
    pos_new = 1
    While pos_old < pos_new
        pos_old = pos_new
        pos_new = InStr(pos_new + 1, response, "; part ")
        resp_no = resp_no + 1
    Wend
    Get_Response_Number = resp_no - 1
End Function
Sub Transfer_Student_Results()
Dim xst, xev As Integer
Dim yst As Integer
Dim number_parts As Integer
Dim y_dataset As Integer
Dim number_dataset As Integer
Dim var_intro_x As Variant
    var_intro_x = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    Call Get_Variables_X(var_intro_x)
    y_dataset = Worksheets("Evaluation").Cells(5, 2).Value
    number_parts = Worksheets("Responses").Cells(2, 2).Value   ' cave: nur Verschiebung - wird zum yst Wert addiert!!!
    xst = 5
    With Worksheets("Responses")
        While .Cells(xst, 1).Value <> ""
            number_dataset = .Cells(xst, y_dataset)
            Call Transfer_Variables(var_intro_x, number_dataset)
            Call Transfer_Student_Data(xst, number_parts)
            Call Write_Student_Data(xst, number_parts)
            xst = xst + 1
        Wend
    End With
    
End Sub
Sub Transfer_Dataset()
' Makro überträgt einzelnen Datensatz
Dim dataset As Integer
Dim var_x_init As Variant
    dataset = Val(InputBox("Dataset Number?", "Dataset", "1"))
    var_x_init = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)    ' hier stehen die zu übertragenden Variablen
    Call Get_Variables_X(var_x_init)
    Call Transfer_Variables(var_x_init, dataset)
End Sub
Sub Mark_Correct(ByVal x_st As Integer, y_shift As Integer)
Dim y As Integer
Dim x_sol As Integer
Dim bewertung As String
Dim err_tol As Variant
    x_sol = 8
    err_tol = Worksheets("Gen_output").Cells(11, 2).Value
    With Worksheets("Evaluation")
        While .Cells(x_sol, 1).Value <> ""
            If .Cells(x_sol, 3).Value <> "" Then
                y = .Cells(x_sol, 3).Value
                If .Cells(x_sol, 7).Value = .Cells(x_sol, 8).Value Then
                    bewertung = "Gut"
                Else
                    If .Cells(x_sol, 8).Value > 0 Then
                        bewertung = "Neutral"
                    Else
                        bewertung = "Schlecht"
                    End If
                End If
                Worksheets("Responses").Cells(x_st, y).Style = bewertung
                
            End If
            If .Cells(x_sol, 4).Value <> "" Then
                If Abs(1 - (Worksheets("Responses").Cells(x_st, y).Value / Worksheets("Responses").Cells(x_st, y + y_shift).Value)) <= err_tol Then
                    Worksheets("Responses").Cells(x_st, y).Font.Bold = True
                    Worksheets("Responses").Cells(x_st, y).Font.Underline = xlUnderlineStyleSingle
                Else
                    Worksheets("Responses").Cells(x_st, y).Font.Bold = False
                    Worksheets("Responses").Cells(x_st, y).Font.Underline = xlUnderlineStyleNone
                End If
            End If
            x_sol = x_sol + 1
        Wend
    End With
    
End Sub
Private Sub Transfer_Student_Data(ByVal xst As Integer, ByVal number_parts As Integer)
Dim xev As Integer
Dim yst As Integer
    xev = 8
    With Worksheets("Responses")
        While Worksheets("Evaluation").Cells(xev, 1).Value <> ""
            yst = Worksheets("Evaluation").Cells(xev, 3).Value
            If yst > 0 Then
                Select Case Worksheets("evaluation").Cells(xev, 2).Value
                Case "NM"
                    Worksheets("Evaluation").Cells(xev, 5).Value = .Cells(xst, yst).Value
                Case "sub"
                    Worksheets("Evaluation").Cells(xev, 5).Value = (.Cells(xst, yst).Value = "Yes")
                Case Else
                    Worksheets("Evaluation").Cells(xev, 5).Value = (.Cells(xst, yst).Value = .Cells(xst, yst + number_parts).Value)
                End Select
            End If
            xev = xev + 1
        Wend
    End With
End Sub
Private Sub Write_Student_Data(ByVal xst As Integer, ByVal y_shift As Integer)
    With Worksheets("Responses")
            .Cells(xst, 4).Value = Worksheets("Evaluation").Cells(6, 8).Value
            .Cells(xst, 5).Value = Worksheets("Evaluation").Cells(5, 8).Value
            .Cells(xst, 6).Value = Worksheets("Evaluation").Cells(4, 2).Value
            Call Mark_Correct(xst, y_shift)
    End With
End Sub
Sub Transfer_Single_Student()
' Makro überträgt einzelnen Studierenden
Dim xst As Integer
Dim dataset As Integer
Dim studi_name As String
Dim var_x_init As Variant
Dim number_parts As Integer
    xst = Val(InputBox("Student Line Number?", "Student X", "5"))
    studi_name = Worksheets("Responses").Cells(xst, 1).Value + ", " + Worksheets("Responses").Cells(xst, 2).Value
    If MsgBox("Transfer data from student " + studi_name + "? ", vbYesNo) = vbNo Then
        Exit Sub
    End If
    dataset = Worksheets("Responses").Cells(xst, Worksheets("Evaluation").Cells(5, 2).Value).Value
    number_parts = Worksheets("Responses").Cells(2, 2).Value
    var_x_init = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)    ' hier stehen die zu übertragenden Variablen
    Call Get_Variables_X(var_x_init)
    Call Transfer_Variables(var_x_init, dataset)
    Call Transfer_Student_Data(xst, number_parts)
End Sub
Sub Write_Single_Student()
' Makro schreibt Ergebnisse eines einzelnen Studierenden (cave: ohne Nachfrage!!! - Enthält alle per Hand geänderten Einträge)
Dim xst As Integer
Dim studi_name As String
    xst = Val(InputBox("Student Line Number?", "Student X", "5"))
    studi_name = Worksheets("Responses").Cells(xst, 1).Value + ", " + Worksheets("Responses").Cells(xst, 2).Value
    If MsgBox("Write data from student " + studi_name + "? ", vbYesNo) = vbNo Then
        Exit Sub
    End If
    Call Write_Student_Data(xst, Worksheets("Responses").Cells(2, 2).Value)
End Sub
Private Sub Print_Replace(ByVal Text_File As Integer, ByVal Repl As Integer, ByVal print_text As String)
Dim start_Zeichen As Long
Dim laenge As Integer
Dim ant_nr As Integer
Dim antwort As String
Dim platzhalter As String
' Funktion schreibt in Datei, falls replace aktiviert werden Imagelinks / Beschriftungen angepasst
    If Repl > 0 Then
        start_Zeichen = 1
        start_Zeichen = InStr(start_Zeichen, print_text, "##")
        While start_Zeichen > 0
            laenge = InStr(start_Zeichen, print_text, ";") - start_Zeichen - 2
            platzhalter = Right(Left(print_text, start_Zeichen + 1 + laenge), laenge)
            If IsNumeric(platzhalter) Then
                ant_nr = CInt(platzhalter) + Anz_Link_Fields
            Else
                ant_nr = get_ant_nr(platzhalter)
                If ant_nr = 0 Then
                    MsgBox ("Replace Term " + platzhalter + " not found! Replacement aborted!")
                    Exit Sub
                End If
            End If
            antwort = Worksheets("Rnd_Matrix").Cells(1 + ant_nr, 1 + Repl).Value
            print_text = Replace(print_text, "##" + platzhalter + ";", antwort, 1, 1)
            
            start_Zeichen = InStr(start_Zeichen, print_text, "##")
        Wend
    End If
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
