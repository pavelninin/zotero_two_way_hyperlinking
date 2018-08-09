Public Const opt_SCRIPT_VERSION = 18
Public Const opt_DisplayUpdatingOptionsDisable = 1  ' for performance acceleration
Public Const opt_detailed_ProgressView = 0          ' print progress/statistics info during execution
Public Const opt_wait_ZoteroRefresh_Sec = 3         ' find for your document experimentally
Public Const opt_performBackwardLinking = 1         ' keep it optional since for some documents Backward Linking is not needed
Public Const opt_use_wdUnderlineNone = 1            ' by default in word references are Underlined
Public Const opt_addRefScreenTipsText = 0           ' add reftips => short ref occurence like "text [ref] text"
Public Const opt_addRefScreenTipsTextLength = 50    ' lenght of text "text text text [ref] text text text
Public Const opt_work_with_ZoteroOffline = 1        ' when all links are added => can close zotero to prevent its unexpected backgroud jobs (updating bib, etc)
Public Const opt_FindNumPageSect = 2                ' 0 - pages only (done faster), 1 - sections only (done slower), 2 - pages & sections only (done slower)
Public Const opt_FindNumPattern_00 = "Referenced in sections: [sect_nb] (page [page_nb])" ' this to to define your custom styles
Public Const opt_FindNumPattern_01 = "Referenced in pages: [page_nb]"
Public Const opt_FindNumPattern_00_one = "Referenced in section: [sect_nb] (page [page_nb])"
Public Const opt_FindNumPattern_01_one = "Referenced in page: [page_nb]"
Public Const opt_Referenced_on_page = "Referenced on page"
Public Const opt_Referenced_in_section = "Referenced in section"
Public Const opt_debug_Mode = 0                     ' for debug
Public Const opt_debug_Mode_iter_Nb_Total = 40      ' for debug
Public Const opt_updateProgressPeriod = 2           ' for debug
Public Const opt_extraPagesBeforePageOne = 11       ' how many extra pages (title page, etc) before page no 1
Public Const opt_use_UI = 1                         ' run script with UI
Public Const opt_proc_code = 0
Public Const opt_skip_code = 1
Public Const opt_ZoteroHyperlinkPrefixCommon = "ZOT637"
Public Const opt_PfxForw = opt_ZoteroHyperlinkPrefixCommon + "F"
Public Const opt_PfxBack = opt_ZoteroHyperlinkPrefixCommon + "B"
Public Const opt_err_Max = 4096        ' for debug
Public Const opt_wrn_Max = 4096        ' for debug
Public Const opt_group_Length = 1024
Public Const opt_info_Length = 10
Public Const opt_offset_BibNames = 0 + opt_info_Length                                                      ' used to keep some tech info
Public Const opt_offset_BibRanges = opt_offset_BibNames + opt_group_Length                                  ' used to make bookmark
Public Const opt_offset_BackRefPageNumbers = opt_offset_BibRanges + opt_group_Length                        ' stores pages occurence aka 1,3,5,6
Public Const opt_offset_BackRefPageNumbers_AnchorBackward_to_RefGroup = opt_offset_BackRefPageNumbers + opt_group_Length    ' stores bookmark for each page occurence need to make hypelink to this occurence
Public Const opt_offset_BackRefSectNumbers = opt_offset_BackRefPageNumbers_AnchorBackward_to_RefGroup + opt_group_Length                ' stores sections occurence
Public Const opt_offset_BackRefSectNumbers_AnchorBackward_to_RefGroup = opt_offset_BackRefSectNumbers + opt_group_Length                ' stores bookmark for each section occurence need to make hypelink to this occurence
Public Const opt_offset_RefTip = opt_offset_BackRefSectNumbers_AnchorBackward_to_RefGroup + opt_group_Length ' stores tips like "text [ref] text" for each ref
' NOTE: here PageNumbers_AnchorBackward_to_RefGroup == RefSectNumbersBookMarks to make format "referenced in sections: 3 (page 12)" => click
'       on "3" and on "12" points to the same reference in text "text |reference_group_bookmark|[1,5,7] text"
' NOTE: of course, possible to make other format, where click on page will point to page number in the doc and click
'       on ref num will point to reference_group_bookmark (if exect no is not present aka '6' in 1-7) or exact reference
'       number if present in reference_group_bookmark
Public Const opt_bibEntrMax_Num = opt_offset_RefTip + opt_group_Length                           ' total storage length
Public Const string_Zotero_Bibliography = "Zotero_Bibliography"
Public Const string_ADDIN_ZOTERO_ITEM = "ADDIN ZOTERO_ITEM"
Public g_err_Array(opt_err_Max) As String   ' errors storage
Public g_wrn_Array(opt_wrn_Max) As String   ' warnings storage
Public g_err_num As Integer
Public g_wrn_num As Integer
Public g_bibPositionInDocument As Long
Public g_iter_Nb As Integer
Public g_proc_Nb As Integer
Public g_found_aFields As Integer
Public g_PageNbOffst As Integer
Public g_BibEntries(opt_bibEntrMax_Num) As String
Public g_BibEntriesLen As Integer
Public g_foundBib_Entries As Integer
Public g_zotero_did_FastRefresh As Integer
Public g_zotero_did_LongRefresh  As Integer
Public g_BibEntriesMoved As Integer
Public g_BibEntriesFile_Found  As Integer
Public g_rebuild_BibEntries  As Integer
Public g_UI_init As Integer
Function TOOLS__INIT_GLOBAL_VARIABLES()
    g_PageNbOffst = -1
    g_err_num = 0
    g_wrn_num = 0
    g_iter_Nb = 0
    g_proc_Nb = 0
    g_found_aFields = 0
    Erase g_err_Array
    Erase g_wrn_Array
End Function
Public Sub PERFORM_TWO_WAY_HYPERLINKING()
    Call AA__ENTRY_POINT
End Sub
Function AA__ENTRY_POINT()
    If opt_DisplayUpdatingOptionsDisable Then
        Application.ScreenUpdating = False
        Application.DisplayStatusBar = False
        Application.Visible = False
    End If
    If opt_use_UI Then
        g_UI_init = 0
        Call UI_PROGRESS_BAR.Show(vbModeless)
    Else
        Call BB__PERFORM_TWO_WAY_LINKING
    End If
End Function
Function ZZ_INTERNAL__UI_PROGRESS_BAR_CALLBACK()
    Call BB__PERFORM_TWO_WAY_LINKING
End Function
Function BB__PERFORM_TWO_WAY_LINKING()
    Dim TotaElapsedTimeMin As String
    Debug.Print vbCrLf & "BB__PERFORM_TWO_WAY_LINKING version " & opt_SCRIPT_VERSION
    '===================================================
    Call TOOLS__INIT_GLOBAL_VARIABLES
    ' build references collection in format [title|nb]
    '===================================================
    Call TOOLS__BUILD_BIBLIOGRAPHY_ENTRIES
    g_BibEntriesLen = CInt(g_BibEntries(0))
    Dim StartTime As Double
    Dim TotaElapsedTimeSec As Double
    '===================================================
    StartTime = Timer
    '===================================================
    Call TOOLS__BUILD_FORWARD_REFERENCES
    Debug.Print vbCrLf & "Main loop done " & vbCrLf
    '===================================================
    If opt_performBackwardLinking = 1 Then
        Call TOOLS__BUILD_BACKWARD_REFERENCES(g_BibEntries)
    End If
    '===================================================
    TotaElapsedTimeMin = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    '===================================================
    If ActiveDocument.Saved = False Then ActiveDocument.Save
    '===================================================
    Call TOOLS__SHOW_ERRORS_WARNINGS(TotaElapsedTimeMin)
    '===================================================
    If opt_DisplayUpdatingOptionsDisable Then
        Application.ScreenUpdating = True
        Application.DisplayStatusBar = True
        Application.Visible = True
    End If
End Function
Function TOOLS__BUILD_FORWARD_REFERENCES()
    Dim n1&, n2&
    'Dim iCount As Integer
    Dim fieldCode As String
    Dim plain_Cit As String
    ' loop through each field in the document
    g_found_aFields = 0
    For Each aField In ActiveDocument.Fields
        ' check if the field is a Zotero in-text reference
        If InStr(aField.Code, string_ADDIN_ZOTERO_ITEM) > 0 Then
            g_found_aFields = g_found_aFields + 1
        End If
    Next aField ' next field
    If opt_use_UI Then
        Call UI_PROGRESS_BAR.CONFIGURE(g_found_aFields, 3)
    End If
    For Each aField In ActiveDocument.Fields
    ' check if the field is a Zotero in-text reference
        If InStr(aField.Code, string_ADDIN_ZOTERO_ITEM) > 0 Then
            fieldCode = aField.Code
            inc = 1
            tit_Found = 0
            plain_Cit_Found = 0
            skip_iteration = 0
            Dim tit_Arr(128) As String
            Dim plain_CitNb_Array(128) As String
            Dim plCitStrBeg, plCitStrEnd As String
            plCitStrBeg = """plainCitation"":""["
            plCitStrEnd = "]"""
            IterStartTime = Timer
            ' ....................... start iteration .........................
            n1 = InStr(fieldCode, plCitStrBeg)
            If n1 = 0 Then
                g_err_Array(g_err_num) = "plainCitation begin not found ([ref] format expected)"
                g_err_num = g_err_num + 1
                skip_iteration = 1
            Else
                n1 = n1 + Len(plCitStrBeg)
            End If
            If skip_iteration = 1 Then
                skip_iteration = 1
            Else
                n2 = InStr(Mid(fieldCode, n1, Len(fieldCode) - n1), plCitStrEnd) - 1 + n1
                plain_Cit = Mid$(fieldCode, n1, n2 - n1)
                ret_code = TOOLS__PROCESS_REFERENCE(plain_Cit, aField, g_BibEntries)
                If ret_code = opt_proc_code Then
                    g_proc_Nb = g_proc_Nb + 1
                ElseIf ret_code = opt_skip_code Then
                    ' nothing for the moment
                End If
            End If
            IterElapsedTimeSec = Round(Timer - IterStartTime, 2)
            IterElapsedTimeMin = Format((Timer - IterStartTime) / 86400, "hh:mm:ss")
            TotaElapsedTimeSec = Round(Timer - StartTime, 2)
            TotaElapsedTimeMin = Format((Timer - StartTime) / 86400, "hh:mm:ss")
            g_iter_Nb = g_iter_Nb + 1
            If opt_use_UI Then
                Call UI_PROGRESS_BAR.SET_PROGRESS(g_iter_Nb)
            End If
            If g_iter_Nb - (opt_updateProgressPeriod * (g_iter_Nb \ opt_updateProgressPeriod)) = 0 Then
                If opt_detailed_ProgressView Then
                    Debug.Print "Iterations done : " & g_iter_Nb & " / " & g_found_aFields & " (" & Round(g_iter_Nb * 100 / g_found_aFields) & "%) " _
                                                    ; "(" & g_proc_Nb & " processed " & g_iter_Nb - g_proc_Nb & " skipped) " & vbCrLf & _
                                "Time elapsed cur: " & IterElapsedTimeSec & " seconds (" & IterElapsedTimeMin & " minutes) "; vbCrLf & _
                                "Time elapsed tot: " & TotaElapsedTimeSec & " seconds (" & TotaElapsedTimeMin & " minutes)" & vbCrLf & _
                                "Errors/warnings : " & g_err_num & " / " & g_wrn_num & vbCrLf
                Else
                    Debug.Print "Progress " & Round(g_iter_Nb * 100 / g_found_aFields) & "% time elapsed " & _
                                TotaElapsedTimeMin & " errors "; g_err_num & " warnings " & g_wrn_num
                End If
            End If
        End If
    Next aField ' next field
End Function
Function TOOLS__SHOW_ERRORS_WARNINGS(TotaElapsedTimeMin As String)
    If g_err_num = 0 And g_wrn_num = 0 Then
        MsgBox "Debug mode." & vbCrLf & _
                "Script finished." & vbCrLf & "NO errors NO warnings" & vbCrLf & _
               "Iterations  : total " & g_found_aFields & " executed " & g_iter_Nb & " (" & g_proc_Nb & " processed " & g_iter_Nb - g_proc_Nb & " skipped) " & vbCrLf & _
               "Time elapsed: " & TotaElapsedTimeSec & " seconds (" & TotaElapsedTimeMin & " minutes)"
    Else
        Dim err_str, wrn_str  As String
        If g_err_num > 0 Then
            err_limit = 10
            err_limited = 0
            If g_err_num > err_limit Then
                g_err_num = err_limit
                err_limited = 1
            End If
            err_str = "Errors:" + vbCrLf
            If err_limited = 1 Then
                err_str = err_str + "too many errors (limited to " + CStr(err_limit) + ")" + vbCrLf
            End If
            For ki = 0 To g_err_num - 1 Step 1
                err_str = err_str + CStr(ki) + " " + g_err_Array(ki) + vbCrLf
            Next ki
        Else
            err_str = ""
        End If
        If g_wrn_num > 0 Then
            wrn_limit = 10
            wrn_limited = 0
            If g_wrn_num > wrn_limit Then
                g_wrn_num = wrn_limit
                wrn_limited = 1
            End If
            wrn_str = "Warnings:" + vbCrLf
            If err_limited = 1 Then
                wrn_str = wrn_str + "too many warnings (limited to " + CStr(wrn_limit) + ")" + vbCrLf
            End If
            For ki = 0 To g_wrn_num - 1 Step 1
                wrn_str = wrn_str + CStr(ki) + " " + g_wrn_Array(ki) + vbCrLf
            Next ki
        Else
            wrn_str = ""
        End If
        MsgBox "Debug mode." & vbCrLf & _
               "Script finished." & vbCrLf & g_err_num & " errors " & g_wrn_num & " warnings" & vbCrLf & _
               "Iterations  : total " & g_found_aFields & " executed " & g_iter_Nb & " (" & g_proc_Nb & " processed " & g_iter_Nb - g_proc_Nb & " skipped) " & vbCrLf & _
               "Iterations  : total " & g_found_aFields & " executed " & g_iter_Nb & " (" & g_proc_Nb & " processed " & g_iter_Nb - g_proc_Nb & " skipped) " & vbCrLf & _
               "Time elapsed: " & TotaElapsedTimeSec & " seconds (" & TotaElapsedTimeMin & " minutes)" & vbCrLf & _
               err_str & wrn_str
    End If
End Function
Function TOOLS__PROCESS_REFERENCE(plain_Cit As String, aField As Variant, g_BibEntries As Variant) As Integer
    Dim plain_CitNb_Array(128) As String
    Dim page_nb_ext As String
    Dim sect_nb_ext As String
    Dim page_nb_ext_i As Integer
    plain_Cit_Found = 0
    refsNum = 0
    MList = 0
    TList = 0
    PList = 0
    SimpC = 0
    If InStr(plain_Cit, ",") > 0 And (InStr(plain_Cit, "–") > 0 Or InStr(plain_Cit, "-") > 0) Then
        ' mixed list aka 1,3,5-7
        MList = 1
    ElseIf InStr(plain_Cit, "–") > 0 Or InStr(plain_Cit, "-") > 0 Then
        ' trait separated list aka 1-3
        Dim Tags(128) As String
        tagsFound = 0
        pch = TOOLS__STRTOK(plain_Cit, "-–")
        Do While pch <> vbNullString
            Tags(tagsFound) = pch
            tagsFound = tagsFound + 1
            pch = TOOLS__STRTOK(vbNullString, "-–")
        Loop
        TListS = CInt(Tags(0))
        TListE = CInt(Tags(1))
        TList = 1
        refsNum = 2
    ElseIf InStr(plain_Cit, ",") > 0 Then
        pch = TOOLS__STRTOK(plain_Cit, ",")
        Do While pch <> vbNullString
            plain_CitNb_Array(plain_Cit_Found) = pch
            plain_Cit_Found = plain_Cit_Found + 1
            pch = TOOLS__STRTOK(vbNullString, " ,")
        Loop
        PList = 1
        refsNum = plain_Cit_Found
    Else
        SimpC = 1
        refsNum = 1
    End If
    If MList Then
        g_wrn_Array(g_wrn_num) = "mixed list " & plain_Cit & " skipped (not yet supported)"
        g_wrn_num = g_wrn_num + 1
        TOOLS__PROCESS_REFERENCE = opt_skip_code
        Exit Function
    End If
    ' .................... process forward part ...............
    aField.Select
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "[0-9]{1,}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    For iCount = 1 To refsNum Step 1
        Selection.Find.Execute
        If Selection.Find.Found = True Then
            refNo = Selection.Range.Text
            If TList Then
                If (iCount = 1 And refNo <> TListS) Or _
                   (iCount = 2 And refNo <> TListE) Then
                   ' must match => pb
                End If
            ElseIf PList Then
                If iCount <> plain_CitNb_Array(iCount - 1) Then
                   ' must match => pb
                End If
            ElseIf SimpC Then
                ' nothing to check
            End If
            Dim titleName As String:
            titleName = g_BibEntries(opt_offset_BibNames + refNo - 1)
            titleAnchorForward = Left(opt_PfxForw + TOOLS__HASH12(titleName), 40)
            ' store current style
            style = Selection.style
            ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, _
                address:="", SubAddress:=titleAnchorForward, _
                ScreenTip:=Left(titleName, 1024), TextToDisplay:="" & refNo
        End If
    Next iCount
    If opt_use_wdUnderlineNone Then
        ' uncomment to force black font/no underline
        aField.Select
        With Selection.Font
            .Underline = wdUnderlineNone
            '.ColorIndex = wdBlack
        End With
    End If
    ' .................... process backward part ....................
    If opt_performBackwardLinking = 1 Then
        'aField.Select
        find_p = 0
        find_s = 0
        ' 0 - pages only (done faster), 1 - sections only (done slower), 2 - pages & sections only (done slower)
        If opt_FindNumPageSect = 0 Then find_p = 1: find_s = 0
        If opt_FindNumPageSect = 1 Then find_p = 0: find_s = 1
        If opt_FindNumPageSect = 2 Then find_p = 1: find_s = 1
        If find_p Then
            page_nb_ext_i = Selection.Information(wdActiveEndPageNumber) - opt_extraPagesBeforePageOne
            If page_nb_ext_i < 1 Then
                page_nb_ext_i = 1
            End If
            page_nb_ext = CStr(page_nb_ext_i)
        End If
        If find_s Then
            sect_nb_ext = ""
            ' find link lication in section
            aField.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            search_in_sect = Selection.Information(wdActiveEndSectionNumber)
            Do While sect_nb_found = 0
                Selection.MoveUp Unit:=wdParagraph, Count:=1
                cur_sect = Selection.Information(wdActiveEndSectionNumber)
                sel_text = Selection.Range.Text
                list_str = Selection.Paragraphs(1).Range.ListFormat.ListString
                name_loc = Selection.style.NameLocal
                ' FIXME differnet stop policies based on reaching list nb, based on style, etc
               
                If 0 Then
                    'If name_loc = "Titre 1 - Chapter Name" Then
                    If name_loc = "Titre 1 - List No Hidden" Then
                        With Selection.Find
                            ' https://wordmvp.com/FAQs/General/UsingWildcards.htm
                            .Text = "Chapter [0-9]{1,}"
                            '.Text = "^#. ^t"
                            '.Text = titleName
                            .Replacement.Text = ""
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .MatchCase = False
                            .MatchWholeWord = False
                            .MatchWildcards = True
                            .MatchSoundsLike = False
                            .MatchAllWordForms = False
                        End With
                        Selection.Find.Execute
                        If Selection.Find.Found = True Then
                            sect_nb_ext = Selection.Range.Text
                            Exit Do 'sect_nb_ext found => normal exit
                        Else
                            sect_nb_ext = "not found (err 03)"
                            g_wrn_Array(g_wrn_num) = "Titre 1 - Chapter Name is found, but pattern is not (err 02) for ref " & refNo
                            g_wrn_num = g_wrn_num + 1
                            Exit Do
                        End If
                    End If
                End If
                If 1 Then
                    If list_str <> "" Then
                        lin_ind = Selection.Paragraphs(1).LeftIndent
                        fli_ind = Selection.Paragraphs(1).FirstLineIndent
                        If name_loc = "Titre 1,List No Hidden" Then
                            sect_nb_ext = list_str
                            Exit Do 'sect_nb_ext found => normal exit
                        ElseIf Abs(lin_ind) - Abs(fli_ind) = 0 Then
                            sect_nb_ext = list_str
                            Exit Do 'sect_nb_ext found => normal exit
                        End If
                    End If
                End If
                prev_sect_reached = 0
                first_line_reached = 0
                ' disable first_line_reached because on 1st line of page 2 of section 1 cur_line_nb
                ' returns 1 => so it is relative to page, not to section/document
                'cur_line_nb = Selection.Range.Information(wdFirstCharacterLineNumber)
                'If cur_line_nb = 1 Then
                '    first_line_reached = 1
                '    wrn_Array(wrn_num) = "first_line_reached during section search (did you forget to layout your doc?)"
                '    wrn_num = wrn_num + 1
                'End If
                If cur_sect < search_in_sect Then
                    prev_sect_reached = 1
                    g_wrn_Array(g_wrn_num) = "prev_sect_reached during section search (search went to previous section, unexpected)"
                    g_wrn_num = g_wrn_num + 1
                End If
                If first_line_reached = 1 Then
                    sect_nb_ext = "not found (err 01)"
                    g_wrn_Array(g_wrn_num) = "first_line_reached (err 01) for ref " & refNo
                    g_wrn_num = g_wrn_num + 1
                    Exit Do
                End If
                If prev_sect_reached = 1 Then
                    sect_nb_ext = "not found (err 02)"
                    g_wrn_Array(g_wrn_num) = "prev_sect_reached (err 02) for ref " & refNo
                    g_wrn_num = g_wrn_num + 1
                    Exit Do
                End If
            Loop
        End If
        beg_idx = 0
        end_idx = 0
        ' ....................... add bookmark to this reference_group (reference_group_bookmark) .......................
        aField.Select
        unique_reference_group_id = Selection.Information(wdActiveEndPageNumber) * Selection.Range.Start
        titleName = plain_Cit   ' must be same as if you find it with aField.Select => Selection.Find => plain_Cit = Selection.Range.Text
        titleAnchorBackward_to_RefGroup = Left(opt_PfxBack + TOOLS__HASH12(titleName + CStr(unique_reference_group_id)), 40)
        With ActiveDocument.Bookmarks
            .Add Range:=Selection.Range, name:=titleAnchorBackward_to_RefGroup
            .DefaultSorting = wdSortByName
            .ShowHidden = True
        End With
        ' ....................... loop through all reference_group members .......................
        If TList Then beg_idx = TListS: end_idx = TListE
        If PList Then beg_idx = 0: end_idx = plain_Cit_Found - 1
        If SimpC Then beg_idx = 0: end_idx = 0
        For iCount = beg_idx To end_idx Step 1
            relative_offset = 0
            If TList Then relative_offset = iCount
            If PList Then relative_offset = plain_CitNb_Array(iCount)
            If SimpC Then relative_offset = refNo
            If find_p Then
                ' add BackRefPageNumbers: make Referenced on page: XXX
                g_BibEntries(opt_offset_BackRefPageNumbers - 1 + relative_offset) = _
                g_BibEntries(opt_offset_BackRefPageNumbers - 1 + relative_offset) + page_nb_ext + ","
            End If
            If find_s Then
                ' add BackRefSectNumbers: make Referenced in sections: XXX
                ' FIXME: not yet implemented => search if section number for reference group takes lot of time => need to use para-by-para scanning
                g_BibEntries(opt_offset_BackRefSectNumbers - 1 + relative_offset) = _
                g_BibEntries(opt_offset_BackRefSectNumbers - 1 + relative_offset) + sect_nb_ext + ","
            End If
            ' add PageNumbers_AnchorBackward_to_RefGroup: to click on page number and go to RefGroup
            g_BibEntries(opt_offset_BackRefPageNumbers_AnchorBackward_to_RefGroup - 1 + relative_offset) = _
            g_BibEntries(opt_offset_BackRefPageNumbers_AnchorBackward_to_RefGroup - 1 + relative_offset) + titleAnchorBackward_to_RefGroup + ","
            ' add SectNumbers_AnchorBackward_to_RefGroup: to click on sect number and go to RefGroup
            ' NOTE: of course, possible to make other format, where click on page will point to page number in the doc and click
            '       on ref num will point to reference_group_bookmark (if exect no is not present aka '6' in 1-7) or exact reference
            '       number if present in reference_group_bookmark
            g_BibEntries(opt_offset_BackRefSectNumbers_AnchorBackward_to_RefGroup - 1 + relative_offset) = _
            g_BibEntries(opt_offset_BackRefSectNumbers_AnchorBackward_to_RefGroup - 1 + relative_offset) + titleAnchorBackward_to_RefGroup + ","
            Dim RefTip As String
            If opt_addRefScreenTipsText Then ' add reftips => short ref occurence like "text [ref] text
                ref_beg = Selection.Range.Start
                ref_end = Selection.Range.End
                halpTipLen = opt_addRefScreenTipsTextLength / 2
                ' FIXME hope that new range is valid
                ' FIXME need to check border cases:
                ' - text [ref] text text text text text text
                ' - text text text text text text [ref].
                Selection.SetRange Start:=ref_beg - halpTipLen, End:=ref_end + halpTipLen
                RefTip = Selection.Range.Text
                Selection.SetRange Start:=ref_beg, End:=ref_end
            Else
                If find_p Then RefTip = CStr(page_nb_ext)
                If find_s Then RefTip = CStr(sect_nb_ext)
            End If
            ' add reftips => short ref occurence like "text [ref] text
            g_BibEntries(opt_offset_RefTip - 1 + relative_offset) = _
            g_BibEntries(opt_offset_RefTip - 1 + relative_offset) + RefTip + ","
        Next iCount
    End If
    TOOLS__PROCESS_REFERENCE = opt_proc_code
End Function
Function TOOLS__STRTOK(strVar As String, delims As String) As String
      Static stmp As String
      Dim i As Long
      TOOLS__STRTOK = vbNullString
      'Initialize first time calling
      If strVar <> vbNullString Then
        stmp = strVar
      End If
      'Nothing left to tokenize!
      If stmp = vbNullString Then
        Exit Function
      End If
search_for_next_delimiter:
      'Loop until we find a delimiter
      For i = 1 To Len(stmp)
        If InStr(1, delims, Mid$(stmp, i, 1), vbBinaryCompare) > 0 Then
          If i > 1 Then
            TOOLS__STRTOK = Left(stmp, i - 1)
            stmp = Mid$(stmp, i + 1, Len(stmp) - i)
            Exit Function
          Else
            ' string starts with delimiter, so skip one
            stmp = Right$(stmp, Len(stmp) - 1)
            GoTo search_for_next_delimiter
          End If
        End If
      Next i
      'Did not find any, return whatever is left in stmp
      TOOLS__STRTOK = stmp
      stmp = vbNullString
End Function
Function TOOLS__HASH12(s As String)
    ' create a 12 character hash from string s
    Dim l As Integer, l3 As Integer
    Dim s1 As String, s2 As String, s3 As String
    l = Len(s)
    l3 = Int(l / 3)
    s1 = Mid(s, 1, l3)      ' first part
    s2 = Mid(s, l3 + 1, l3) ' middle part
    s3 = Mid(s, 2 * l3 + 1) ' the rest of the string...
    TOOLS__HASH12 = TOOLS__HASH4(s1) + TOOLS__HASH4(s2) + TOOLS__HASH4(s3)
End Function
Function TOOLS__HASH4(txt)
    ' copied from the example
    Dim x As Long
    Dim mask, i, j, nC, crc As Integer
    Dim c As String
    crc = &HFFFF
    For nC = 1 To Len(txt)
        j = Asc(Mid(txt, nC)) ' <<<<<<< new line of code - makes all the difference
        ' instead of j = Val("&H" + Mid(txt, nC, 2))
        crc = crc Xor j
        For j = 1 To 8
            mask = 0
            If crc / 2 <> Int(crc / 2) Then mask = &HA001
            crc = Int(crc / 2) And &H7FFF: crc = crc Xor mask
        Next j
    Next nC
    c = Hex$(crc)
    ' <<<<< new section: make sure returned string is always 4 characters long >>>>>
    ' pad to always have length 4:
    While Len(c) < 4
      c = "0" & c
    Wend
    TOOLS__HASH4 = c
End Function
Function TOOLS__BUILD_BACKWARD_REFERENCES_ON_ENTRY(ByVal refNo As String, ByVal refIn As String, ByVal bibText As String, ByVal title As String, g_BibEntries As Variant)
    Dim BackRefPageNumbers As String
    Dim BackRefSectNumbers As String
    Dim BackRefPageNumbers_AnchorBackward_to_RefGroup As String
    Dim BackRefSectNumbers_AnchorBackward_to_RefGroup As String
    Dim RefTip As String
    Dim str_ext As String
    Dim str_ext2 As String
    relative_offset = CInt(refNo)
    find_p = 0
    find_s = 0
    ' 0 - pages only (done faster), 1 - sections only (done slower), 2 - pages & sections only (done slower)
    If opt_FindNumPageSect = 0 Then find_p = 1: find_s = 0
    If opt_FindNumPageSect = 1 Then find_p = 0: find_s = 1
    If opt_FindNumPageSect = 2 Then find_p = 1: find_s = 1
    If find_p Then
        ' add BackRefPageNumbers: make Referenced on page: XXX
        BackRefPageNumbers = g_BibEntries(opt_offset_BackRefPageNumbers - 1 + relative_offset)
        BackRefPageNumbers = Mid(BackRefPageNumbers, 1, Len(BackRefPageNumbers) - 1) ' remove last comma
    End If
    If find_s Then
        ' add BackRefSectNumbers: make Referenced in sections: XXX
        ' FIXME: not yet implemented => search if section number for reference group takes lot of time => need to use para-by-para scanning
        BackRefSectNumbers = g_BibEntries(opt_offset_BackRefSectNumbers - 1 + relative_offset)
        BackRefSectNumbers = Mid(BackRefSectNumbers, 1, Len(BackRefSectNumbers) - 1) ' remove last comma
    End If
    ' add PageNumbers_AnchorBackward_to_RefGroup: to click on page number and go to RefGroup
    BackRefPageNumbers_AnchorBackward_to_RefGroup = g_BibEntries(opt_offset_BackRefPageNumbers_AnchorBackward_to_RefGroup - 1 + relative_offset)
    BackRefPageNumbers_AnchorBackward_to_RefGroup = Mid(BackRefPageNumbers_AnchorBackward_to_RefGroup, 1, Len(BackRefPageNumbers_AnchorBackward_to_RefGroup) - 1) ' remove last comma
    ' add SectNumbers_AnchorBackward_to_RefGroup: to click on sect number and go to RefGroup
    ' NOTE: of course, possible to make other format, where click on page will point to page number in the doc and click
    '       on ref num will point to reference_group_bookmark (if exect no is not present aka '6' in 1-7) or exact reference
    '       number if present in reference_group_bookmark
    BackRefSectNumbers_AnchorBackward_to_RefGroup = g_BibEntries(opt_offset_BackRefSectNumbers_AnchorBackward_to_RefGroup - 1 + relative_offset)
    BackRefSectNumbers_AnchorBackward_to_RefGroup = Mid(BackRefSectNumbers_AnchorBackward_to_RefGroup, 1, Len(BackRefSectNumbers_AnchorBackward_to_RefGroup) - 1) ' remove last comma
    ' add reftips => short ref occurence like "text [ref] text
    RefTip = g_BibEntries(opt_offset_RefTip - 1 + relative_offset)
    RefTip = Mid(RefTip, 1, Len(RefTip) - 1) ' remove last comma
    ' find num_BackRefPageNumbers
    num_BackRefPageNumbers = 0
    num_BackRefSectNumbers = 0
    num_BackRefPageNumbers_AnchorBackward_to_RefGroup = 0
    num_BackRefSectNumbers_AnchorBackward_to_RefGroup = 0
    num_RefTip = 0
    Dim array_BackRefPageNumbers(128) As String
    Dim array_BackRefSectNumbers(128) As String
    Dim array_BackRefPageNumbers_AnchorBackward_to_RefGroup(128) As String
    Dim array_BackRefSectNumbers_AnchorBackward_to_RefGroup(128) As String
    Dim array_RefTip(128) As String
    If find_p Then
        pch = TOOLS__STRTOK(BackRefPageNumbers, ",")
        Do While pch <> vbNullString
            array_BackRefPageNumbers(num_BackRefPageNumbers) = pch
            num_BackRefPageNumbers = num_BackRefPageNumbers + 1
            pch = TOOLS__STRTOK(vbNullString, ",")
        Loop
    End If
    If find_s Then
        pch = TOOLS__STRTOK(BackRefSectNumbers, ",")
        Do While pch <> vbNullString
            array_BackRefSectNumbers(num_BackRefSectNumbers) = pch
            num_BackRefSectNumbers = num_BackRefSectNumbers + 1
            pch = TOOLS__STRTOK(vbNullString, ",")
        Loop
    End If
    pch = TOOLS__STRTOK(BackRefPageNumbers_AnchorBackward_to_RefGroup, ",")
    Do While pch <> vbNullString
        array_BackRefPageNumbers_AnchorBackward_to_RefGroup(num_BackRefPageNumbers_AnchorBackward_to_RefGroup) = pch
        num_BackRefPageNumbers_AnchorBackward_to_RefGroup = num_BackRefPageNumbers_AnchorBackward_to_RefGroup + 1
        pch = TOOLS__STRTOK(vbNullString, ",")
    Loop
    pch = TOOLS__STRTOK(BackRefSectNumbers_AnchorBackward_to_RefGroup, ",")
    Do While pch <> vbNullString
        array_BackRefSectNumbers_AnchorBackward_to_RefGroup(num_BackRefSectNumbers_AnchorBackward_to_RefGroup) = pch
        num_BackRefSectNumbers_AnchorBackward_to_RefGroup = num_BackRefSectNumbers_AnchorBackward_to_RefGroup + 1
        pch = TOOLS__STRTOK(vbNullString, ",")
    Loop
    pch = TOOLS__STRTOK(RefTip, ",")    ' FIXME COMMA
    Do While pch <> vbNullString
        array_RefTip(num_RefTip) = pch
        num_RefTip = num_RefTip + 1
        pch = TOOLS__STRTOK(vbNullString, ",")
    Loop
    '................. integrity check ......................
    If find_p Then
        If num_BackRefPageNumbers <> num_BackRefPageNumbers_AnchorBackward_to_RefGroup And _
            num_BackRefPageNumbers <> num_RefTip Then   ' FIXME COMPARE
            MsgBox "integrity error"
        End If
    End If
    If find_s Then
        If num_BackRefSectNumbers <> num_BackRefSectNumbers_AnchorBackward_to_RefGroup And _
            num_BackRefSectNumbers <> num_RefTip Then   ' FIXME COMPARE
            MsgBox "integrity error"
        End If
    End If
    '................. human readable text generation (you can write plugins here) .....................
    ' 0 - pages only (done faster), 1 - sections only (done slower), 2 - pages & sections only (done slower)
    If opt_FindNumPageSect = 0 Then find_p = 1: find_s = 0
    If opt_FindNumPageSect = 1 Then find_p = 0: find_s = 1
    If opt_FindNumPageSect = 2 Then find_p = 1: find_s = 1
    If opt_FindNumPageSect = 0 Then ' 0 - pages only (done faster)
        If num_BackRefPageNumbers = 1 Then
            str_ext2 = " " + opt_Referenced_on_page + ": "
            refIn = str_ext2 + BackRefPageNumbers + "."
        Else
            'add spaces after each comma => or better re-generate
            str_ext = array_BackRefPageNumbers(0)
            For eCount = 1 To num_BackRefPageNumbers - 1 Step 1
                str_ext = str_ext + ", " + array_BackRefPageNumbers(eCount)
            Next eCount
            str_ext2 = " " + opt_Referenced_on_page + "s: "
            refIn = str_ext2 + str_ext + "."
        End If
    End If
    If opt_FindNumPageSect = 1 Then ' 1 - sections only (done slower)
        If num_BackRefSectNumbers = 1 Then
            str_ext2 = " " + opt_Referenced_in_section + ": "
            refIn = str_ext2 + BackRefSectNumbers + "."
        Else
            'add spaces after each comma => or better re-generate
            str_ext = array_BackRefSectNumbers(0)
            For eCount = 1 To num_BackRefSectNumbers - 1 Step 1
                str_ext = str_ext + ", " + array_BackRefSectNumbers(eCount)
            Next eCount
            str_ext2 = " " + opt_Referenced_in_section + "s: "
            refIn = str_ext2 + str_ext + "."
        End If
    End If
    If opt_FindNumPageSect = 2 Then ' 2 - pages & sections only (done slower)
        If num_BackRefSectNumbers = 1 Then
            str_ext2 = " " + opt_Referenced_in_section + ": "
            refIn = str_ext2 + BackRefSectNumbers + " (page " + BackRefPageNumbers + ")."
        Else
            'add spaces after each comma => or better re-generate
            str_ext = array_BackRefSectNumbers(0) + " (page " + array_BackRefPageNumbers(0) + ")"
            For eCount = 1 To num_BackRefSectNumbers - 1 Step 1
                str_ext = str_ext + ", " + array_BackRefSectNumbers(eCount) + " (page " + array_BackRefPageNumbers(eCount) + ")"
            Next eCount
            str_ext2 = " " + opt_Referenced_in_section + "s: "
            refIn = str_ext2 + str_ext + "."
        End If
    End If
    Selection.Range.InsertAfter refIn
    Dim new_beg As Long
    Dim new_end As Long
    new_beg = Selection.Range.Start + Len(title) + Len(str_ext2)
    new_beg = new_beg - 1 ' FIXME
    new_end = new_beg + 1 ' keep selection on space to let word find next occurence FIXME replace with wildcard/regex search
    Selection.SetRange Start:=new_beg, End:=new_end
    If opt_FindNumPageSect = 0 Then up_limit = num_BackRefPageNumbers
    If opt_FindNumPageSect = 1 Then up_limit = num_BackRefSectNumbers
    If opt_FindNumPageSect = 2 Then up_limit = num_BackRefPageNumbers
    If opt_FindNumPageSect = 0 Then
        With Selection.Find
            ' https://wordmvp.com/FAQs/General/UsingWildcards.htm
            .Text = "[0-9]{1,}"
            '.Text = "^#. ^t"
            '.Text = titleName
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        For eCount = 1 To up_limit Step 1
            Selection.Find.Execute
            If Selection.Find.Found = True Then
                PageSectNo = Selection.Range.Text
                Dim PageSectNumber As String:
                BackRefPageNumber = array_BackRefPageNumbers(eCount - 1)
                BackRefSectNumber = array_BackRefSectNumbers(eCount - 1)
                BackRefPageNumber_AnchorBackward_to_RefGroup = array_BackRefPageNumbers_AnchorBackward_to_RefGroup(eCount - 1)
                BackRefSectNumber_AnchorBackward_to_RefGroup = array_BackRefSectNumbers_AnchorBackward_to_RefGroup(eCount - 1)
                If opt_FindNumPageSect = 0 Then PageSectNumber = BackRefPageNumber
                If opt_FindNumPageSect = 1 Then PageSectNumber = BackRefSectNumber
                If opt_addRefScreenTipsText Then ' add reftips => short ref occurence like "text [ref] text
                    RefTip = array_RefTip(eCount - 1)
                Else
                    RefTip = Left(PageSectNumber, 1024) ' by default
                End If
                ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, _
                    address:="", SubAddress:=BackRefPageNumber_AnchorBackward_to_RefGroup, _
                    ScreenTip:=RefTip, TextToDisplay:="" & PageSectNumber
                If opt_use_wdUnderlineNone Then
                    Selection.Select
                    With Selection.Font
                        .Underline = wdUnderlineNone
                        '.ColorIndex = wdBlack
                    End With
                End If
                ' to continue search
                ref_new_beg = Selection.Range.End
                ref_new_end = ref_new_beg + 1  'keep selection on space to let word find next  occurence ' FIXME
                Selection.SetRange Start:=ref_new_beg, End:=ref_new_end
            End If
        Next eCount
    End If
    If opt_FindNumPageSect = 1 Or opt_FindNumPageSect = 2 Then
        Dim sect_p, page_p As Integer
        Dim pattern As String
        If opt_FindNumPageSect = 1 Then sect_p = 1: page_p = 0: pattern = "\ *\ " ' FIXME PATTERN
        If opt_FindNumPageSect = 2 Then sect_p = 1: page_p = 1: pattern = "\ *\ "
        For eCount = 1 To up_limit Step 1
            If sect_p Then ' secitons: can be either in format "num" or "num.num" or "num.num.num"
                With Selection.Find
                    ' https://wordmvp.com/FAQs/General/UsingWildcards.htm
                    .Text = pattern     ' In section 12.24 (page 23)
                    '.Text = "^#. ^t"
                    '.Text = titleName
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = True
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute
                If Selection.Find.Found = True Then
                    SectNo = Selection.Range.Text
                    Dim SectNumber As String:
                    SectNumber = array_BackRefSectNumbers(eCount - 1)
                    ' remove search spaces
                    SectNo = Mid(SectNo, 2, Len(SectNo) - 2)
                    ' adjust selection range due to search spaces
                    new_beg = Selection.Range.Start + 1
                    new_end = Selection.Range.End - 1
                    Selection.SetRange Start:=new_beg, End:=new_end
                    BackRefSectNumber_AnchorBackward_to_RefGroup = array_BackRefSectNumbers_AnchorBackward_to_RefGroup(eCount - 1)
                    If opt_addRefScreenTipsText Then ' add reftips => short ref occurence like "text [ref] text
                        RefTip = array_RefTip(eCount - 1)
                    Else
                        RefTip = Left(SectNumber, 1024) ' by default
                    End If
                    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, _
                        address:="", SubAddress:=BackRefSectNumber_AnchorBackward_to_RefGroup, _
                        ScreenTip:=RefTip, TextToDisplay:="" & SectNumber
                    If opt_use_wdUnderlineNone Then
                        With Selection.Font
                            .Underline = wdUnderlineNone
                            '.ColorIndex = wdBlack
                        End With
                    End If
                    ' to continue search
                    ref_new_beg = Selection.Range.End
                    ref_new_end = ref_new_beg + 1  'keep selection on space to let word find next  occurence ' FIXME
                    Selection.SetRange Start:=ref_new_beg, End:=ref_new_end
                Else
                End If
            End If
            If page_p Then ' pages: always in format "num"
                With Selection.Find
                    ' https://wordmvp.com/FAQs/General/UsingWildcards.htm
                    .Text = "[0-9]{1,}"
                    '.Text = "^#. ^t"
                    '.Text = titleName
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = True
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute
                If Selection.Find.Found = True Then
                    PageNo = Selection.Range.Text
                    Dim PageNumber As String:
                    PageNumber = array_BackRefPageNumbers(eCount - 1)
                    BackRefPageNumber_AnchorBackward_to_RefGroup = array_BackRefPageNumbers_AnchorBackward_to_RefGroup(eCount - 1)
                    If opt_addRefScreenTipsText Then ' add reftips => short ref occurence like "text [ref] text
                        RefTip = array_RefTip(eCount - 1)
                    Else
                        RefTip = Left(PageNumber, 1024) ' by default
                    End If
                    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, _
                        address:="", SubAddress:=BackRefPageNumber_AnchorBackward_to_RefGroup, _
                        ScreenTip:=RefTip, TextToDisplay:="" & PageNumber
                    If opt_use_wdUnderlineNone Then
                        With Selection.Font
                            .Underline = wdUnderlineNone
                            '.ColorIndex = wdBlack
                        End With
                    End If
                    ' to continue search
                    ref_new_beg = Selection.Range.End
                    ref_new_end = ref_new_beg + 1  'keep selection on space to let word find next  occurence ' FIXME
                    Selection.SetRange Start:=ref_new_beg, End:=ref_new_end
                End If
            End If
        Next eCount
    End If
End Function
Function TOOLS__BUILD_BACKWARD_REFERENCES(g_BibEntries As Variant)
    Dim title As String
    Dim proc_Entries As Integer
    StartTime = Timer
    Debug.Print "TOOLS__BUILD_BACKWARD_REFERENCES start => BibEntriesLen " & g_BibEntriesLen
    proc_Entries = 0
    skip_Entries = 0
    If opt_use_UI Then
        Call UI_PROGRESS_BAR.CONFIGURE(g_BibEntriesLen, 4)
    End If
    ' find the Zotero bibliography
    Selection.GoTo What:=wdGoToBookmark, name:=string_Zotero_Bibliography
    'Selection.Find.ClearFormatting
    Do
        With Selection.Find
            ' https://wordmvp.com/FAQs/General/UsingWildcards.htm
            .Text = "[0-9]{1,}.\ ^t"
            '.Text = "^#. ^t"
            '.Text = titleName
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute
        If Selection.Find.Found = True Then
            bibText = Selection.Range.Text
            refNo = Left(bibText, InStr(bibText, ".") - 1)
            title = g_BibEntries(opt_offset_BibNames - 1 + refNo)
            If opt_work_with_ZoteroOffline Then ' remove "referenced on xxx" from previous run
                ref_beg_s = Selection.Range.Start
                ref_end_s = Selection.Range.End
                ref_beg_m = Selection.Range.Start + Len(bibText)
                ref_end_m = ref_beg_m + Len(title)
                Selection.SetRange Start:=ref_beg_m, End:=ref_end_m
                Selection.MoveEnd Unit:=wdParagraph
                ref_beg_n = Selection.Range.Start
                ref_end_n = Selection.Range.End
                Diff = ref_end_n - ref_end_m
                If Diff > 2 Then
                    Selection.SetRange Start:=ref_end_m + 1, End:=ref_end_n - 1
                    Selection.Delete
                End If
                Selection.SetRange Start:=ref_beg_s, End:=ref_end_s
            End If
            offst = 0
            ' 0 - pages only (done faster), 1 - sections only (done slower), 2 - pages & sections only (done slower)
            If opt_FindNumPageSect = 0 Then offst = opt_offset_BackRefPageNumbers
            If opt_FindNumPageSect = 1 Then offst = opt_offset_BackRefSectNumbers
            If opt_FindNumPageSect = 2 Then offst = opt_offset_BackRefSectNumbers
            refIn = g_BibEntries(offst - 1 + refNo)
            ' select currect title
            new_beg = Selection.Range.Start + Len(bibText)
            new_end = Selection.Range.Start + Len(bibText) + Len(title)
            Selection.SetRange Start:=new_beg, End:=new_end
            ' if does not exist then probably it was mixed section aka [1,2,4,5-8]
            If refIn = "" Then
                skip_Entries = skip_Entries + 1
            Else
                ' ............ insert back hyperlinks ...........
                Call TOOLS__BUILD_BACKWARD_REFERENCES_ON_ENTRY(refNo, refIn, bibText, title, g_BibEntries)
                proc_Entries = proc_Entries + 1
            End If
            If opt_use_UI Then
                Call UI_PROGRESS_BAR.SET_PROGRESS(proc_Entries + skip_Entries)
            End If
         End If
    Loop Until (proc_Entries + skip_Entries) = g_BibEntriesLen
    ' FIXME bug found START => after adding "referenced in XXX" annotation to bib list, previously created bookmarks are not valid
    ' because ranges in them are not valid anymore => need to re-update ranges in these bookmarks
    Call ZZ_INTERNAL__READ_BIB_ENTRIES(g_BibEntries, 2)
    ' clear Bookmarks type forward => hope previously created hyperlinks will work with new bookmarks
    For ki = ActiveDocument.Bookmarks.Count To 1 Step -1
        bName = ActiveDocument.Bookmarks(ki).name
        If InStr(bName, opt_PfxForw) > 0 Then
            ActiveDocument.Bookmarks(ki).Delete
        End If
    Next ki
    Call ZZ_INTERNAL__CREATE_BOOKMARKS_ON_BIB_ENTRIES(g_BibEntries, 2)
    ' FIXME bug found END
    ElapsedTimeSec = Round(Timer - StartTime, 2)
    Debug.Print "TOOLS__BUILD_BACKWARD_REFERENCES done => proc_Entries / skip_Entries " & proc_Entries & " / " & skip_Entries
    Debug.Print "TOOLS__BUILD_BACKWARD_REFERENCES done => time elapsed " & ElapsedTimeSec & " seconds"
End Function
Function TOOLS__WAIT_SEC(n As Long)
    Dim t As Date
    t = Now
    Do
        DoEvents
    Loop Until Now >= DateAdd("s", n, t)
End Function
Function ZZ_INTERNAL__READ_BIB_ENTRIES(ByRef g_BibEntries, passNb As Integer)
    Dim title As String
    If opt_use_UI Then
        If passNb = 1 Then
            If g_BibEntriesMoved Then
                Call UI_PROGRESS_BAR.CONFIGURE(g_foundBib_Entries, 1) ' no of entries did not change in this case
            Else
                Call UI_PROGRESS_BAR.CONFIGURE(opt_group_Length, 1) ' unknown yet no if found entries => set to max
            End If
        ElseIf passNb = 2 Then
            Call UI_PROGRESS_BAR.CONFIGURE(g_foundBib_Entries, 5)
        Else
            MsgBox "ZZ_INTERNAL__READ_BIB_ENTRIES unknown passNb " & passNb & " (set to last 2)"
            Call UI_PROGRESS_BAR.CONFIGURE(g_foundBib_Entries, 5)
        End If
    End If
    ' .................... clear g_BibEntries .....................
    Erase g_BibEntries
    g_foundBib_Entries = 0
    ' ..................... read bib entries ....................
    Selection.GoTo What:=wdGoToBookmark, name:=string_Zotero_Bibliography
    Selection.Find.ClearFormatting
    With Selection.Find
        ' https://wordmvp.com/FAQs/General/UsingWildcards.htm
        .Text = "[0-9]{1,}.\ ^t"
        '.Text = "^#. ^t"
        '.Text = titleName
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    For iCount = 1 To opt_group_Length Step 1
        Selection.Find.Execute
        If Selection.Find.Found = True Then
            bibText = Selection.Range.Text
            If iCount = 1 Then
                g_bibPositionInDocument = Selection.Information(wdActiveEndPageNumber) * Selection.Range.Start
            End If
            refNo = Left(bibText, InStr(bibText, ".") - 1)
            If CInt(refNo) = iCount Then
                Selection.MoveEnd Unit:=wdParagraph
                bibTextFull = Selection.Range.Text
                posit = Len(bibText) + 1
                title = Mid(bibTextFull, posit, Len(bibTextFull) - posit - 1)
                new_beg = Selection.Range.Start + Len(bibText)
                new_end = new_beg + Len(title)
                Selection.SetRange Start:=new_beg, End:=new_end
                ' FIXME need to correct title => sinon titleAnchorForward will be not correct
                Dim title_modified As String
                ' probabilty that title includes "Referenced on page" is likely LOW
                If opt_FindNumPageSect = 0 Then title_modified = InStr(title, opt_Referenced_on_page) '
                If opt_FindNumPageSect = 1 Then title_modified = InStr(title, opt_Referenced_in_section) '
                If opt_FindNumPageSect = 2 Then title_modified = InStr(title, opt_Referenced_in_section) '
                If title_modified Then
                    title_end = title_modified - 2
                    title = Mid(title, 1, title_end)
                End If
                If TOOLS__IS_IN_ARRAY(title, g_BibEntries) = 0 Then
                    ' ............... save title .............
                    g_BibEntries(opt_offset_BibNames + g_foundBib_Entries) = title
                    ' ............... save range (used to make bookmark) .............
                    range_str = CStr(new_beg) + "_" + CStr(new_end)
                    g_BibEntries(opt_offset_BibRanges + g_foundBib_Entries) = range_str
                    g_foundBib_Entries = g_foundBib_Entries + 1
                    If opt_use_UI Then
                        Call UI_PROGRESS_BAR.SET_PROGRESS(g_foundBib_Entries)
                    End If
                Else
                    'title already in array => very likely VBA started repeat search results...
                    g_foundBib_Entries = g_foundBib_Entries
                    ' Exit For ' ???
                End If
            Else
                 'MsgBox "integrity problem: refNo <> iCount"
                 ' very likely VBA started repeat search results
                 Exit For
            End If
        Else
            Exit For
        End If
    Next iCount
End Function
Function TOOLS__BUILD_BIBLIOGRAPHY_ENTRIES()
    Dim title As String
    Dim xwait_ZoteroRefresh_sec As Long
    CacheName = Replace(ActiveDocument.name, ".", "_") & "_cache"
    BibEntriesFile = ActiveDocument.Path & Application.PathSeparator & CacheName & "\g_BibEntries.array"
    If opt_work_with_ZoteroOffline Then
        Debug.Print "note: work_with_ZoteroOffline"
        g_zotero_did_FastRefresh = 1 ' do logic as g_zotero_did_FastRefresh => bib list is not changed
    Else
        ' request to refresh zotero generated content (to remove old results generated by this macro)
        'Debug.Print "start ZoteroRefresh"
        StartTime = Timer
        under_dev = 0
        If under_dev Then
             'Application.Run "Zotero.ZoteroRefresh"
             strProgramName = "C:\Program Files (x86)\Zotero\zotero.exe"
             Retval = Shell(strProgramName, 1)
             Call TOOLS__WAIT_SEC(2)
             ElapsedTimeSec = Round(Timer - StartTime, 2)
             Const opt_wait_ZoteroLaunch_Sec = 3
             'Call TOOLS__WAIT_SEC(opt_wait_ZoteroLaunch_Sec)
             Application.Run "Zotero.ZoteroRefresh"
             AppActivate Retval
             SendKeys "%{F4}", True
        End If
        Application.Run "Zotero.ZoteroRefresh"
        ElapsedTimeSec = Round(Timer - StartTime, 2)
        'Debug.Print "end ZoteroRefresh elapsed time: " & ElapsedTimeSec & " seconds"
        'Debug.Print "looks like need to wait zotero async updating bib list"
        'Debug.Print "start wait " & xwait_ZoteroRefresh_sec & " sec "
        Call TOOLS__WAIT_SEC(opt_wait_ZoteroRefresh_Sec) ' wait for async ZoteroRefresh => how to check if ZoteroRefresh done?
        'Debug.Print "end wait " & xwait_ZoteroRefresh_sec & " sec "
        ' ............ study ElapsedTimeSec to see if zotero did refresh .............
        If ElapsedTimeSec < 1 Then
            'Debug.Print "looks like zotero did not update bib list"
            ' zotero detected that nothing is changed
            g_zotero_did_FastRefresh = 1
        Else
            'Debug.Print "looks like zotero update bib list"
            ' zotero detected something is changed
            g_zotero_did_LongRefresh = 1
        End If
    End If
    ' ............. redo the Zotero_Bibliography if need ..................
    Zotero_BibliographyFound = 0
    For ki = ActiveDocument.Bookmarks.Count To 1 Step -1
        bName = ActiveDocument.Bookmarks(ki).name
        If InStr(bName, string_Zotero_Bibliography) > 0 Then
            Zotero_BibliographyFound = 1
            Exit For
        End If
    Next ki
    If Zotero_BibliographyFound = 0 Then
        ' create Zotero_Bibliography
        ActiveWindow.View.ShowFieldCodes = True
        Selection.Find.ClearFormatting
        With Selection.Find
            .Text = "^d ADDIN ZOTERO_BIBL"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute
        If Selection.Find.Found = True Then
            ' add bookmark for the Zotero bibliography
            With ActiveDocument.Bookmarks
                .Add Range:=Selection.Range, name:=string_Zotero_Bibliography
                .DefaultSorting = wdSortByName
                .ShowHidden = True
            End With
        Else
            MsgBox "cannot find ADDIN ZOTERO_BIBL"
        End If
        ActiveWindow.View.ShowFieldCodes = False
    End If
    ' ............. redo the Zotero_Bibliography if need ..................
    If g_zotero_did_LongRefresh = 1 Then
        g_rebuild_BibEntries = 1
    End If
    If g_zotero_did_FastRefresh = 1 Then
        ' zotero detected that nothing is changed
        ' check if g_BibEntries g_BibEntriesMoved => need to re-generate bookmarks at new places
        If Len(Dir(BibEntriesFile)) <> 0 Then
            g_BibEntriesFile_Found = 1
        End If
        If g_BibEntriesFile_Found = 0 Then
            ' cannot check if g_BibEntries g_BibEntriesMoved
            g_rebuild_BibEntries = 1
        Else
            ' ........................... read from disk ...........................
            If Not TOOLS__READ_ARRAY_FROM_DISK(BibEntriesFile, g_BibEntries) Then
                MsgBox "g_BibEntries Read error"
            End If
            g_foundBib_Entries = g_BibEntries(0)
            ' .......... check if g_BibEntries g_BibEntriesMoved ...............
            Selection.GoTo What:=wdGoToBookmark, name:=string_Zotero_Bibliography
            Selection.Find.ClearFormatting
            With Selection.Find
                ' https://wordmvp.com/FAQs/General/UsingWildcards.htm
                .Text = "[0-9]{1,}.\ ^t"
                '.Text = "^#. ^t"
                '.Text = titleName
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = True
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute
            If Selection.Find.Found = True Then
                bibText = Selection.Range.Text
                g_bibPositionInDocument = CLng(Selection.Information(wdActiveEndPageNumber)) * Selection.Range.Start
                bib_pos_last = CLng(g_BibEntries(1))
                If g_bibPositionInDocument <> bib_pos_last Then ' FIXME CHECK PAGE NO + LINE/COL NO
                    g_BibEntriesMoved = 1
                End If
            Else
                MsgBox "cannot find 1st bib entry => format pb?"
            End If
            If g_BibEntriesMoved Then
                g_rebuild_BibEntries = 1
            End If
        End If
    End If
    ' ......................... main processing ................
    If g_rebuild_BibEntries = 0 Then
        ' NO need to again read from disk (already did)
        Debug.Print "NO need g_rebuild_BibEntries => " & g_foundBib_Entries & " read from cache"
        ' .................... emulate 100% progress .....................
        If opt_use_UI Then
            Call UI_PROGRESS_BAR.CONFIGURE(g_foundBib_Entries, 1)
            Call UI_PROGRESS_BAR.SET_PROGRESS(g_foundBib_Entries)
        End If
    Else
        Debug.Print "g_rebuild_BibEntries start => FastRefresh/LongRefresh/Moved " & g_zotero_did_FastRefresh & " / " & g_zotero_did_LongRefresh & " / " & g_BibEntriesMoved
        StartTime = Timer
        Call ZZ_INTERNAL__READ_BIB_ENTRIES(g_BibEntries, 1)
        g_BibEntries(0) = g_foundBib_Entries ' save length
        g_BibEntries(1) = g_bibPositionInDocument 'save position to check for outdated
        ' ........................... write to disk ...........................
        If Len(Dir(ActiveDocument.Path & Application.PathSeparator & CacheName, vbDirectory)) = 0 Then
            MkDir ActiveDocument.Path & Application.PathSeparator & CacheName
        End If
        If Not TOOLS__DUMP_ARRAY_TO_DISK(BibEntriesFile, g_BibEntries) Then
            MsgBox "g_BibEntries Write error"
        End If
        ElapsedTimeSec = Round(Timer - StartTime, 2)
        Debug.Print "g_rebuild_BibEntries done => g_foundBib_Entries " & g_foundBib_Entries & " elapsed time " & ElapsedTimeSec & " seconds"
    End If
    ' ......................... post processing ................
    ' clear all back refs because ALL back refs are NEVER cached and generated each run
    For iCount = opt_offset_BackRefPageNumbers To opt_offset_BackRefPageNumbers + g_foundBib_Entries Step 1
        g_BibEntries(iCount) = ""
    Next iCount
    ' ...................... clear Bookmarks .....................
    If g_rebuild_BibEntries = 1 Then
        ' clear Bookmarks type forward & backward
        For ki = ActiveDocument.Bookmarks.Count To 1 Step -1
            bName = ActiveDocument.Bookmarks(ki).name
            If InStr(bName, opt_PfxForw) > 0 Then
                ActiveDocument.Bookmarks(ki).Delete
            End If
            If InStr(bName, opt_PfxBack) > 0 Then
                ActiveDocument.Bookmarks(ki).Delete
            End If
        Next ki
    Else
        ' clear Bookmarks type backward & keep forward (because they
        ' are the same since NO bib changes detected, same position in the doc)
        ' backward could change (text moving, moved to other position/page, etc)
        For ki = ActiveDocument.Bookmarks.Count To 1 Step -1
            bName = ActiveDocument.Bookmarks(ki).name
            If InStr(bName, opt_PfxBack) > 0 Then
                ActiveDocument.Bookmarks(ki).Delete
            End If
        Next ki
    End If
    ' .........................clear ALL Hyperlinks .................
    ' ALL because
    ' - forward  => position could change (text moving, moved to other position/page, etc)
    ' - backward => position could change (text moving, moved to other position/page, etc)
    For ki = ActiveDocument.Hyperlinks.Count To 1 Step -1
        subAddr = ActiveDocument.Hyperlinks(ki).SubAddress
        If InStr(subAddr, opt_ZoteroHyperlinkPrefixCommon) > 0 Then
            ActiveDocument.Hyperlinks(ki).Delete
        End If
    Next ki
    ' ...................... add Bookmarks type forward .....................
    If g_rebuild_BibEntries Then
        Call ZZ_INTERNAL__CREATE_BOOKMARKS_ON_BIB_ENTRIES(g_BibEntries, 1)
    Else
        ' .................... emulate 100% progress .....................
        If opt_use_UI Then
            Call UI_PROGRESS_BAR.CONFIGURE(g_foundBib_Entries, 2)
            Call UI_PROGRESS_BAR.SET_PROGRESS(g_foundBib_Entries)
        End If
    End If
    ' ..................... all done .....................
    TOOLS__BUILD_BIBLIOGRAPHY_ENTRIES = g_BibEntries
End Function
Function ZZ_INTERNAL__CREATE_BOOKMARKS_ON_BIB_ENTRIES(g_BibEntries As Variant, passNb As Integer)
    Dim title As String
    Dim iCount As Integer
    If opt_use_UI Then
        If passNb = 1 Then
            Call UI_PROGRESS_BAR.CONFIGURE(g_foundBib_Entries, 2)
        ElseIf passNb = 2 Then
            Call UI_PROGRESS_BAR.CONFIGURE(g_foundBib_Entries, 6)
        Else
            MsgBox "ZZ_INTERNAL__READ_BIB_ENTRIES unknown passNb " & passNb & " (set to last 6)"
            Call UI_PROGRESS_BAR.CONFIGURE(g_foundBib_Entries, 6)
        End If
    End If
    ' add Bookmarks type forward (will be used make connection [1] == > Bookmark)
    ' NOTE: not possible to cache bookmarks because zotero removes them during update operation
    For iCount = 1 To g_foundBib_Entries Step 1
        title = g_BibEntries(opt_offset_BibNames + iCount - 1)
        titleAnchorForward = Left(opt_PfxForw + TOOLS__HASH12(title), 40)
        range_str = g_BibEntries(opt_offset_BibRanges + iCount - 1)
        delim_idx = InStr(range_str, "_")
        bm_beg = CLng(Mid(range_str, 1, delim_idx - 1))
        bm_end = CLng(Mid(range_str, delim_idx + 1, Len(range_str) - delim_idx))
        Set BookmarkRange = ActiveDocument.Range(Start:=0, End:=0)
        BookmarkRange.SetRange Start:=bm_beg, End:=bm_end
        With ActiveDocument.Bookmarks
            .Add Range:=BookmarkRange, name:=titleAnchorForward
            .DefaultSorting = wdSortByName
            .ShowHidden = True
        End With
        If opt_use_UI Then
            Call UI_PROGRESS_BAR.SET_PROGRESS(iCount)
        End If
    Next iCount
End Function
Function TOOLS__IS_IN_ARRAY(stringToBeFound As String, Arr As Variant) As Boolean
  TOOLS__IS_IN_ARRAY = (UBound(Filter(Arr, stringToBeFound)) > -1)
End Function
Function TOOLS__DUMP_ARRAY_TO_DISK(ByVal FileName As String, ByRef Arr, _
    Optional ByVal OverWrite As Boolean = True) As Boolean
  'Writes an array to disk
  Dim ff As Integer
  If Dir(FileName) <> "" Then
    If Not OverWrite Then Exit Function
  End If
  On Error GoTo ExitPoint
  ff = FreeFile
  Open FileName For Binary Access Write Lock Read Write As #ff
  Put #ff, , Arr
  Close #ff
  TOOLS__DUMP_ARRAY_TO_DISK = True
ExitPoint:
End Function
Function TOOLS__READ_ARRAY_FROM_DISK(ByVal FileName As String, ByRef Arr) As Boolean
    'Reads an array from disk
    Dim ff As Integer
    If Dir(FileName) = "" Then Exit Function
    On Error GoTo ExitPoint
    ff = FreeFile
    Open FileName For Binary Access Read Lock Write As #ff
    Get #ff, , Arr
    Close #ff
    TOOLS__READ_ARRAY_FROM_DISK = True
ExitPoint:
End Function
Function TOOLS__MANUAL_REMOVE_HYPERLINKS_AND_BOOKMARKS()
    ' clear Hyperlinks
    For ki = ActiveDocument.Hyperlinks.Count To 1 Step -1
        subAddr = ActiveDocument.Hyperlinks(ki).SubAddress
        If InStr(subAddr, opt_ZoteroHyperlinkPrefixCommon) > 0 Then
            ActiveDocument.Hyperlinks(ki).Delete
        End If
    Next ki
    ' clear Bookmarks
    For ki = ActiveDocument.Bookmarks.Count To 1 Step -1
        bName = ActiveDocument.Bookmarks(ki).name
        If InStr(bName, opt_ZoteroHyperlinkPrefixCommon) > 0 Then
            ActiveDocument.Bookmarks(ki).Delete
        End If
    Next ki
End Function
Function BB__PERFORM_TWO_WAY_LINKING__PERF_TEST_01()
    TestDocName = Replace(ActiveDocument.name, ".", "_") & "_cache"
    BibEntriesFile = ActiveDocument.Path & Application.PathSeparator & TestDocName
    For ki = 0 To 10
        strProgramName = "C:\Program Files (x86)\Zotero\zotero.exe"
        Retval = Shell(strProgramName, 1)
        Call TOOLS__WAIT_SEC(5)
        Application.Run "Zotero.ZoteroRefresh"
        Call TOOLS__WAIT_SEC(10)
        AppActivate Retval
        SendKeys "%{F4}", True
        Call TOOLS__WAIT_SEC(2)
        Call BB__PERFORM_TWO_WAY_LINKING__PERF_TEST_01
        TestDocName = "T" & ki & "_" & Replace(ActiveDocument.name, ".docm", ".PDF")
        TestDocName = ActiveDocument.Path & Application.PathSeparator & TestDocName
        ActiveDocument.ExportAsFixedFormat OutputFileName:= _
            TestDocName, _
            ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True, OptimizeFor:= _
            wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
            Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
            CreateBookmarks:=wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
            BitmapMissingFonts:=True, UseISO19005_1:=False
        Call TOOLS__WAIT_SEC(2)
    Next ki
End Function



