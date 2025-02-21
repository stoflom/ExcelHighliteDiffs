'Subroutine to scan the ActiveCell text for <ins>..</ins> and <del>...</del> tags
'then delete the tags and take note of their positions. The cleaned-up text is then copied to the cell
'at the right (any data there is overwritten, a blank column must be inserted before) and the
'text in the right cell is then turned red.strikethrough if it was deleted and blue.underlined
'if it was inserted. Then routine then moves one row down and repeats until the last row. 
'The Workbook is saved every once-in-a-while, this seems to help Excell from going non-responsive.
'The Ctl-Break interrupt is also enabled.

'If the tags do not come in opening-closing pairs I do not know what will happen!

'The subroutine caters for a combined total of 201 tags per cell. This can be changed
'by resizing the first index of array updates(200,2) below.
'
'C Lombard (4 Feb 2025)

Sub TurnTextRedBlue()
    On Error GoTo MyErrorHandler 'See https://stackoverflow.com/questions/3279826/interrupt-abort-a-vba-loop

    Dim startPos As Long
    Dim startPosDel As Long
    Dim startPosIns As Long
    Dim endPos As Long
    Dim textLength As Long
    Dim var As String
    Dim updates(200, 2) As Long 'Array of (start,length, del=0/ins=1) updates
    Dim numUpdates As Long
    Dim cellR As Range
    Dim progress As Long
    Dim StopStop As Long
    Dim LastRow As Long
    Dim SaveMod As Long 'Save sheet every SaveMod rows

    SaveMod = 1000 'Save every 1000 rows
    Application.ScreenUpdating = False 'Do not update screen
    Application.EnableCancelKey = xlErrorHandler 'Enable Ctl-Break interrupt (ctl-Fn-End on my keyboard)

    LastRow = ActiveSheet.Range("A:A").SpecialCells(xlCellTypeLastCell).Row
    StopStop = LastRow 'Maximum number of cells to do
    progress = 0

    Do While progress < StopStop
        If progress Mod SaveMod = 0 Then 'Save every 1000 rows
            ThisWorkbook.Save
        End If

        var = ActiveCell.Value
        If Len(var) > 11 Then 'Must have at least <ins></ins> (11 chars)
            numUpdates = 0 'Counts updates
            startPos = 1

            ' Check if the cell contains the <del> or <ins> tag next
            Do
                startPosIns = InStr(startPos, var, "<ins>", 1)
                startPosDel = InStr(startPos, var, "<del>", 1)
                If startPosIns = 0 Then startPosIns = Len(var)
                If startPosDel = 0 Then startPosDel = Len(var)

                If (startPosDel < Len(var)) And (startPosDel < startPosIns) Then
                    endPos = InStr(startPosDel, var, "</del>", 1) - 5
                    textLength = endPos - startPosDel
                    updates(numUpdates, 0) = startPosDel
                    updates(numUpdates, 1) = textLength
                    updates(numUpdates, 2) = 0 '0->del
                    startPos = endPos + 11 'Start next after </del>
                    numUpdates = numUpdates + 1
                ElseIf (startPosIns < Len(var)) And (startPosIns < startPosDel) Then
                    endPos = InStr(startPosIns, var, "</ins>", 1) - 5
                    textLength = endPos - startPosIns
                    updates(numUpdates, 0) = startPosIns
                    updates(numUpdates, 1) = textLength
                    updates(numUpdates, 2) = 1 '1->ins
                    startPos = endPos + 11 'Start next after </ins>
                    numUpdates = numUpdates + 1
                End If
            Loop While (startPosIns < Len(var) - 10) Or (startPosDel < Len(var) - 10)

            'Now do deletes
            For i = 0 To numUpdates - 1
                startPos = updates(i, 0) - i * 11
                updates(i, 0) = startPos 'Fix startpos for deletion
                textLength = updates(i, 1)
                var = DelChars(var, startPos, 5) 'delete the <tag> at startPos
                var = DelChars(var, startPos + textLength, 6) 'delete the </tag> at startPos + textLength
            Next i

            Set cellR = ActiveCell.Offset(0, 1) 'Cell to the right (assumed blank)
            With cellR
                .Value = var
                With .Font
                    .Color = vbBlack
                    .Strikethrough = False
                    .Underline = False
                End With
            End With

            ' Turn the text red/blue
            For i = 0 To numUpdates - 1
                startPos = updates(i, 0)
                textLength = updates(i, 1)
                If updates(i, 2) = 0 Then
                    With cellR.Characters(startPos, textLength).Font
                        .Color = vbRed
                        .Strikethrough = True
                    End With
                ElseIf updates(i, 2) = 1 Then
                    With cellR.Characters(startPos, textLength).Font
                        .Color = vbBlue
                        .Underline = True
                    End With
                End If
            Next i
        End If

        ActiveCell.Offset(1, 0).Select
        progress = progress + 1
    Loop

MyErrorHandler:
    If Err.Number = 18 Then '18 =User interrupt
        MsgBox "You clicked Ctrl + Break"
        Exit Sub
    End If
End Sub

Function DelChars(str As String, start As Long, length As Long) As String
    DelChars = Left(str, start - 1) & Mid(str, start + length)
End Function






