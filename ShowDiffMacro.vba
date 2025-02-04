'Subroutine to scan the ActiveCell text for <ins>..</ins> and <del>...</del> tags
'then delete the tags and take note of their positions. The cleanup text is then copied to the cell
'at the right (any data there is overwritten, a blank column must be inserted before) and the
'text in the right cell is then turned red.strikethrough if it was deleted and blue.underlined
'if it was inserted. Then routine then moves one cell down and repeats until an empty cell is found.

'If the tags do not come in opening-closing pairs I do not know what will happen!

'The subroutine caters for a combined total of 201 tags/per cell. This can bechanged
'by resizing the first index of array updates(200,2) below.
'
'C Lombard (4 Feb 2025)

Sub TurnTextRedBlue()
    Dim startPos As Long
    Dim startPosDel As Long
    Dim startPosIns As Long
    Dim endPos As Long
    Dim textLength As Long
    Dim var As String
    Dim sleft As String
    Dim sright As String
    Dim updates(200, 2) As Long   'Array of (start,length, del=0/ins=1) updates
    Dim numUpdates As Long
    Dim cellR As range

   
        ' Loop through each cell in the worksheet
        Do While Not IsEmpty(ActiveCell)
           
            Set cellR = ActiveCell.Offset(0, 1)
            
            var = ActiveCell.Value
            slength = Len(var)
           
            numUpdates = 0  'Counts updates
            startPos = 1
            endPos = 1
            ' Check if the cell contains the <del> or <ins> tag next
            Do
                startPosIns = InStr(startPos, var, "<ins>", 1)
                startPosDel = InStr(startPos, var, "<del>", 1)
                If (startPosDel > 0) And (startPosDel < startPosIns) Then
                    endPos = InStr(startPosDel, var, "</del>", 1) - 5 '<del> will be deleted
                    textLength = endPos - startPosDel
                    updates(numUpdates, 0) = startPosDel
                    updates(numUpdates, 1) = textLength
                    updates(numUpdates, 2) = 0 '0->del
                    startPos = endPos + 11  'Start next after </del>
                    numUpdates = numUpdates + 1
                Else
                    If (startPosIns > 0) Then
                        endPos = InStr(startPosIns, var, "</ins>", 1) - 5 '<ins> will be deleted
                        textLength = endPos - startPosIns
                        updates(numUpdates, 0) = startPosIns
                        updates(numUpdates, 1) = textLength
                        updates(numUpdates, 2) = 1 '1->ins
                        startPos = endPos + 11  'Start next after </ins>
                        numUpdates = numUpdates + 1
                    End If
                End If
            
            Loop While (startPosIns > 0) Or (startPosDel > 0)   ' no tags found
                
             'Now do deletes

            For i = 0 To numUpdates - 1
                startPos = updates(i, 0) - i * 11
                textLength = updates(i, 1)
                'delete the <tag> at startPos
                var = DelChars(var, startPos, 5)
                '   delete the </tag> at startPos + textLength
                var = DelChars(var, startPos + textLength, 6)
                'Reduce all next startPos'with 6
            Next i
             
             
             'Shift to cell on right (assumed empty)
             cellR.Value = var
                        
        
            ' Turn the text red/blue
            With cellR
                For i = 0 To numUpdates - 1  'del->red
                    startPos = updates(i, 0) - i * 11
                   textLength = updates(i, 1)
                   If updates(i, 2) = 0 Then
                       With .Characters(startPos, textLength).Font
                           .Color = vbRed
                           .Strikethrough = True
                       End With
                    ElseIf updates(i, 2) = 1 Then 'ins->blue
                        With .Characters(startPos, textLength).Font
                           .Color = vbBlue
                           .Underline = True
                       End With
                    End If
                Next i
            End With
            
             
           ' Move 1 row down
            ActiveCell.Offset(1, 0).Select
        Loop

End Sub

Function DelChars(str As String, start As Long, length As Long)
    'NOTE string.Delete does not work if more than 255 characters in the string,
    'therefore this code.
    slength = Len(str)
    sleft = Left(str, start - 1) 'Delete from right
    sright = Right(str, slength - start - length + 1) 'Delete from left
    DelChars = sleft + sright 'Concatenate
End Function
