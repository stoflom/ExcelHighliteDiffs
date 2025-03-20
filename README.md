Excel Subroutine to scan the ActiveCell text for '\<ins\>...\<\/ins\>' and '\<del\>...\<\/del\>' tags
then delete the tags and take note of their positions. The cleaned-up text is then copied to the cell
at the right (any data there is overwritten, a blank column must be inserted before) and the
text in the right cell is then turned red.strikethrough if it was deleted and blue.underlined
if it was inserted. Then routine then moves one row down and repeats until the last row. 
The Workbook is saved every once-in-a-while, this seems to help Excel from going non-responsive.
The Ctl-Break interrupt is also enabled.

If the tags do not come in opening-closing pairs I do not know what will happen!


C Lombard (4 Feb 2025)
