using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the Excel file as the first argument.");
            return;
        }

        string excelFilePath = args[0];

        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
        Excel.Worksheet worksheet = workbook.Sheets[1];
        Excel.Range range = worksheet.UsedRange;
        int lastRow = range.Rows.Count;

        int saveMod = 1000; // Save every 1000 rows
        int progress = 0;

        excelApp.ScreenUpdating = false;
        excelApp.EnableCancelKey = Excel.XlEnableCancelKey.xlErrorHandler;

        try
        {
            for (int i = 1; i <= lastRow; i++)
            {
                if (progress % saveMod == 0)
                {
                    workbook.Save();
                }

                Excel.Range cell = worksheet.Cells[i, 1];
                string cellValue = cell.Value?.ToString() ?? string.Empty;

                if (cellValue.Length > 11)
                {
                    var updates = new (int startPos, int length, int type)[201];
                    int numUpdates = 0;
                    int startPos = 0;

                    while (true)
                    {
                        int startPosIns = cellValue.IndexOf("<ins>", startPos, StringComparison.Ordinal);
                        int startPosDel = cellValue.IndexOf("<del>", startPos, StringComparison.Ordinal);

                        if (startPosIns == -1) startPosIns = cellValue.Length;
                        if (startPosDel == -1) startPosDel = cellValue.Length;

                        if (startPosDel < cellValue.Length && startPosDel < startPosIns)
                        {
                            int endPos = cellValue.IndexOf("</del>", startPosDel, StringComparison.Ordinal) - 5;
                            int textLength = endPos - startPosDel;
                            updates[numUpdates] = (startPosDel, textLength, 0);
                            startPos = endPos + 11;
                            numUpdates++;
                        }
                        else if (startPosIns < cellValue.Length && startPosIns < startPosDel)
                        {
                            int endPos = cellValue.IndexOf("</ins>", startPosIns, StringComparison.Ordinal) - 5;
                            int textLength = endPos - startPosIns;
                            updates[numUpdates] = (startPosIns, textLength, 1);
                            startPos = endPos + 11;
                            numUpdates++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    for (int j = 0; j < numUpdates; j++)
                    {
                        startPos = updates[j].startPos - j * 11;
                        cellValue = DeleteChars(cellValue, startPos, 5);
                        cellValue = DeleteChars(cellValue, startPos + updates[j].length, 6);
                    }

                    Excel.Range cellRight = cell.Offset[0, 1];
                    cellRight.Value = cellValue;
                    cellRight.Font.Color = ConsoleColor.Black;
                    cellRight.Font.Strikethrough = false;
                    cellRight.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone;

                    for (int j = 0; j < numUpdates; j++)
                    {
                        startPos = updates[j].startPos;
                        int textLength = updates[j].length;
                        if (updates[j].type == 0)
                        {
                            cellRight.Characters[startPos + 1, textLength].Font.Color = ConsoleColor.Red;
                            cellRight.Characters[startPos + 1, textLength].Font.Strikethrough = true;
                        }
                        else if (updates[j].type == 1)
                        {
                            cellRight.Characters[startPos + 1, textLength].Font.Color = ConsoleColor.Blue;
                            cellRight.Characters[startPos + 1, textLength].Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
                        }
                    }
                }

                progress++;
            }

            workbook.Save();
        }
        catch (COMException ex)
        {
            if (ex.ErrorCode == -2146827284)
            {
                Console.WriteLine("You clicked Ctrl + Break");
            }
            else
            {
                throw;
            }
        }
        finally
        {
            workbook.Close(false);
            excelApp.Quit();
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);
        }
    }

    static string DeleteChars(string str, int start, int length)
    {
        return str.Substring(0, start) + str.Substring(start + length);
    }
}
