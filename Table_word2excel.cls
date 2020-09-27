VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Word2Excel()
Dim tb As Table
Dim r As Row
Dim cl As Cell
Dim txt As String
Dim head As String
head = ""
' create excel
    Dim exc As New Excel.Application
    exc.Visible = True
    Dim wb As Workbook
    Dim ws As Excel.Worksheet
    Dim cel As Excel.Range
    Set wb = exc.Workbooks.Add
    Set ws = wb.Worksheets(wb.Worksheets.Count)
    Set cel = ws.Range("b1")
' end create excel
For Each tb In Tables
    For Each r In tb.Rows
        If r.Cells.Count = 1 Then
            head = r.Cells(1).Range.FormattedText
        Else
            cel.Offset(0, -1).Value = head
            For Each cl In r.Cells
                txt = cl.Range.FormattedText
                txt = Replace(txt, Chr(7), "", , , vbTextCompare)
                For i = 1 To Len(cl.Range.FormattedText) - 1
                    Debug.Print i & vbTab & AscW(Mid(txt, i, 1))
                    If (AscW(Mid(txt, i, 1)) < 35) And (AscW(Mid(txt, i, 1)) <> 13) And (AscW(Mid(txt, i, 1)) <> 32) Then
                        Debug.Print Mid(txt, i, 1) & vbTab & AscW(Mid(txt, i, 1)) & vbTab & txt
                    End If
                Next i
                txt = Replace(txt, Chr(13), vbCrLf, , , vbTextCompare)
                cel.Value = txt
                Set cel = cel.Offset(0, 1)
            Next cl
            Set cel = cel.Offset(1, 2 - cel.Column)
        End If
    Next r
Next tb
MsgBox ("end")
End Sub
