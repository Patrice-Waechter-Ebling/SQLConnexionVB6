Attribute VB_Name = "Noyau"
Option Explicit
Dim g_connData As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim fld As ADODB.Field
Dim recCount As Long
Dim i As Integer, x As Integer
Dim li As ListItem

'tables
Public Employes As New ListitemsView
Public Function AquerirEmployes()
    Employes.ListView1.ListItems.Clear
    rs.Open "GrbEmploye", g_connData, adOpenForwardOnly, adLockReadOnly
    LoadListViewFromRecordset Employes.ListView1, rs
    frmLogin.Combo1.ListIndex = frmLogin.Combo1.ListCount - 1
    rs.Close
End Function
Public Function Login(username As String, password As String)
For x = 1 To Employes.ListView1.ListItems.Count - 1
    If username = Employes.ListView1.ListItems(x).SubItems(3) Then
        MsgBox "Bienvenue " + Employes.ListView1.ListItems(x).SubItems(3)
    Else
    Debug.Print Employes.ListView1.ListItems(x).SubItems(3) + "=" + Employes.ListView1.ListItems(x).SubItems(2)
    End If
Next
End Function
Sub main()
    g_connData.Open "Driver={SQL Server};Server=192.168.1.17;Database=GRB2023;Trusted_Connection=Yes;"
    frmLogin.Show
End Sub
Sub LoadListViewFromRecordset(LV As ListView, rs As ADODB.Recordset, Optional MaxRecords As Long)
 On Error Resume Next
   Dim fld As ADODB.Field, alignment As Integer
    Dim recCount As Long, i As Long, fldName As String
    Dim li As ListItem
    LV.ListItems.Clear
    LV.ColumnHeaders.Clear
    For Each fld In rs.Fields
        Select Case fld.Type
            Case adBoolean, adCurrency, adDate, adDecimal, adDouble
                alignment = lvwColumnRight
            Case adInteger, adNumeric, adSingle, adSmallInt, adVarNumeric
                alignment = lvwColumnRight
            Case adBSTR, adChar, adVarChar, adVariant
                alignment = lvwColumnLeft
            Case Else
                alignment = 0
        End Select
        If alignment <> -1 Then
            If LV.ColumnHeaders.Count = 0 Then alignment = lvwColumnLeft
            LV.ColumnHeaders.Add , , fld.Name, fld.DefinedSize * 200, _
                alignment
        End If
    Next
    If LV.ColumnHeaders.Count = 0 Then Exit Sub
    rs.MoveFirst
    Do Until rs.EOF
        recCount = recCount + 1
        fldName = LV.ColumnHeaders(1).Text
        Set li = LV.ListItems.Add(, , rs.Fields(fldName) & "")
        For i = 2 To LV.ColumnHeaders.Count
            fldName = LV.ColumnHeaders(i)
            li.ListSubItems.Add , , rs.Fields(fldName) & ""
        Next
        If recCount = MaxRecords Then Exit Do
        rs.MoveNext
    Loop
End Sub
Sub LoadComboFromRecordset(LV As ComboBox, rs As ADODB.Recordset, Optional MaxRecords As Long)
On Error Resume Next
    Dim fld As ADODB.Field, alignment As Integer
    Dim recCount As Long, i As Long, fldName As String
    LV.Clear
    rs.MoveFirst
    Do Until rs.EOF
        recCount = recCount + 1
        LV.AddItem rs.Fields(1)
        If recCount = MaxRecords Then Exit Do
        rs.MoveNext
    Loop
End Sub
Sub ListViewAdjustColumnWidth(LV As ListView, Optional AccountForHeaders As Boolean)
#If USE_API Then
    Dim col As Integer, lParam As Long
    If AccountForHeaders Then
        lParam = LVSCW_AUTOSIZE_USEHEADER
    Else
        lParam = LVSCW_AUTOSIZE
    End If
    For col = 1 To LV.ColumnHeaders.Count
        SendMessage LV.hWnd, LVM_SETCOLUMNWIDTH, col, lParam
    Next
#Else
    Dim row As Long, col As Long
    Dim width As Single, maxWidth As Single
    Dim saveFont As StdFont, saveScaleMode As Integer
    Dim cellText As String
    If LV.ListItems.Count = 0 Then Exit Sub
    Set saveFont = LV.Parent.Font
    Set LV.Parent.Font = LV.Font
    saveScaleMode = LV.Parent.ScaleMode
    LV.Parent.ScaleMode = vbTwips
    For col = 1 To LV.ColumnHeaders.Count
        maxWidth = 0
        If AccountForHeaders Then
            maxWidth = LV.Parent.TextWidth(LV.ColumnHeaders(col).Text) + 200
        End If
        For row = 1 To LV.ListItems.Count
            If col = 1 Then
                cellText = LV.ListItems(row).Text
            Else
                cellText = LV.ListItems(row).ListSubItems(col - 1).Text
            End If
            width = LV.Parent.TextWidth(cellText) + 200
            If width > maxWidth Then maxWidth = width
        Next
        LV.ColumnHeaders(col).width = maxWidth
    Next
    Set LV.Parent.Font = saveFont
    LV.Parent.ScaleMode = saveScaleMode
#End If
End Sub
Sub ListViewSortOnNonStringField(LV As ListView, ByVal ColumnIndex As Integer, Optional SortOrder As ListSortOrderConstants, Optional IsDateValue As Boolean)
    Dim li As ListItem, number As Double, newIndex As Integer
    Dim minValue As Double
    LV.Visible = False
    LV.Sorted = False
    LV.ColumnHeaders.Add , , "dummy column", 1000
    newIndex = LV.ColumnHeaders.Count - 1
    For Each li In LV.ListItems
        If IsDateValue Then
            number = DateValue(li.ListSubItems(ColumnIndex - 1))
        Else
            number = CDbl(li.ListSubItems(ColumnIndex - 1))
        End If
        If number < minValue Then minValue = number
        li.ListSubItems.Add , , Format$(number, "000000000000000.000")
    Next
    If minValue < 0 Then
        For Each li In LV.ListItems
            number = CDbl(li.ListSubItems(newIndex)) - minValue
            li.ListSubItems(newIndex).Text = Format$(number, "000000000000000.000")
        Next
    End If
        LV.SortKey = newIndex
    LV.SortOrder = SortOrder
    LV.Sorted = True
    LV.ColumnHeaders.Remove newIndex + 1
    For Each li In LV.ListItems
        li.ListSubItems.Remove newIndex
    Next
    LV.Visible = True
End Sub




