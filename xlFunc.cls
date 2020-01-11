Option Compare Database
Option Explicit


Public Sub xlsheetInsert(ByVal fName As String, ByVal sheetName As String)
Dim xlObj As Object, xlBook As Object 'workbook型は、Excel Libraryの参照が必要となるため、Objectで宣言

    Set xlObj = CreateObject("Excel.Application")
    Set xlBook = xlObj.Workbooks
    xlBook.Open fName
    xlObj.Worksheets.Add
    xlObj.ActiveSheet.Name = sheetName
    xlObj.ActiveWorkbook.Close SaveChanges:=True

Set xlBook = Nothing
Set xlObj = Nothing

End Sub


Public Sub xlCreate(ByVal fName As String, ByVal sheetName As String)
Dim xlObj As Object, xlBook As Object 'workbook型は、Excel Libraryの参照が必要となるため、Objectで宣言


If strCheck(fName, ".xlsx", 5, True) Then
    '処理を続行する
Else
    'ファイル名不正のため、処理中断
    MsgBox fName & vbCrLf & "の拡張子が.xlsx でないため、処理を中断します。"
    End
End If

    
    Set xlObj = CreateObject("Excel.Application")
    Set xlBook = xlObj.Workbooks.Add
        xlBook.Sheets(1).Name = sheetName
        xlBook.SaveAs (fName)
        xlBook.Close

Set xlBook = Nothing
Set xlObj = Nothing



End Sub


Private Function strCheck(ByVal str As String, ByVal CheckStr As String, ByVal strLen As Long, Optional isRight As Boolean = False)

If isRight = True Then
    If Right(str, strLen) = CheckStr Then
        strCheck = True
    Else
        strCheck = False
    End If

ElseIf isRight = False Then
    If Left(str, strLen) = CheckStr Then
        strCheck = True
    Else
        strCheck = False
    End If

Else
    strCheck = False
End If

End Function


Public Property Get Name()
    Name = "xlFunc"
End Property

Private Sub Class_Initialize()
    Debug.Print Now & vbTab & Me.Name & "はコンストラクトされました。"
End Sub

Private Sub Class_Terminate()
    Debug.Print Now & vbTab & Me.Name & "はデトラクトされました。"
End Sub
