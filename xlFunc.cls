Option Compare Database

Public savedName As String

Public Sub xlCreate(ByVal fName As String, ByVal sheetName As String)

If strCheck(fName, ".xlsx", 5, True) Then
    '処理を続行する
Else
    'ファイル名不正のため、処理中断
    MsgBox fName & vbCrLf & "の拡張子が.xlsx でないため、処理を中断します。"
    End
End If


Dim xlObj As Object, xlBook As Object 'workbook型は、Excel Libraryの参照が必要となるため、Objectで宣言
Dim path As String

    path = Application.CurrentProject.path & "\"
    fName = path & fName
    
    Set xlObj = CreateObject("Excel.Application")
    Set xlBook = xlObj.Workbooks.Add
        xlBook.SaveAs (fName)
        savedName = fName
        xlBook.Close


Set xlObj = Nothing
Set xlBook = Nothing


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
