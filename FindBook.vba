Function FindBook (nme As String) As Workbook 
'принимает название книги, пролистывает открытые и возвращает такую, начало имени которой совпадает с переменной nme

Dim bk As Workbook
Dim swc As Boolean

Set FindBook = ThisWorkbook
swc = True
For Each bk in Workbooks
    'MsgBox bk.Name
    If Left(bk.name, Len (nme)) = nme Then
        Set FindBook = bk
        swc = False
    End If
Next bk

If swc Then
    MsgBox "Книга " & nme & " не открыта"
    End
End If

End Function

' пример использования 
' IF Findbook("testBook").ReadOnly Then 
'   MsgBox "Книга testBook открыта только для чтения." & vbNewLine & "Необходимо открыть книгу для записи и повторить попытку"
'End If
