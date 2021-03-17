# blog.github.io
'''vb
Sub getExcelFile(sFolderPath As String)
On Error Resume Next
Dim f As String
Dim file() As String
Dim x
k = 1
 
ReDim file(1)
file(1) = sFolderPath & "\"
 
    f = Dir(file(1) & "*.xlsx")     '通配符*.*表示所有文件，*.xlsx Excel文件
    Do Until f = ""
       'Range("a" & x) = f
       Range("a" & x).Hyperlinks.Add Anchor:=Range("a" & x), Address:=file(i) & f, TextToDisplay:=f
        x = x + 1
        f = Dir
    Loop
 
End Sub
'''
