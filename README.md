# blog.github.io
```vb
Function ItemInArrayLoop(aData, vEle) As Boolean
    ItemInArrayLoop = False
    For i = 1 To ubund(aData, 1)
        For j = 1 To UBound(aData, 2)
            If aData(i, j) = vItem Then
                ItemInArrayLoop = True
                Exit Function
            End If
        Next j
    Next i
End Function
    
```
BF8Y8-GN2QH-T84XB-QVY3B-RC4DF
NYWVH-HT4XC-R2WYW-9Y3CM-X4V3Y
