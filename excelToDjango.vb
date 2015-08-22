Dim sh As Worksheet
 
Sub genCode()
Set sh = ActiveWorkbook.Sheets(1)
 Dim attribWasBlank As Boolean
 Dim constructor As String
 
 constructor = vbCrLf + "def __unicode__(self):  #Python 3.3 __str__" + vbCrLf + vbTab
 constructor = constructor + "return self.name"
 
 Dim lastRow As Long


 For k = 2 To 1048576
    If sh.Cells(k, 1).Value = "" Then
        Exit For
    End If
    
        Dim modelCode, attribCode As String
        Dim classOccurance As Long
        className = sh.Cells(k, 1).Value
        fieldName = sh.Cells(k, 2).Value
        fieldType = sh.Cells(k, 3).Value
        attribName = sh.Cells(k, 4).Value
        attribVal = sh.Cells(k, 5).Value
        
        'If Class does not change
        If sh.Cells(k + 1, 1).Value = sh.Cells(k, 1).Value And sh.Cells(k + 1, 1).Value <> "" Then
            If modelCode = "" Then
                classOccurance = k
                modelCode = "class " + className + "(models.Model):" + vbCrLf
            End If
            
            'If next field does not change
            If sh.Cells(k + 1, 2).Value = sh.Cells(k, 2).Value Then
                If attribCode = "" Then
                    attribCode = attribCode + vbTab + fieldName + " = models." + fieldType + "("
                    attribWasBlank = True
                End If
                
                If sh.Cells(k, 4).Value <> "" Then
                    attribCode = attribCode + CStr(sh.Cells(k, 4).Value) + "=" + CStr(sh.Cells(k, 5).Value)
                End If
                
                If sh.Cells(k + 1, 2).Value = sh.Cells(k, 2).Value Then
                    attribCode = attribCode + ", "
                Else
                    attribCode = attribCode + ")" + vbCrLf
                End If
                            
            'If next field changes
            ElseIf sh.Cells(k + 1, 2).Value <> sh.Cells(k, 2).Value And sh.Cells(k + 1, 2).Value <> "" Then
                If attribCode = "" Then
                    attribCode = attribCode + vbTab + fieldName + " = models." + fieldType + "("
                    attribWasBlank = True
                End If
                
                If sh.Cells(k, 4).Value <> "" Then
                    attribCode = attribCode + CStr(sh.Cells(k, 4).Value) + "=" + CStr(sh.Cells(k, 5).Value)
                End If
                
                If sh.Cells(k + 1, 2).Value = sh.Cells(k, 2).Value Then
                    attribCode = attribCode + ", "
                Else
                    attribCode = attribCode + ")" + vbCrLf
                    modelCode = modelCode + attribCode
                    attribCode = ""
                End If
            End If
                    
        'If Class changes
        ElseIf sh.Cells(k + 1, 1).Value <> sh.Cells(k, 1).Value Then
            
            If modelCode = "" Then
                classOccurance = k
                modelCode = "class " + className + "(models.Model):" + vbCrLf
            End If
            
            'If field does not change
            If sh.Cells(k + 1, 2).Value = sh.Cells(k, 2).Value Then
                If attribCode = "" Then
                    attribCode = attribCode + vbTab + fieldName + " = models." + fieldType + "("
                   
                    attribWasBlank = True
                End If
                
                If sh.Cells(k, 4).Value <> "" Then
                    attribCode = attribCode + CStr(sh.Cells(k, 4).Value) + "=" + CStr(sh.Cells(k, 5).Value)
                End If
                
                If sh.Cells(k + 1, 2).Value = sh.Cells(k, 2).Value Then
                    attribCode = attribCode + ", "
                Else
                    attribCode = attribCode + ")" + vbCrLf
                End If
                            
            'If field changes. Code must hit here
            ElseIf sh.Cells(k + 1, 2).Value <> sh.Cells(k, 2).Value Then
             If attribCode = "" Then
                    attribCode = attribCode + vbTab + fieldName + " = models." + fieldType + "("
                   
                    attribWasBlank = True
                End If
                
                If sh.Cells(k, 4).Value <> "" Then
                    attribCode = attribCode + CStr(sh.Cells(k, 4).Value) + "=" + CStr(sh.Cells(k, 5).Value)
                End If
                
                If sh.Cells(k + 1, 2).Value = sh.Cells(k, 2).Value Then
                    attribCode = attribCode + ", "
                Else
                    attribCode = attribCode + ")" + vbCrLf
                End If
                sh.Cells(classOccurance, 6).Value = modelCode + attribCode + constructor
                modelCode = ""
                attribCode = ""
            End If
        End If
        
    Next k
End Sub
