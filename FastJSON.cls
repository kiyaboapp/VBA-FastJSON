VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FastJSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================================================
' Class Module Name: FastJSON
' ONE CLASS - SIMPLE AND WORKING - DEEP NESTING FIXED
' November 15, 2025 - Arusha, Tanzania
'====================================================================
Option Explicit

Private Type JSONItem
    key As String
    Value As String
    ItemType As String  ' "V"=value, "A"=array, "O"=object
End Type

Private mItems() As JSONItem
Private mCount As Long

Private Sub Class_Initialize()
    ReDim mItems(0 To 9)
    mCount = 0
End Sub

'====================================================================
' Add simple value
'====================================================================
Public Sub Add(key As String, Value As Variant)
    AddItem key, CStr(Nz(Value, "")), "V"
End Sub

'====================================================================
' Add array
'====================================================================
Public Sub AddArray(key As String, ParamArray Values() As Variant)
    Dim i As Long
    Dim arrStr As String
    
    For i = LBound(Values) To UBound(Values)
        If i > LBound(Values) Then arrStr = arrStr & "|"
        arrStr = arrStr & CStr(Nz(Values(i), ""))
    Next i
    
    AddItem key, arrStr, "A"
End Sub

'====================================================================
' Add nested object
'====================================================================
Public Sub AddObject(key As String, Obj As FastJSON)
    AddItem key, Obj.ToRaw, "O"
End Sub

'====================================================================
' Get value (supports dot notation)
'====================================================================
Public Function GetValue(Path As String) As String
    Dim parts() As String
    Dim i As Long, idx As Long
    Dim tempObj As FastJSON
    
    parts = Split(Path, ".")
    
    ' Get first level
    idx = FindKey(parts(0))
    If idx = -1 Then Exit Function
    
    ' Simple value at root level
    If mItems(idx).ItemType = "V" Then
        If UBound(parts) = 0 Then
            GetValue = mItems(idx).Value
        End If
        Exit Function
    End If
    
    ' Array at root level (can't navigate into arrays)
    If mItems(idx).ItemType = "A" Then
        Exit Function
    End If
    
    ' Object navigation
    If mItems(idx).ItemType = "O" Then
        ' If just requesting the object key itself (no dots)
        If UBound(parts) = 0 Then Exit Function
        
        ' Parse the nested object
        Set tempObj = New FastJSON
        tempObj.Parse mItems(idx).Value
        
        ' Build remaining path and recurse
        Dim remainingPath As String
        For i = 1 To UBound(parts)
            If i > 1 Then remainingPath = remainingPath & "."
            remainingPath = remainingPath & parts(i)
        Next i
        
        ' Recursive call
        GetValue = tempObj.GetValue(remainingPath)
    End If
End Function

'====================================================================
' Get array
'====================================================================
Public Function GetArray(Path As String) As Variant
    Dim parts() As String
    Dim idx As Long
    Dim tempObj As FastJSON
    Dim i As Long
    
    parts = Split(Path, ".")
    
    ' Get first level
    idx = FindKey(parts(0))
    If idx = -1 Then
        GetArray = Array()
        Exit Function
    End If
    
    If UBound(parts) = 0 Then
        ' Direct array
        If mItems(idx).ItemType = "A" Then
            If Len(mItems(idx).Value) = 0 Then
                GetArray = Array()
            Else
                GetArray = Split(mItems(idx).Value, "|")
            End If
        Else
            GetArray = Array()
        End If
    Else
        ' Nested array
        If mItems(idx).ItemType = "O" Then
            Set tempObj = New FastJSON
            tempObj.Parse mItems(idx).Value
            
            ' Build remaining path
            Dim remainingPath As String
            For i = 1 To UBound(parts)
                If i > 1 Then remainingPath = remainingPath & "."
                remainingPath = remainingPath & parts(i)
            Next i
            
            GetArray = tempObj.GetArray(remainingPath)
        Else
            GetArray = Array()
        End If
    End If
End Function

'====================================================================
' Get nested object
'====================================================================
Public Function GetObject(Path As String) As FastJSON
    Dim parts() As String
    Dim idx As Long
    Dim tempObj As FastJSON
    Dim i As Long
    
    parts = Split(Path, ".")
    
    idx = FindKey(parts(0))
    If idx = -1 Then Exit Function
    
    If mItems(idx).ItemType <> "O" Then Exit Function
    
    Set tempObj = New FastJSON
    tempObj.Parse mItems(idx).Value
    
    If UBound(parts) = 0 Then
        ' Return this object
        Set GetObject = tempObj
    Else
        ' Navigate deeper
        Dim remainingPath As String
        For i = 1 To UBound(parts)
            If i > 1 Then remainingPath = remainingPath & "."
            remainingPath = remainingPath & parts(i)
        Next i
        
        Set GetObject = tempObj.GetObject(remainingPath)
    End If
End Function

'====================================================================
' Check if key exists
'====================================================================
Public Function HasKey(Path As String) As Boolean
    Dim parts() As String
    parts = Split(Path, ".")
    
    Dim idx As Long
    idx = FindKey(parts(0))
    
    If idx = -1 Then
        HasKey = False
        Exit Function
    End If
    
    If UBound(parts) = 0 Then
        HasKey = True
        Exit Function
    End If
    
    ' Check nested path
    If mItems(idx).ItemType = "O" Then
        Dim tempObj As FastJSON
        Dim i As Long
        Dim remainingPath As String
        
        Set tempObj = New FastJSON
        tempObj.Parse mItems(idx).Value
        
        For i = 1 To UBound(parts)
            If i > 1 Then remainingPath = remainingPath & "."
            remainingPath = remainingPath & parts(i)
        Next i
        
        HasKey = tempObj.HasKey(remainingPath)
    Else
        HasKey = False
    End If
End Function

'====================================================================
' Get all root keys
'====================================================================
Public Function GetKeys() As Variant
    Dim result() As String
    Dim i As Long
    
    If mCount = 0 Then
        GetKeys = Array()
        Exit Function
    End If
    
    ReDim result(0 To mCount - 1)
    For i = 0 To mCount - 1
        result(i) = mItems(i).key
    Next i
    
    GetKeys = result
End Function

'====================================================================
' Count items
'====================================================================
Public Function Count() As Long
    Count = mCount
End Function

'====================================================================
' Update value
'====================================================================
Public Function UpdateValue(key As String, NewValue As Variant) As Boolean
    Dim idx As Long
    idx = FindKey(key)
    
    If idx <> -1 And mItems(idx).ItemType = "V" Then
        mItems(idx).Value = CStr(Nz(NewValue, ""))
        UpdateValue = True
    End If
End Function

'====================================================================
' Delete key
'====================================================================
Public Function DeleteKey(key As String) As Boolean
    Dim idx As Long
    Dim i As Long
    
    idx = FindKey(key)
    If idx = -1 Then Exit Function
    
    ' Shift items down
    For i = idx To mCount - 2
        mItems(i) = mItems(i + 1)
    Next i
    
    mCount = mCount - 1
    DeleteKey = True
End Function

'====================================================================
' Clear all
'====================================================================
Public Sub Clear()
    ReDim mItems(0 To 9)
    mCount = 0
End Sub

'====================================================================
' Pretty print
'====================================================================
Public Function ToPretty(Optional Indent As Long = 0) As String
    ToPretty = BuildPretty(Indent)
End Function

'====================================================================
' To JSON
'====================================================================
Public Function ToJSON() As String
    ToJSON = BuildJSON(0)
End Function

'====================================================================
' Serialize to string
'====================================================================
Public Function ToRaw() As String
    Dim result As String
    Dim i As Long
    
    For i = 0 To mCount - 1
        If i > 0 Then result = result & Chr(2)
        
        ' For objects, encode their raw data to prevent Chr(1) and Chr(2) conflicts
        If mItems(i).ItemType = "O" Then
            result = result & mItems(i).key & Chr(1) & mItems(i).ItemType & Chr(1) & EncodeForStorage(mItems(i).Value)
        Else
            result = result & mItems(i).key & Chr(1) & mItems(i).ItemType & Chr(1) & mItems(i).Value
        End If
    Next i
    
    ToRaw = result
End Function

Private Function EncodeForStorage(ByVal s As String) As String
    ' Replace Chr(1) and Chr(2) with safe control characters
    s = Replace(s, Chr(1), Chr(3))
    s = Replace(s, Chr(2), Chr(4))
    EncodeForStorage = s
End Function

Private Function DecodeFromStorage(ByVal s As String) As String
    ' Restore Chr(1) and Chr(2) from safe control characters
    s = Replace(s, Chr(3), Chr(1))
    s = Replace(s, Chr(4), Chr(2))
    DecodeFromStorage = s
End Function

'====================================================================
' Deserialize from string
'====================================================================
Public Sub Parse(RawData As String)
    Clear
    
    If Len(RawData) = 0 Then Exit Sub
    
    Dim pairs() As String
    Dim parts() As String
    Dim i As Long
    
    pairs = Split(RawData, Chr(2))
    
    For i = 0 To UBound(pairs)
        If Len(pairs(i)) = 0 Then GoTo NextPair
        
        parts = Split(pairs(i), Chr(1), 3)  ' Limit to 3 parts: key, type, value
        
        If UBound(parts) >= 2 Then
            ' For objects, decode their stored data
            If parts(1) = "O" Then
                AddItem parts(0), DecodeFromStorage(parts(2)), parts(1)
            Else
                AddItem parts(0), parts(2), parts(1)
            End If
        End If
        
NextPair:
    Next i
End Sub

'====================================================================
' SQL-safe
'====================================================================
Public Function ToSQLSafe() As String
    ToSQLSafe = Replace(ToRaw(), "'", "''")
End Function

'====================================================================
' PRIVATE HELPERS
'====================================================================
Private Sub AddItem(key As String, Value As String, ItemType As String)
    If mCount > UBound(mItems) Then
        ReDim Preserve mItems(0 To UBound(mItems) + 10)
    End If
    
    mItems(mCount).key = key
    mItems(mCount).Value = Value
    mItems(mCount).ItemType = ItemType
    mCount = mCount + 1
End Sub

Private Function FindKey(key As String) As Long
    Dim i As Long
    For i = 0 To mCount - 1
        If mItems(i).key = key Then
            FindKey = i
            Exit Function
        End If
    Next i
    FindKey = -1
End Function

Private Function BuildPretty(Indent As Long) As String
    Dim result As String
    Dim sp As String
    Dim i As Long
    
    sp = String(Indent * 2, " ")
    
    If mCount = 0 Then
        BuildPretty = sp & "{}"
        Exit Function
    End If
    
    result = sp & "{" & vbCrLf
    
    For i = 0 To mCount - 1
        result = result & sp & "  """ & mItems(i).key & """: "
        
        Select Case mItems(i).ItemType
            Case "V"
                If IsNumeric(mItems(i).Value) Or LCase(mItems(i).Value) = "true" Or _
                   LCase(mItems(i).Value) = "false" Or LCase(mItems(i).Value) = "null" Then
                    result = result & mItems(i).Value
                Else
                    result = result & """" & mItems(i).Value & """"
                End If
                
            Case "A"
                result = result & "[ "
                If Len(mItems(i).Value) > 0 Then
                    result = result & """" & Replace(mItems(i).Value, "|", """, """) & """"
                End If
                result = result & " ]"
                
            Case "O"
                Dim childObj As New FastJSON
                childObj.Parse mItems(i).Value
                result = result & vbCrLf & childObj.ToPretty(Indent + 1)
        End Select
        
        If i < mCount - 1 Then result = result & ","
        result = result & vbCrLf
    Next i
    
    result = result & sp & "}"
    BuildPretty = result
End Function

Private Function BuildJSON(Indent As Long) As String
    Dim result As String
    Dim sp As String
    Dim i As Long
    
    sp = String(Indent * 2, " ")
    
    If mCount = 0 Then
        BuildJSON = sp & "{}"
        Exit Function
    End If
    
    result = sp & "{" & vbCrLf
    
    For i = 0 To mCount - 1
        result = result & sp & "  """ & EscapeJSON(mItems(i).key) & """: "
        
        Select Case mItems(i).ItemType
            Case "V"
                If IsNumeric(mItems(i).Value) Or LCase(mItems(i).Value) = "true" Or _
                   LCase(mItems(i).Value) = "false" Or LCase(mItems(i).Value) = "null" Then
                    result = result & mItems(i).Value
                Else
                    result = result & """" & EscapeJSON(mItems(i).Value) & """"
                End If
                
            Case "A"
                result = result & "["
                If Len(mItems(i).Value) > 0 Then
                    Dim arrParts() As String
                    Dim j As Long
                    arrParts = Split(mItems(i).Value, "|")
                    For j = 0 To UBound(arrParts)
                        If j > 0 Then result = result & ", "
                        result = result & """" & EscapeJSON(arrParts(j)) & """"
                    Next j
                End If
                result = result & "]"
                
            Case "O"
                Dim childObj As New FastJSON
                childObj.Parse mItems(i).Value
                result = result & vbCrLf & childObj.BuildJSON(Indent + 1)
        End Select
        
        If i < mCount - 1 Then result = result & ","
        result = result & vbCrLf
    Next i
    
    result = result & sp & "}"
    BuildJSON = result
End Function

Private Function EscapeJSON(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJSON = s
End Function

