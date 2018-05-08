VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADT_Heap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSize As Long
Private pArr As Variant

'======================================================================================================================================
'   Initialize Head Properties
'======================================================================================================================================
Sub InitHeap()
    pSize = 0       'size
    ReDim pArr(100)  'capacity
End Sub

'======================================================================================================================================
'   Header Properties
'======================================================================================================================================

Private Function Compare(Val1 As String, Val2 As String) As Integer

    Compare = 0
    
    If (IsNumeric(Val1) = True) And (IsNumeric(Val2) = True) Then 'compare two numbers
        If CDbl(Val1) > CDbl(Val2) Then Compare = 1
        If CDbl(Val1) < CDbl(Val2) Then Compare = -1
        
    ElseIf IsNumeric(Val1) Then Compare = -1
    ElseIf IsNumeric(Val2) Then Compare = 1
    
    Else 'compare two strings
        Compare = StrComp(Val1, Val2, vbBinaryCompare)
    End If
        
End Function

Sub ConsolePrint()

    Dim i As Long

    For i = 0 To pSize Step 1
        Debug.Print pArr(i)
    Next i

End Sub



'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+
'+                                                      CLASS PROPERTIES
'+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'======================================================================================================================================
'   NUMBER OF HELD VALUES IN ARRAY
'======================================================================================================================================
Public Property Get GetSize() As Long
    Size = pSize
End Property

'======================================================================================================================================
'   ARRAY HOLDING VALUES
'======================================================================================================================================
Public Property Get GetArray() As Variant
    Arr = pArr
End Property


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+
'+                                                      CLASS METHODS
'+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'======================================================================================================================================
'   ADD TO ARRAY
'======================================================================================================================================
Private Sub AddToArray(Value As String)
    
    If pSize = UBound(pArr) Then
        ReDim Preserve pArr(0 To (2 * pSize))
    End If
    
    pArr(pSize) = Value
    pSize = pSize + 1

End Sub

'======================================================================================================================================
'   REMOVE FROM ARRAY AT POSITION
'======================================================================================================================================
Sub RemovePos(pos As Long)
    
    While (pos < pSize)
        pArr(pos) = pArr(pos + 1)
        pos = pos + 1
    Wend
    
    If (pos >= 0 And pos < pSize) Then
        pSize = pSize - 1
    End If

End Sub

'======================================================================================================================================
'   GET VALUE AT POSITION
'======================================================================================================================================
Function ValueAtPos(pos As Long) As String
    ValueAtPos = pArr(pos)
End Function


'======================================================================================================================================
'   EXCHANGE VALUES
'======================================================================================================================================
Private Sub ExchangePos(i As Long, j As Long)
    
    Dim temp As String
    
    temp = pArr(i)
    pArr(i) = pArr(j)
    pArr(j) = temp
    
End Sub


'======================================================================================================================================
'   Smaller of Two Positions
'======================================================================================================================================
Private Function SmallerPosition(i As Long, j As Long) As Long

    If Compare(ValueAtPos(i), ValueAtPos(j)) = -1 Then
        SmallerPosition = i
        Exit Function
    End If
    
    SmallerPosition = j

End Function


'======================================================================================================================================
'   PUT IN ARRAY AT LOCATION
'======================================================================================================================================
Private Sub PutArray(i As Long, Value As String)
    pArr(i) = Value
End Sub


'======================================================================================================================================
'   GET MIN HEAP
'======================================================================================================================================
Private Function GetMinHeap() As String
    GetMinHeap = ValueAtPos(0)
End Function


'======================================================================================================================================
'   ADD TO HEAP
'======================================================================================================================================
Sub Add(Value As String)
    
    Dim parent As Long, _
        pos As Long
    
    AddToArray Value 'node added to end
    pos = pSize - 1  'position of last value
    
    While (pos <> 0)
        parent = (pos - 1) / 2
        
        If Compare(ValueAtPos(pos), ValueAtPos(parent)) = -1 Then
            ExchangePos parent, pos
            pos = parent
        Else
            Exit Sub
        End If
    Wend

End Sub


Sub AddArray(ArrData As Variant)

    Dim i As Long
    
    For i = LBound(ArrData) To UBound(ArrData)
        Add CStr(ArrData(i))
    Next i

End Sub



'======================================================================================================================================
'   ADJUST HEAP TO MAINTAIN HEAP PROPERTY
'======================================================================================================================================
Private Sub AdjustHeap(max As Long, pos As Long)
    
    Dim leftChild As Long, _
        rightChild As Long, _
        smallestChild As Long

    leftChild = (2 * pos) + 1
    rightChild = (2 * pos) + 2
    smallestChild = 0
    
    If rightChild < max Then
        smallestChild = SmallerPosition(rightChild, leftChild)
        
        If Compare(ValueAtPos(pos), ValueAtPos(smallestChild)) = 1 Then
            ExchangePos pos, smallestChild
            AdjustHeap max, smallestChild
        End If
    
    ElseIf leftChild < max Then
        
        If Compare(ValueAtPos(pos), ValueAtPos(leftChild)) = 1 Then
            ExchangePos pos, leftChild
            AdjustHeap max, leftChild
        End If
    
    End If

End Sub


'======================================================================================================================================
'   REMOVE MIN HEAP
'======================================================================================================================================
Private Sub RemoveHeapMin()

    Dim last As Long
    
    last = pSize - 1
    PutArray 0, ValueAtPos(last)
    RemovePos last
    AdjustHeap last, 0

End Sub


'======================================================================================================================================
'   BUILD HEAP
'======================================================================================================================================
Private Sub BuildHeap()

    Dim index As Long
    
    index = (pSize / 2) - 1
    
    ExchangePos 0, pSize - 1
    
    While index >= 0
        AdjustHeap pSize, index
        index = index - 1
    Wend


End Sub


'======================================================================================================================================
'   SORT HEAP
'======================================================================================================================================

Sub Sort()

    Dim index As Long
    
    BuildHeap
    index = pSize - 1
    
    While index > 0
        ExchangePos 0, index
        AdjustHeap index, 0
        index = index - 1
    Wend

End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+
'+                                                      Utility Methods
'+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'======================================================================================================================================
'   Print to Sheet
'======================================================================================================================================
Sub PrintToSheet(CellLoc As Range, Optional LeftToRight As Boolean)
    Dim i As Long
    
    i = 0
    
    While i < pSize
        CellLoc = ValueAtPos(i)
        If (LeftToRight) Then
            Set CellLoc = CellLoc.Offset(0, 1)
        Else
            Set CellLoc = CellLoc.Offset(1, 0)
        End If
        i = i + 1
    Wend
    
    
End Sub
