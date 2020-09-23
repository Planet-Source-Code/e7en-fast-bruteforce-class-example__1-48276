VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim sCharSet() As String
Dim sFirstPass As String
Dim sCurrentPass As String
Dim iStartLen As Integer
Dim PassPos() As Integer
Dim sTotalCombo As String
Public Event CombinationsPerSec(Combos As Long)
Public Event TotalCombinations(Combos As String)


Private Declare Function GetTickCount Lib "kernel32" () As Long

Property Let CharacterSet(value As String)
    ReDim sCharSet(Len(value))
    
    For X = 1 To UBound(sCharSet)
        sCharSet(X) = Mid(value, X, 1)
    Next
End Property

Property Get CharacterSet() As String
    Dim sTmp As String

    For X = 1 To UBound(sCharSet)
        sTmp = sTmp & sCharSet(X)
    Next
    
    CharacterSet = sTmp
End Property

Property Let FirstPassword(value As String)
    sFirstPass = value
End Property

Property Get FirstPassword() As String
    FirstPassword = sFirstPass
End Property

Property Let CurrentPassword(value As String)
    sCurrentPass = value
End Property

Property Get CurrentPassword() As String
    CurrentPassword = sCurrentPass
End Property

Property Let StartLength(value As Integer)
    iStartLen = value
End Property

Property Get StartLength() As Integer
    StartLength = iStartLen
End Property
'======================================================================================

Private Sub Combinations()
    Static lComboCount As Long
    Static lComboTime As Long
    lComboCount = lComboCount + 1

    If GetTickCount - lComboTime >= 1000 Then 'if a second has past then

        DoEvents
            RaiseEvent CombinationsPerSec(lComboCount) 'Display # of combinations For that second
            lComboTime = GetTickCount
            sTotalCombo = sTotalCombo + lComboCount 'add to total of combinations
            lComboCount = 0 'reset number of Combinations For the Next second
            RaiseEvent TotalCombinations(sTotalCombo) 'Display # of total combinations
        End If
    End Sub

Sub Initialize()
    sTotalCombo = 0
    If StartLength = 0 Then StartLength = 1

    If FirstPassword = "" Or Len(FirstPassword) < StartLength Then
        CurrentPassword = String(StartLength, sCharSet(1))
    Else
        CurrentPassword = sFirstPass
    End If
        
        ReDim PassPos(Len(CurrentPassword))

        For X = 1 To Len(CurrentPassword)

            For Y = 1 To UBound(sCharSet)

                If Mid(StrReverse(CurrentPassword), X, 1) = sCharSet(Y) Then
                    PassPos(X) = Y
                    Exit For
                End If
            Next
        Next
    
End Sub


Function BruteForce() As String
    Dim sTmp As String
    Dim X As Integer, Y As Integer
    Dim TmpPass() As Integer

    For X = 1 To Len(CurrentPassword)

        If X = Len(CurrentPassword) And PassPos(X) >= UBound(sCharSet) Then
            ReDim TmpPass(UBound(PassPos))

            For Y = 1 To UBound(PassPos)
                TmpPass(Y) = PassPos(Y)
            Next
            ReDim PassPos(UBound(PassPos) + 1)
            PassPos(UBound(PassPos)) = 0

            For Y = 1 To UBound(TmpPass)
                PassPos(Y) = TmpPass(Y)
            Next
        End If

        If PassPos(X) >= UBound(sCharSet) Then
            PassPos(X) = 1
            PassPos(X + 1) = PassPos(X + 1) + 1
            If PassPos(X + 1) <= UBound(sCharSet) Then Exit For
        Else
            PassPos(X) = PassPos(X) + 1
            Exit For
        End If
    Next

    For X = UBound(PassPos) To 1 Step -1
        sTmp = sTmp & sCharSet(PassPos(X))
    Next
    
    Combinations
    CurrentPassword = sTmp
    BruteForce = CurrentPassword
End Function


