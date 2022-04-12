Attribute VB_Name = "Demo"
Option Explicit

Private Const MODULE_NAME As String = "Demo"

Sub mSub00()
    Call XX_PRINT_HISTO_XX(MODULE_NAME, "mSub00")

    On Error GoTo errHandle
    
    Dim start As Double
    start = Timer

    Dim i As Integer
    
    For i = 1 To 5

        Call mSub01
         
        Call mSub02

    Next i

    Call XX_PRINT_HISTO_XX(histEND)
    Exit Sub
errHandle:
    Call FatalError(False, True)
End Sub
Sub mSub01()
    Call XX_PRINT_HISTO_XX(MODULE_NAME, "mSub01")
        Call mSub11

        Call mSub12
    Call XX_PRINT_HISTO_XX(histEND)
End Sub

Sub mSub11()
    Call XX_PRINT_HISTO_XX(MODULE_NAME, "mSub11")
    
    On Error GoTo errHandle

    Call XX_PRINT_HISTO_XX(histEND)
    Exit Sub
    
errHandle:
    Call XX_PRINT_HISTO_XX(histERROR)
End Sub

Sub mSub12()
    Call XX_PRINT_HISTO_XX(MODULE_NAME, "mSub12")
    
    Call XX_PRINT_HISTO_XX(histEXIT)
    Exit Sub
    
    Call XX_PRINT_HISTO_XX(histEND)
End Sub

Sub mSub02()
    Call XX_PRINT_HISTO_XX(MODULE_NAME, "mSub02")

    On Error GoTo errHandle
    
    Call mSub21

    Call XX_PRINT_HISTO_XX(histEND)
    Exit Sub
    
errHandle:
    Call XX_PRINT_HISTO_XX(histERROR)
End Sub

Sub mSub21()
    Call XX_PRINT_HISTO_XX(MODULE_NAME, "mSub21")
        
    On Error GoTo errHandle
    
    Call mSub31
    
    Dim i As Integer: i = "tototo"
    
    Call XX_PRINT_HISTO_XX(histEND)
    Exit Sub
    
errHandle:
    Call XX_PRINT_HISTO_XX(histERROR)
End Sub

Sub mSub31()
    Call XX_PRINT_HISTO_XX(MODULE_NAME, "mSub31")
    
    On Error GoTo errHandle
    
    Dim i As Integer: i = "tototo"
    
    Call XX_PRINT_HISTO_XX(histEND)
    Exit Sub
    
errHandle:
    Call XX_PRINT_HISTO_XX(histERROR)
End Sub




























