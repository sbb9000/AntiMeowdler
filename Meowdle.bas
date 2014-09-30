Attribute VB_Name = "Meowdle"
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetLastInputInfo Lib "user32" (pLII As Any) As Long

Public BumpTime As Double       ' TimeSerial value at which the scheduled bump will be triggered
Public RefreshTime As Double    ' TimeSerial value at which the scheduled form refresh will be triggered
Public NapTime As Double        ' TimeSerial value at which the scheduled nap will be triggered

Public BumpCount As Long        ' Running tally of bumps
Public MeowTime As frmMeow      ' Instance of the userform

Public BumpIntervalMinutes As Integer   ' Number of minutes added to the current TimeSerial value to calculate BumpTime
Public NapPeriodHours As Integer        ' Number of hours added to the current TimeSerial value to calculate NapTime
Public Const cStatusIntervalSeconds = 1 ' Number of seconds added to the current TimeSerial value to calculate RefreshTime

Public Const cBumpIt = "BumpIt"         ' String that names the procedure to execute when the clock strikes BumpTime
Public Const cStatusSet = "StatusSet"   ' String that names the procedure to execute when the clock strikes RefreshTime
Public Const cNap = "CatNap"            ' String that names the procedure to execute when the clock strikes NapTime

Private Type LastInputInformation
    cbSize As Long
    dwTime As Long
End Type

Public Sub QueueBump()
    BumpTime = Now + TimeSerial(0, BumpIntervalMinutes, 0)
    Application.OnTime EarliestTime:=BumpTime, Procedure:=cBumpIt, Schedule:=True
End Sub
Public Sub QueueStatusUpdate()
    RefreshTime = Now + TimeSerial(0, 0, cStatusIntervalSeconds)
    Application.OnTime EarliestTime:=RefreshTime, Procedure:=cStatusSet, Schedule:=True
End Sub
Public Sub QueueNap()
    NapTime = Now + TimeSerial(NapPeriodHours, 0, 0)
    Application.OnTime EarliestTime:=NapTime, Procedure:=cNap, Schedule:=True
End Sub

Public Sub CancelBump()
    On Error Resume Next
        Application.OnTime EarliestTime:=BumpTime, Procedure:=cBumpIt, Schedule:=False
    On Error GoTo 0
End Sub
Public Sub EndStatusUpdates()
    On Error Resume Next
        Application.OnTime EarliestTime:=RefreshTime, Procedure:=cStatusSet, Schedule:=False
    On Error GoTo 0
End Sub
Public Sub CancelNap()
    On Error Resume Next
        Application.OnTime EarliestTime:=NapTime, Procedure:=cNap, Schedule:=False
    On Error GoTo 0
End Sub

Public Function GetUsersIdleTime() As Long
    Dim LII As LastInputInformation
    LII.cbSize = Len(LII)
    Call GetLastInputInfo(LII)
    GetUsersIdleTime = FormatNumber((GetTickCount() - LII.dwTime) / 1000, 2)
End Function

Sub BumpIt()
    If BumpIntervalMinutes > 0 Then
        If GetUsersIdleTime() >= ((BumpIntervalMinutes * 60) - 1) Then
            SendKeys "+"
            BumpCount = BumpCount + 1
        End If
        Call QueueBump
    End If
End Sub
Private Sub StatusSet()
    If RefreshTime > 0 Then
        MeowTime.txtElapsed.Value = CStr(Format(((GetUsersIdleTime - (GetUsersIdleTime Mod 60)) / 60), "00")) & ":" & CStr(Format(GetUsersIdleTime Mod 60, "00"))
        MeowTime.txtPounces.Value = BumpCount
        Call QueueStatusUpdate
    End If
End Sub
Public Sub Nap()
        Call CancelBump
        Call EndStatusUpdates
        MeowTime.btnMAOU.Enabled = False
        MeowTime.btnMeow.Enabled = True
        MeowTime.chkCatNap.Enabled = True
        MeowTime.txtThreshold.Enabled = True
End Sub

Public Sub MeowMeow()
    If Not MeowTime Is Nothing Then
        ' nothing
        If MeowTime.btnMAOU.Enabled Then
            Call QueueStatusUpdate
        End If
    Else
        Set MeowTime = New frmMeow
        BumpCount = 0
    End If
    MeowTime.Show
End Sub
Public Sub Miew(Optional ByVal bKill As Boolean)
    MeowTime.Hide
    If MeowTime.btnMeow.Enabled Or bKill Then
        Call Nap
        Set MeowTime = Nothing
    Else
        Call EndStatusUpdates
    End If
End Sub
Public Sub AddButton()
    Call RemoveButton
    Dim btn As CommandBarControl
    Set btn = Application.CommandBars("Worksheet Menu Bar").FindControl(ID:=30007).CommandBar.Controls.Add( _
            Office.MsoControlType.msoControlButton, Before:=1, Temporary:=False)
    With btn
        .Style = Office.MsoButtonStyle _
        .msoButtonIconAndCaption
        .Caption = ""
        '.FaceId = 65
        .Picture = frmMeow.imgIcon.Picture
        .OnAction = ThisWorkbook.Name & "!MeowMeow"
    End With
End Sub
Public Sub RemoveButton()
    Dim I As Integer
    With Application.CommandBars("Worksheet Menu Bar").FindControl(ID:=30007).CommandBar
        For I = .Controls.Count To 1 Step -1
            If .Controls(I).Caption Like "" Then
                If Replace(.Controls(I).OnAction, "'", "") Like "*!Meow*" Then
                    .Controls(I).Delete
                End If
            End If
        Next I
    End With
End Sub
