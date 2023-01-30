VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStart 
   Caption         =   "��ܶ}�l�ɶ�"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   4815
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1860
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�з���"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   138674177
      CurrentDate     =   44930
   End
   Begin VB.Label lblQueryDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "Query Date:"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msTitle As String = "frmStart."

Private Sub cmdStart_Click()

On Error GoTo ErrorHandler
    
    Dim i, j, k As Integer
    Dim iSheet As Integer
    Dim iTop As Integer '�Ϫ��l��m
    
    Dim sTitle As String
    Dim sFileName As String
    Dim sSavePath As String
    Dim sSQL As String
    Dim arrsColumn() As String '��ƪ����W��
    Dim sEQPID As String
    Dim sTime As String
    Dim sLotID As String '�帹�W��
    Dim sLblName As String
    Dim sFixTable As String '��m�����
    Dim sTemErr As String '�ūײ��`
    
    Dim oExlApp As New Excel.Application
    Dim oWorkBook As Workbook
    Dim oSheet As Worksheet
    Dim oChart As Chart
    Dim cEQP As Collection '���x�����X
    Dim cLot As Collection '�帹�����X
    Dim cEDC As Collection '�帹�����q���
    Dim dic As New Dictionary '����ūײ��`��
    
    sTitle = msTitle & "cmdStart_Click"
    LogMsg sTitle & ": Enter"
    
    '�ˬd�s�ɥت����L�ɮ�,���N�R���Ф@�ӷs��
    sFileName = "EDC.xlsx"
    sSavePath = "C:\Users\user\Desktop\"
    If goFSO.FileExists(sSavePath & sFileName) Then goFSO.DeleteFile sSavePath & sFileName
    Set oWorkBook = oExlApp.Workbooks.Add
    
    With oExlApp
        .Visible = True '�����i��
        .WindowState = xlMaximized '�����̤j��
        .DisplayAlerts = False 'ĵ�i��������
        
        '�w�]���T��Sheet,�R��ѤU�@��
        For i = 1 To oWorkBook.Worksheets.Count - 1
            oWorkBook.Worksheets(1).Delete
        Next
        
        .Sheets(1).Name = "�Ϫ�"
        .ActiveWindow.DisplayGridlines = False '������u
        
        '�b��Ʈw��X�Ҧ����x
        sSQL = "SELETE DISTINCT EQPID"
        sSQL = sSQL & " FORM PCS_EQP"
        sSQL = sSQL & " WHERE STEPNAME = 'Slicing'"
        sSQL = sSQL & " AND EQPID LIKE 'Slicing%'"
        sSQL = sSQL & " ORDER BY EQPID"
        Set cEQP = PDBDataBase(sSQL)
        
        iSheet = 2 '�q�ĤG��sheet�}�l����
        iTop = 10
        
        LogMsg "�j�M�ɶ�:" & DTPicker1.Value
        
        sTime = Format$(DTPicker1.Value, "yyyymmdd")
        
        arrsColumn = Array("TABLE_POSITION", "TEMPERATURE_WORKING", "LEFT_MAIN_GUIDE", "RIGHT_MAIN_GUIDE", "SLURRY_IN_TEMP", "WG_R_OUT_TEMP", "R_MAIN_ROLLER_TEMP", "L_MAIN_ROLLER_TEMP")
        
        For i = 1 To cEQP.Count
            sEQPID = cEQP(i).Item(1)
            LogMsg sEQPID
            
            '������x�b�ҿ諸����̷s��Lot
            sSQL = "SELECT LOT_ID"
            sSQL = sSQL & " FROM AI_SL_PROSERVER"
            sSQL = sSQL & " WHERE EQUIP = " & Q(sEQPID)
            sSQL = sSQL & " AND COLLECT_TIME LIKE '" & sTime & "%'"
            sSQL = sSQL & " AND TABLE_START = 1"
            sSQL = sSQL & " AND TABLE_STOP = 0"
            sSQL = sSQL & " AND ROWNUM = 1"
            sSQL = sSQL & " AND LOT_ID IS NOT NULL"
            sSQL = sSQL & " ORDER BY COLLECT_TIME DESC"
            Set cLot = PDBDataBase(sSQL)
            
            If cLot.Count = 0 Then
                LogMsg "�LLot" & vbCrLf & sSQL
                GoTo NextEQP
            End If
            
            sLotID = cLot(1).Item(1)
            sLblName = sLotID & "_" & Right$(sEQPID, 3)
            LogMsg sLotID
            
            '���o�帹���˴����
            sSQL = "SELECT TABLE_POSITION ,TEMPERATURE_WORKING ,LEFT_MAIN_GUIDE ,RIGHT_MAIN_GUIDE ,SLURRY_IN_TEMP ,WG_R_OUT_TEMP ,R_MAIN_ROLLER_TEMP,L_MAIN_ROLLER_TEMP"
            sSQL = sSQL & " FROM AI_SL_PROSERVER"
            sSQL = sSQL & " WHERE LOT_ID = " & Q(sLotID)
            sSQL = sSQL & " AND TABLE_START=1"
            sSQL = sSQL & " AND TABLE_STOP = 0"
            sSQL = sSQL & " ORDER BY COLLECT_TIME"
            Set cEDC = PDBDataBase(sSQL)
            
            If cEDC.Count = 0 Then
                LogMsg "Lot�LEDC" & vbCrLf & sSQL
                GoTo NextEQP
            End If
            
            '�s�Wsheet
            .Sheets.Add after:=.Sheets(iSheet - 1)
            .Sheets(iSheet).Name = sLblName
            
            Set oSheet = .Worksheets(sLblName)
            
            For j = 0 To UBound(arrsColumn)
                oSheet.Cells(1, j + 1) = arrsColumn(j) '��J���
                
                For k = 1 To cEDC.Count
                    oSheet.Cells(k + 1, j + 1) = cEDC.Item(k).Item(arrsColumn(j)) '�v����J���
                    
                    If arrsColumn(j) = "R_MAIN_ROLLER_TEMP" Or arrsColumn = "L_MAIN_ROLLER_TEMP" Then
                        sFixTable = Fix(cEDC.Item(k).Item("TABLE_POSITIONM"))
                        
                        If sFixTable > 0 Then '�p��1������,�]����}�l�[�u�ū��٤�í�w
                            '���o�Cmm���̤j��
                            If dic.Exists(sFixTable) Then dic.Remove (sFixTable)
                            dic.Add sFixTable, cEDC.Item(k).Item(arrsColumn(j))
                        End If
                        
                        '�̫�@���S��B�z
                        If k = cEDC.Count Then
                            If dic(CStr(sFixTable)) - dic(CStr(sFixTable - 1)) < -0.5 Then
                                sTemErr = sTemErr & TempErr(sLblName, cEDC.Item(k).Item("TABLE_POSITION"), arrsColumn(j))
                            End If
                        End If
                        
                        If dic.Count < 3 Then GoTo Continue
                        
                        '����Cmm�ū׬O�_�C��-0.5
                        If dic(CStr(sFixTable - 1)) - dic(CStr(sFixTable - 2)) < -0.5 Then
                            sTemErr = sTemErr & TempErr(sLblName, cEDC.Item(k).Item("TABLE_POSITION"), arrsColumn(j))
                        End If
                        
                        dic.Remove (CStr(sFixTable - 2))
                        
                    End If
Continue:
                    DoEvents
                    
                Next
                
                dic.RemoveAll
                
            Next
            
            oSheet.Cells.EntireColumn.AutoFill '�۰ʽվ���e
            LogMsg "EDC��J����"
            
            If cEDC.Count = 1 Then
                LogMsg sLotID & "-�u���@�����,�L�k�X�Ϫ�"
                iSheet = iSheet + 1
                GoTo NextEQP
            End If
            
            '���i�Ϫ�
            For j = 1 To 5
                Set oChart = oExlApp.Sheets("�Ϫ�").ChartObjects.Add(Left:=(j - 1 * 360), Top:=iTop, Width:=350, Height:=250).Chart
                
                If oChart Is Nothing Or oSheet Is Nothing Then Exit For
                
                Call SetChart(j, oChart, , cEDC.Count, oSheet)
                
                With oChart
                    .Legend.Position = xlLegendPositionBottom '���u�W�٦�m
                    
                    '�]�wX�b�̤j�ȳ̤p��
                    .Axes(xlCategory).MinimumScale = 0
                    .Axes(xlCategory).MaximumScale = 300
                End With
            Next
            
            iSheet = iSheet + 1
            iTop = iTop + 260
            
            LogMsg "�Ϫ�ͦ�����"
            
NextEQP:
        Next
        
        If sTemErr <> "" Then
            Call HandleWithMail(sTemErr)
        End If
    End With
    
    oExlApp.DisplayAlerts = True '�}��ĵ�i
    
    oWorkBook.SaveAs sSavePath & sFileName '�s�ɦܫ��w���|
    oWorkBook.Close True
    
    
    
    GoTo Finish

ErrorHandler:
    Dim lErrNo As Long
    Dim sErrMsg As String

    lErrNo = Err.Number
    sErrMsg = Err.Description

    LogMsg sTitle & " Error-" & lErrNo & ":" & sErrMsg
    MsgBox lErrNo & "-" & sErrMsg, vbOKOnly, sTitle
    Err.Clear

    If Not oWorkBook Is Nothing Then oWorkBook.Close False
    
Finish:
    If Not oExlApp Is Nothing Then oExlApp.Quit
    
    LogMsg sTitle & " : Exit"
End Sub

Private Function TempErr(sLblName As String, sTablePos As String, sColName As String) As String
'����Cmm�ū׬O�_�C��-0.5
On Error GoTo Err

    Dim sTitle As String
    Dim sTe As String
    
    sTitle = msTitle & "TempErr"
    
    TempErr = sLblName & " �� " & sTablePos
    
    If Left$(sColName, 1) = "R" Then
        TempErr = TempErr & " ��m�k���D�ɽ��b�ӷū׵o�ͧC��[-0.5��]�����`"
    Else
        TempErr = TempErr & " ��m�����D�ɽ��b�ӷū׵o�ͧC��[-0.5��]�����`"
    End If
    
    Exit Function

Err:
    Dim lErrNo As Long
    Dim sErrMsg As String
    
    lErrNo = Err.Number
    sErrMsg = Err.Description
    
    LogMsg sTitle & " Error-" & lErrNo & ":" & sErrMsg
    MsgBox lErrNo & "-" & sErrMsg, , sTitle
    Err.Clear
End Function

Private Sub SetChart(iChart As Integer, oChart As Chart, lRow As Long, oSheet As Worksheet)
'�]�w�Ϫ����
On Error GoTo Err
    
    Dim sTitle As String
    
    sTitle = msTitle & "SetChart"
    
    With oChart
        .ChartType = xlXYScatterSmoothNoMarkers '�]�w�Ϫ�����
        
        Select Case iChart
        
            Case 1
                .SetSourceData oSheet.Range("A2:A" & lRow, "B2:B" & lRow) '����ƽd��
                .SeriesCollection(1).Name = oSheet.Cells(1, 2) '�]�w���u�W��
                
                '�]�w�Ϫ�W��
                .HasTitle = True
                .ChartTitle.Text = oSheet.Name & " �����ū�(10mm��m)"
                
                '�]�wY�b�̤j�̤p��
                .Axes(xlValue).MinimumScale = 20
                .Axes(xlValue).MaximumScale = 34
                
            Case 2
                .SetSourceData oSheet.Range("A2:A" & lRow, "C2:D" & lRow) '����ƽd��
                
                '�]�w���u�W��
                .SeriesCollection(2).Name = oSheet.Cells(1, 3)
                .SeriesCollection(3).Name = oSheet.Cells(1, 4)
                .SeriesCollection(1).Delete
                
                '�]�w�Ϫ�W��
                .HasTitle = True
                .ChartTitle.Text = oSheet.Name & "�D�ɽ��ܦ�"
                
                '�]�wY�b�̤j�̤p��
                .Axes(xlValue).MinimumScale = -20
                .Axes(xlValue).MaximumScale = 10
                
            Case 3
                .SetSourceData oSheet.Range("A2:A" & lRow, "E2:E" & lRow) '����ƽd��
                
                '�]�w���u�W��
                .SeriesCollection(4).Name = oSheet.Cells(1, 5)
                .SeriesCollection(3).Delete
                .SeriesCollection(2).Delete
                .SeriesCollection(1).Delete
                
                '�]�w�Ϫ�W��
                .HasTitle = True
                .ChartTitle.Text = oSheet.Name & "SLURRY��ڷū�"
                
                '�]�wY�b�̤j�̤p��
                .Axes(xlValue).MinimumScale = 21
                .Axes(xlValue).MaximumScale = 25
                
            Case 4
                .SetSourceData oSheet.Range("A2:A" & lRow, "F2:G" & lRow) '����ƽd��
                
                '�]�w���u�W��
                .SeriesCollection(5).Name = oSheet.Cells(1, 6)
                .SeriesCollection(6).Name = oSheet.Cells(1, 7)
                .SeriesCollection(4).Delete
                .SeriesCollection(3).Delete
                .SeriesCollection(2).Delete
                .SeriesCollection(1).Delete
                
                '�]�w�Ϫ�W��
                .HasTitle = True
                .ChartTitle.Text = oSheet.Name & "�D�ɽ��ū�"
                
                '�]�wY�b�̤j�̤p��
                .Axes(xlValue).MinimumScale = 21
                .Axes(xlValue).MaximumScale = 25
                
            Case 5
                .SetSourceData oSheet.Range("A2:A" & lRow, "H2:I" & lRow) '����ƽd��
                
                '�]�w���u�W��
                .SeriesCollection(7).Name = oSheet.Cells(1, 8)
                .SeriesCollection(8).Name = oSheet.Cells(1, 9)
                .SeriesCollection(6).Delete
                .SeriesCollection(5).Delete
                .SeriesCollection(4).Delete
                .SeriesCollection(3).Delete
                .SeriesCollection(2).Delete
                .SeriesCollection(1).Delete
                
                '�]�w�Ϫ�W��
                .HasTitle = True
                .ChartTitle.Text = oSheet.Name & "�D�ɽ��b�ӷū�"
                
                '�]�wY�b�̤j�̤p��
                .Axes(xlValue).MinimumScale = 20
                .Axes(xlValue).MaximumScale = 40
                
        End Select
    End With
    
    Exit Sub
    
Err:
    Dim lErrNo As Long
    Dim sErrMsg As String
    
    lErrNo = Err.Number
    sErrMsg = Err.Description
    
    LogMsg sTitle & " Error-" & lErrNo & ":" & sErrMsg
    MsgBox lErrNo & "-" & sErrMsg, , sTitle
    Err.Clear
End Sub

Private Sub HandleWithMail(sTemErr As String)
    '�H�H�������H��������
End Sub
