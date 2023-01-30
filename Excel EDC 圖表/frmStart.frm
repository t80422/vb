VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStart 
   Caption         =   "選擇開始時間"
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
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "標楷體"
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
         Name            =   "標楷體"
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
      BackStyle       =   0  '透明
      Caption         =   "Query Date:"
      BeginProperty Font 
         Name            =   "標楷體"
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
    Dim iTop As Integer '圖表初始位置
    
    Dim sTitle As String
    Dim sFileName As String
    Dim sSavePath As String
    Dim sSQL As String
    Dim arrsColumn() As String '資料表欄位名稱
    Dim sEQPID As String
    Dim sTime As String
    Dim sLotID As String '批號名稱
    Dim sLblName As String
    Dim sFixTable As String '位置的整數
    Dim sTemErr As String '溫度異常
    
    Dim oExlApp As New Excel.Application
    Dim oWorkBook As Workbook
    Dim oSheet As Worksheet
    Dim oChart As Chart
    Dim cEQP As Collection '機台的集合
    Dim cLot As Collection '批號的集合
    Dim cEDC As Collection '批號的測量資料
    Dim dic As New Dictionary '比較溫度異常用
    
    sTitle = msTitle & "cmdStart_Click"
    LogMsg sTitle & ": Enter"
    
    '檢查存檔目的有無檔案,有就刪掉創一個新的
    sFileName = "EDC.xlsx"
    sSavePath = "C:\Users\user\Desktop\"
    If goFSO.FileExists(sSavePath & sFileName) Then goFSO.DeleteFile sSavePath & sFileName
    Set oWorkBook = oExlApp.Workbooks.Add
    
    With oExlApp
        .Visible = True '視窗可見
        .WindowState = xlMaximized '視窗最大化
        .DisplayAlerts = False '警告視窗關閉
        
        '預設有三個Sheet,刪到剩下一個
        For i = 1 To oWorkBook.Worksheets.Count - 1
            oWorkBook.Worksheets(1).Delete
        Next
        
        .Sheets(1).Name = "圖表"
        .ActiveWindow.DisplayGridlines = False '關閉格線
        
        '在資料庫找出所有機台
        sSQL = "SELETE DISTINCT EQPID"
        sSQL = sSQL & " FORM PCS_EQP"
        sSQL = sSQL & " WHERE STEPNAME = 'Slicing'"
        sSQL = sSQL & " AND EQPID LIKE 'Slicing%'"
        sSQL = sSQL & " ORDER BY EQPID"
        Set cEQP = PDBDataBase(sSQL)
        
        iSheet = 2 '從第二個sheet開始塞資料
        iTop = 10
        
        LogMsg "搜尋時間:" & DTPicker1.Value
        
        sTime = Format$(DTPicker1.Value, "yyyymmdd")
        
        arrsColumn = Array("TABLE_POSITION", "TEMPERATURE_WORKING", "LEFT_MAIN_GUIDE", "RIGHT_MAIN_GUIDE", "SLURRY_IN_TEMP", "WG_R_OUT_TEMP", "R_MAIN_ROLLER_TEMP", "L_MAIN_ROLLER_TEMP")
        
        For i = 1 To cEQP.Count
            sEQPID = cEQP(i).Item(1)
            LogMsg sEQPID
            
            '抓取機台在所選的日期最新的Lot
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
                LogMsg "無Lot" & vbCrLf & sSQL
                GoTo NextEQP
            End If
            
            sLotID = cLot(1).Item(1)
            sLblName = sLotID & "_" & Right$(sEQPID, 3)
            LogMsg sLotID
            
            '取得批號的檢測資料
            sSQL = "SELECT TABLE_POSITION ,TEMPERATURE_WORKING ,LEFT_MAIN_GUIDE ,RIGHT_MAIN_GUIDE ,SLURRY_IN_TEMP ,WG_R_OUT_TEMP ,R_MAIN_ROLLER_TEMP,L_MAIN_ROLLER_TEMP"
            sSQL = sSQL & " FROM AI_SL_PROSERVER"
            sSQL = sSQL & " WHERE LOT_ID = " & Q(sLotID)
            sSQL = sSQL & " AND TABLE_START=1"
            sSQL = sSQL & " AND TABLE_STOP = 0"
            sSQL = sSQL & " ORDER BY COLLECT_TIME"
            Set cEDC = PDBDataBase(sSQL)
            
            If cEDC.Count = 0 Then
                LogMsg "Lot無EDC" & vbCrLf & sSQL
                GoTo NextEQP
            End If
            
            '新增sheet
            .Sheets.Add after:=.Sheets(iSheet - 1)
            .Sheets(iSheet).Name = sLblName
            
            Set oSheet = .Worksheets(sLblName)
            
            For j = 0 To UBound(arrsColumn)
                oSheet.Cells(1, j + 1) = arrsColumn(j) '輸入欄位
                
                For k = 1 To cEDC.Count
                    oSheet.Cells(k + 1, j + 1) = cEDC.Item(k).Item(arrsColumn(j)) '逐筆輸入資料
                    
                    If arrsColumn(j) = "R_MAIN_ROLLER_TEMP" Or arrsColumn = "L_MAIN_ROLLER_TEMP" Then
                        sFixTable = Fix(cEDC.Item(k).Item("TABLE_POSITIONM"))
                        
                        If sFixTable > 0 Then '小於1的不抓,因為剛開始加工溫度還不穩定
                            '取得每mm的最大值
                            If dic.Exists(sFixTable) Then dic.Remove (sFixTable)
                            dic.Add sFixTable, cEDC.Item(k).Item(arrsColumn(j))
                        End If
                        
                        '最後一筆特殊處理
                        If k = cEDC.Count Then
                            If dic(CStr(sFixTable)) - dic(CStr(sFixTable - 1)) < -0.5 Then
                                sTemErr = sTemErr & TempErr(sLblName, cEDC.Item(k).Item("TABLE_POSITION"), arrsColumn(j))
                            End If
                        End If
                        
                        If dic.Count < 3 Then GoTo Continue
                        
                        '比較每mm溫度是否低於-0.5
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
            
            oSheet.Cells.EntireColumn.AutoFill '自動調整欄寬
            LogMsg "EDC輸入完畢"
            
            If cEDC.Count = 1 Then
                LogMsg sLotID & "-只有一筆資料,無法出圖表"
                iSheet = iSheet + 1
                GoTo NextEQP
            End If
            
            '五張圖表
            For j = 1 To 5
                Set oChart = oExlApp.Sheets("圖表").ChartObjects.Add(Left:=(j - 1 * 360), Top:=iTop, Width:=350, Height:=250).Chart
                
                If oChart Is Nothing Or oSheet Is Nothing Then Exit For
                
                Call SetChart(j, oChart, , cEDC.Count, oSheet)
                
                With oChart
                    .Legend.Position = xlLegendPositionBottom '曲線名稱位置
                    
                    '設定X軸最大值最小值
                    .Axes(xlCategory).MinimumScale = 0
                    .Axes(xlCategory).MaximumScale = 300
                End With
            Next
            
            iSheet = iSheet + 1
            iTop = iTop + 260
            
            LogMsg "圖表生成完畢"
            
NextEQP:
        Next
        
        If sTemErr <> "" Then
            Call HandleWithMail(sTemErr)
        End If
    End With
    
    oExlApp.DisplayAlerts = True '開啟警告
    
    oWorkBook.SaveAs sSavePath & sFileName '存檔至指定路徑
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
'比較每mm溫度是否低於-0.5
On Error GoTo Err

    Dim sTitle As String
    Dim sTe As String
    
    sTitle = msTitle & "TempErr"
    
    TempErr = sLblName & " 於 " & sTablePos
    
    If Left$(sColName, 1) = "R" Then
        TempErr = TempErr & " 位置右側主導輪軸承溫度發生低於[-0.5度]之異常"
    Else
        TempErr = TempErr & " 位置左側主導輪軸承溫度發生低於[-0.5度]之異常"
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
'設定圖表條件
On Error GoTo Err
    
    Dim sTitle As String
    
    sTitle = msTitle & "SetChart"
    
    With oChart
        .ChartType = xlXYScatterSmoothNoMarkers '設定圖表類型
        
        Select Case iChart
        
            Case 1
                .SetSourceData oSheet.Range("A2:A" & lRow, "B2:B" & lRow) '取資料範圍
                .SeriesCollection(1).Name = oSheet.Cells(1, 2) '設定曲線名稱
                
                '設定圖表名稱
                .HasTitle = True
                .ChartTitle.Text = oSheet.Name & " 晶塊溫度(10mm位置)"
                
                '設定Y軸最大最小值
                .Axes(xlValue).MinimumScale = 20
                .Axes(xlValue).MaximumScale = 34
                
            Case 2
                .SetSourceData oSheet.Range("A2:A" & lRow, "C2:D" & lRow) '取資料範圍
                
                '設定曲線名稱
                .SeriesCollection(2).Name = oSheet.Cells(1, 3)
                .SeriesCollection(3).Name = oSheet.Cells(1, 4)
                .SeriesCollection(1).Delete
                
                '設定圖表名稱
                .HasTitle = True
                .ChartTitle.Text = oSheet.Name & "主導輪變位"
                
                '設定Y軸最大最小值
                .Axes(xlValue).MinimumScale = -20
                .Axes(xlValue).MaximumScale = 10
                
            Case 3
                .SetSourceData oSheet.Range("A2:A" & lRow, "E2:E" & lRow) '取資料範圍
                
                '設定曲線名稱
                .SeriesCollection(4).Name = oSheet.Cells(1, 5)
                .SeriesCollection(3).Delete
                .SeriesCollection(2).Delete
                .SeriesCollection(1).Delete
                
                '設定圖表名稱
                .HasTitle = True
                .ChartTitle.Text = oSheet.Name & "SLURRY實際溫度"
                
                '設定Y軸最大最小值
                .Axes(xlValue).MinimumScale = 21
                .Axes(xlValue).MaximumScale = 25
                
            Case 4
                .SetSourceData oSheet.Range("A2:A" & lRow, "F2:G" & lRow) '取資料範圍
                
                '設定曲線名稱
                .SeriesCollection(5).Name = oSheet.Cells(1, 6)
                .SeriesCollection(6).Name = oSheet.Cells(1, 7)
                .SeriesCollection(4).Delete
                .SeriesCollection(3).Delete
                .SeriesCollection(2).Delete
                .SeriesCollection(1).Delete
                
                '設定圖表名稱
                .HasTitle = True
                .ChartTitle.Text = oSheet.Name & "主導輪溫度"
                
                '設定Y軸最大最小值
                .Axes(xlValue).MinimumScale = 21
                .Axes(xlValue).MaximumScale = 25
                
            Case 5
                .SetSourceData oSheet.Range("A2:A" & lRow, "H2:I" & lRow) '取資料範圍
                
                '設定曲線名稱
                .SeriesCollection(7).Name = oSheet.Cells(1, 8)
                .SeriesCollection(8).Name = oSheet.Cells(1, 9)
                .SeriesCollection(6).Delete
                .SeriesCollection(5).Delete
                .SeriesCollection(4).Delete
                .SeriesCollection(3).Delete
                .SeriesCollection(2).Delete
                .SeriesCollection(1).Delete
                
                '設定圖表名稱
                .HasTitle = True
                .ChartTitle.Text = oSheet.Name & "主導輪軸承溫度"
                
                '設定Y軸最大最小值
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
    '寄信給相關人員的物件
End Sub
