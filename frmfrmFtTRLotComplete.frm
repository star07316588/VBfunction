VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFtTRLotComplete 
   Caption         =   "FtTRLotComplete"
   ClientHeight    =   6870
   ClientLeft      =   1125
   ClientTop       =   2670
   ClientWidth     =   9810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   9810
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSelfCheck 
      Caption         =   "TA Self Check."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6210
      TabIndex        =   30
      Top             =   6000
      Width           =   1590
   End
   Begin VB.CommandButton cmdSendVersa3 
      Caption         =   "Send FVI_TR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4245
      TabIndex        =   29
      Top             =   6000
      Width           =   1590
   End
   Begin VB.CommandButton cmdGenLotID 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   27
      Top             =   5880
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.ComboBox cboMergeLots 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdLotCancel 
      Caption         =   "LotCancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4245
      TabIndex        =   19
      Top             =   6420
      Width           =   1590
   End
   Begin VB.ComboBox cboEqpID 
      Height          =   315
      Left            =   1620
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   2055
   End
   Begin VB.Frame fraResult 
      Caption         =   "Packing Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   90
      TabIndex        =   9
      Top             =   2910
      Width           =   9600
      Begin VB.TextBox txtGoodQty 
         Height          =   375
         Left            =   2475
         TabIndex        =   28
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtRemain 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   7065
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1012
         Width           =   1575
      End
      Begin VB.TextBox txtReelCount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2475
         TabIndex        =   3
         Top             =   990
         Width           =   1440
      End
      Begin VB.TextBox txtTotalReeledCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   7065
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   1575
      End
      Begin VB.TextBox txtFail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   7065
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1485
         Width           =   1575
      End
      Begin VB.TextBox txtTRQty 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2475
         TabIndex        =   2
         Top             =   510
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Box/Reel Count (箱/捲) 數:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   17
         Top             =   1035
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Packing Count (已包裝總數) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   4050
         TabIndex        =   13
         Top             =   585
         Width           =   2955
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PackingQty (每箱/捲) 數:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   345
         TabIndex        =   12
         Top             =   555
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FailQty (Fail總數) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   5445
         TabIndex        =   11
         Top             =   1530
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RemainQty (剩餘數量) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   5055
         TabIndex        =   10
         Top             =   1035
         Width           =   1950
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel(放棄)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8100
      TabIndex        =   5
      Top             =   6420
      Width           =   1590
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK(確定)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6210
      TabIndex        =   4
      Top             =   6420
      Width           =   1590
   End
   Begin VB.Frame fraLotList 
      Caption         =   " Lot List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   90
      TabIndex        =   8
      Top             =   540
      Width           =   9585
      Begin FPSpreadADO.fpSpread spdLotList 
         Height          =   1815
         Left            =   225
         TabIndex        =   1
         Top             =   405
         Width           =   9180
         _Version        =   524288
         _ExtentX        =   16193
         _ExtentY        =   3201
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   13
         MaxRows         =   5
         Protect         =   0   'False
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmfrmFtTRLotComplete.frx":0000
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame fraPrintLabel 
      Caption         =   "Label Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   825
      Left            =   90
      TabIndex        =   20
      Top             =   5040
      Width           =   9585
      Begin VB.ComboBox cboPrintServer 
         Height          =   315
         ItemData        =   "frmfrmFtTRLotComplete.frx":08FE
         Left            =   1680
         List            =   "frmfrmFtTRLotComplete.frx":0900
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   300
         Width           =   2445
      End
      Begin VB.ComboBox cboLabelSpec 
         Height          =   315
         Left            =   5595
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   300
         Width           =   2400
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Printer Server:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label lblLabelForm 
         AutoSize        =   -1  'True
         Caption         =   "Label Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4470
         TabIndex        =   23
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Label lblGenChildLot 
      AutoSize        =   -1  'True
      Caption         =   "ChildLotID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   26
      Top             =   6360
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblChildLotLabel 
      AutoSize        =   -1  'True
      Caption         =   "Child Lot ID : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   6360
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblChildLot 
      AutoSize        =   -1  'True
      Caption         =   "ChildLotID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   15
      Top             =   6000
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Equipment 
      AutoSize        =   -1  'True
      Caption         =   "Equipment Id :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   14
      Top             =   135
      Width           =   1245
   End
End
Attribute VB_Name = "frmFtTRLotComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msMODULE_ID As String = "frmFtTRLotComplete"
'**********
' Variable Declaration
'**********
Private moCwMbx     As Object
Private moFwCAT     As Object
Private moFwMDL     As Object
Private moFwOPR     As Object
Private moFwPRP     As Object
Private moFwWF      As Object
Private moFwWIP     As Object
Private moProRawSql As Object
Private moAppLog    As Object

Private msRuleName  As String
Private msLotId     As String
Private msEqpId     As String

Private Const msEqpType                 As String = "TR"
Private Const msEqpType2                As String = "FTPACKING"
Private Const msEqpType3                As String = "SCANNER"

Private Const miSpdFieldPos_Check       As Integer = 1
Private Const miSpdFieldPos_TECN        As Integer = 2
Private Const miSpdFieldPos_LotId       As Integer = 3
Private Const miSpdFieldPos_IPN         As Integer = 4
Private Const miSpdFieldPos_CustomerName As Integer = 5
Private Const miSpdFieldPos_LotQty      As Integer = 6
Private Const miSpdFieldPos_FailQty     As Integer = 7
Private Const miSpdFieldPos_Reason      As Integer = 8
Private Const miSpdFieldPos_ReasonCode  As Integer = 9
Private Const miSpdFieldPos_DefectQty   As Integer = 10

'Add by Sam start on 20101026 for ReqNo:JC201000238
Private Const miSpdFieldPos_SumHold     As Integer = 11

Private Const msSpdFieldPos_SumHold     As String = "SumHold"

'Add by Sam End on 20101026 for ReqNo:JC201000238

'Add by Sam start on 20101215 for ReqNo:JC201000277
Private Const miSpdFieldPos_CrackChoose     As Integer = 12
Private Const msSpdFieldPos_CrackChoose     As String = "CrackChoose"
Private Const miSpdFieldPos_CrackComment     As Integer = 13
Private Const msSpdFieldPos_CrackComment     As String = "CrackComment"
Private Const miSpdFieldPos_CrackHold     As Integer = 14
Private Const msSpdFieldPos_CrackHold     As String = "CrackHold"
'Add by Sam End on 20101215 for ReqNo:JC201000277

Private moDefectQty As Collection
Private moReasonCode As Collection

Private mblotSelected As Boolean
Private miResult    As Integer

'Add by Sam start on 20101026 for ReqNo:JC201000238
'Modified by Tony on 2014/08/11 for Req.JC201400246
'Private Const msHoldCode   As String = "MK330"
'Private Const msHoldReason As String = "SUMMARY ABNORMAL"
Private Const msHoldCode   As String = "MK320"
Private Const msHoldReason As String = "FVI Summary Abnormal"


Private msHoldComment As String
'Add by Sam End on 20101026 for ReqNo:JC201000238
Private msHoldCommentPE As String 'Added by Jack on 2025/06/04 for #212677_BE#202400097 MES FVI Lot complete時, 如有: FVI Test type Error , 加Hold MK330


Public Property Let LotID(sLotID As String)
    On Error Resume Next
    msLotId = sLotID
End Property
Public Property Let EqpId(sEqpid As String)
    On Error Resume Next
    msEqpId = sEqpid
End Property
Public Property Set FwCATControl(oFwCAT As Object)
    On Error Resume Next
    Set moFwCAT = oFwCAT
End Property
Public Property Set FwWFControl(oFwWF As Object)
    On Error Resume Next
    Set moFwWF = oFwWF
End Property
Public Property Set FwOPRControl(oFwOPR As Object)
    On Error Resume Next
    Set moFwOPR = oFwOPR
End Property
Public Property Set FwMDLControl(oFwMDL As Object)
    On Error Resume Next
    Set moFwMDL = oFwMDL
End Property
Public Property Set FwPRPControl(oFwPRP As Object)
    On Error Resume Next
    Set moFwPRP = oFwPRP
End Property
Public Property Set FwWIPControl(oFwWIP As Object)
    On Error Resume Next
    Set moFwWIP = oFwWIP
End Property
Public Property Set CwMbxControl(oCwMbx As Object)
    On Error Resume Next
    Set moCwMbx = oCwMbx
End Property
Public Property Set ProRawSqlControl(oProRawSqlControl As Object)
    On Error Resume Next
    Set moProRawSql = oProRawSqlControl
End Property
Public Property Set MainTraceLog(oLogCtrl As Object)
    On Error Resume Next
    Set moAppLog = oLogCtrl
End Property

Public Property Let RuleName(sRuleName As String)
    On Error Resume Next
    msRuleName = sRuleName
End Property
Private Sub ResetFwControls()
'**************************************************
'**************************************************
    On Error Resume Next
    Set moCwMbx = Nothing
    Set moFwCAT = Nothing
    Set moFwMDL = Nothing
    Set moFwOPR = Nothing
    Set moFwPRP = Nothing
    Set moFwWF = Nothing
    Set moFwWIP = Nothing
    Set moProRawSql = Nothing
    Set moAppLog = Nothing

End Sub
Public Property Get Result() As Integer
'**************************************************
'**************************************************
    On Error Resume Next
    Result = miResult

End Property



Private Sub cboEqpID_Change()
    DispLotData
End Sub

Private Sub cboEqpID_Click()
    Call cboEqpID_Change
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    miResult = vbCancel
    Me.Hide
End Sub

'Add by Sam on 20110308 for ReqNo:JC201100064
Private Sub cmdGenLotID_Click()
    
    Dim vLotID As Variant
    Dim iIdx As Integer
    
    With Me.spdLotList
        .Col = 1
        For iIdx = 1 To .MaxRows
            .Row = iIdx
            If .Text = "1" Then
                .GetText miSpdFieldPos_LotId, iIdx, vLotID
                Exit For
            End If
        Next
    End With
    
    If cboMergeLots.Text <> "" Then
        If cboMergeLots.Text = vLotID Then
            lblGenChildLot.Caption = lblChildLot.Caption
        Else
            lblGenChildLot.Caption = GetChildLotId(moAppLog, moFwWIP, moFwWF, moCwMbx, cboMergeLots.Text, gsSTAGE_FT)
        End If
        cboMergeLots.Enabled = False
        cmdGenLotID.Enabled = False
        Call Qty_Change
    End If
End Sub

Private Sub cmdLotCancel_Click()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

'----
' Init
'----
    sProcID = "cmdOk_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
'----
' Condition Checking
'----
    Dim iResult As VBA.VbMsgBoxResult
    Call UtShowMsgBox("Lot will be cancel" & vbNewLine & "此批將取消執行", vbYesNo, , , iResult)

    If iResult = vbCancel Then
        GoTo ExitHandler
    End If
'----
' Action
'----
    Dim oLot As FwLot
    Dim oEqp As FwEquipment
    Dim sOPName As String
    Dim sAttrName As String
    Dim sAttrValue As String
    Dim oAttrs As FwAttributes, oAttr As FwAttribute
    '------------------------------------------------------------------
    Dim sGroupHistory As String
    sGroupHistory = msRuleName & "-" & GetTxnSeq(moProRawSql, moAppLog)
    '------------------------------------------------------------------
    Screen.MousePointer = vbHourglass

    Set oLot = FwuRetrieveLot(moFwWIP, msLotId, moAppLog)
    If TimeStampChange(oLot.Id, oLot.TimeStamp, Me.spdLotList, miSpdFieldPos_LotId) = True Then
        GoTo ExitHandler
    End If
    Set oEqp = FwuRetrieveEqp(moFwMDL, Me.cboEqpID)
    sOPName = moFwOPR.ActiveUser.UserName
    
    '2009/05/06 update lot attributes in one transaction(only not null value)
    Set oAttrs = moFwWIP.CreateFwAttributes
    
    sAttrName = gsLOT_CUSTOMATTR_STATUS
    sAttrValue = gsLOTSTATUS_WAITD
    'Call FwuModifyLotCustomAttribute(oLot, sAttrName, sAttrValue, sOPName, sGroupHistory, moAppLog)
    oAttrs.Add sAttrName, sAttrValue, fwAttrString
    
    sAttrName = gsLOT_CUSTOMATTR_LAST_EQP_ID
    sAttrValue = oEqp.Id
    'Call FwuModifyLotCustomAttribute(oLot, sAttrName, sAttrValue, sOPName, sGroupHistory, moAppLog)
    oAttrs.Add sAttrName, sAttrValue, fwAttrString

    'sAttrName = gsLOT_CUSTOMATTR_CUR_EQP_ID
    'sAttrValue = ""
    'Call FwuModifyLotCustomAttribute(oLot, sAttrName, sAttrValue, sOPName, sGroupHistory, moAppLog)
    
    Set oAttr = oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
    oAttr.Value = ""
    'Modify by Sam on 20090717 for BBD ENABLE
'    oLot.ModifyAttribute oAttr, sOPName, sGroupHistory, , oAttrs, , , , False
    oLot.ModifyAttribute oAttr, sOPName, sGroupHistory, , oAttrs
    oLot.Refresh
    
    Call RepositionToLastRule(moFwPRP, moFwWF, oLot, _
            oLot.CurrentStep.Steps.Item(1).CurrentRule.Id, sOPName, sGroupHistory, _
            oLot.CurrentStep.Steps.Item(1).ResourceType, moAppLog)
    oLot.Refresh
    
    Dim sCurLotIds As String, sCurLotIDsArry() As String
    Dim sCurIPN As String, sCurIPNArry() As String
    Dim iIdx As Integer
    
    'CurLotID : 單批清空，多批扣移出之Lot
    sCurLotIds = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_CUR_LOT_ID)
    If InStr(1, sCurLotIds, oLot.Id) = 0 Then GoTo ExitHandler
    
    sCurIPN = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_CUR_IPN)
    If InStr(1, sCurLotIds, ";") > 0 Then
        sCurLotIDsArry = Split(sCurLotIds, ";")
        sCurIPNArry = Split(sCurIPN, ";")
        sCurLotIds = ""
        sCurIPN = ""
        For iIdx = LBound(sCurLotIDsArry) To UBound(sCurLotIDsArry)
            If sCurLotIDsArry(iIdx) <> oLot.Id Then
                sCurLotIds = sCurLotIds & sCurLotIDsArry(iIdx) & ";"
                sCurIPN = sCurIPN & sCurIPNArry(iIdx) & ";"
            End If
        Next iIdx
        If Right(sCurLotIds, 1) = ";" Then sCurLotIds = Left(sCurLotIds, Len(sCurLotIds) - 1)
        If Right(sCurIPN, 1) = ";" Then sCurIPN = Left(sCurIPN, Len(sCurIPN) - 1)
    Else
        If oEqp.CustomAttributes(gsEQP_CUSTOMATTR_CUR_LOT_ID) = oLot.Id Or _
        oEqp.State = gsEQPSTATUS_IDLE Then
            sCurLotIds = ""
            sCurIPN = ""
        End If
    End If
    
    Set oAttrs = moFwWIP.CreateFwAttributes
    If oEqp.CustomAttributes(gsEQP_CUSTOMATTR_CUR_LOT_ID) <> sCurLotIds Then
        oAttrs.Add gsEQP_CUSTOMATTR_CUR_LOT_ID, sCurLotIds, fwAttrString
        oAttrs.Add gsEQP_CUSTOMATTR_CUR_IPN, sCurIPN, fwAttrString
        If sCurLotIds = "" Then
            oAttrs.Add gsEQP_CUSTOMATTR_CUR_LOT_START_TIME, "", fwAttrString
        End If
    End If
    
    If oAttrs.Count > 0 Then
        oEqp.ModifyAttributes oAttrs, sOPName, sGroupHistory
        oEqp.Refresh
    End If
    
    If oEqp.Capacity < oEqp.CapacityMaximum Then
        oEqp.ChangeCapacity gsCAPACITY_RELEASE, 1, "", sOPName, sGroupHistory
        oEqp.Refresh
    End If
    
    If sCurLotIds = "" Then
        oEqp.ChangeState gsEQPSTATUS_IDLE, , sOPName, sGroupHistory, , ""
        oEqp.Refresh
    End If
    
    miResult = vbOK
    Me.Hide
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    Screen.MousePointer = vbDefault
    ' Cleaning up
    Set oLot = Nothing
    Set oEqp = Nothing
    
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
           Case glERR_INVALIDOBJECT
               ' Retry code goes here...
           Case Else
                 typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
        End Select
        '---- Start of Select Case Block ----
        Screen.MousePointer = vbDefault
        cmdOk.Enabled = True
        miResult = vbCancel
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If

End Sub

'================================================================================
' Sub: cmdOk_Click()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent Wei, CIT 2002/05/07
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
' [REV 02] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
'================================================================================
Private Sub cmdOK_Click()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

'----
' Init
'----
    sProcID = "cmdOk_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    Screen.MousePointer = vbHourglass
'----
' Condition Checking
'----
    If (moFwWIP Is Nothing) Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                        FormatErrorText(gsETX_INVALIDOBJECT, "FwWIP"))
    End If
    
    If Me.txtTRQty = "" Or Me.txtReelCount = "" Then
        UtShowMsgBox "Please input TRQty and Reel Count"
        GoTo ExitHandler
    End If
    If Me.txtRemain = "" Then Me.txtRemain = "0"
    
    If Me.txtRemain <> "0" And lblChildLot.Caption = "" Then
        UtShowMsgBox "子批批號為空" & vbNewLine & vbNewLine & "Child LotID is null"
        GoTo ExitHandler
    End If
'----
' Action
'----
    Dim oLot As FwLot, oChildLot As FwLot
    Dim oEqp As FwEquipment
    Dim sOPName As String
    Dim oFwNullStr As FwStrings
    Dim sAttrName As String, sAttrValue As String
    Dim oAttr As FwAttribute, oAttrs As FwAttributes
    Dim oPlan As FwProcessPlan, oAggrStep As FwAggregateStep, oStep As FwStep
    Dim sSQL As String, colSQLResult As Collection
    Dim sTxnTime As String
    Dim sDefectQty() As String, sReasonCode() As String
    Dim oReasonCode As FwStrings, oDefectQty As Collection
    Dim iIdx As Integer
    Dim iResult As VbMsgBoxResult
    Dim bIsPrintLabel As Boolean
    
'Add by Sam start on 20101027 for ReqNo:JC201000238
    Dim oNextRule    As Object
    Dim vSumHold As Variant
'Add by Sam End on 20101027 for ReqNo:JC201000238
    
'Add by Sam start on 20101027 for ReqNo:JC201000277
    Dim vCrackChoose As Variant
    Dim vCrackComment As Variant
    Dim vCrackHold As Variant
    Dim oComment      As FwComment
'Add by Sam End on 20101027 for ReqNo:JC201000277
    
'Add by Sam Start on 20110308 for ReqNo:JC201100064
    Dim oParentLot  As FwLot
    Dim sMergeTime As String
    Dim oChildLotIds As FwStrings

'Add by Sam End on 20110308 for ReqNo:JC201100064
    
    'Added by Jack on 2012/12/17 for 併批捲/箱 批號管制專案 Project <Start>
    Dim frmTmp          As Form
    Dim colResult       As Collection
    'Added by Jack on 2012/12/17 for 併批捲/箱 批號管制專案 Project <End>
    
    Dim sEqType2    As String 'Add by Tony on 2013/07/19 for Req.JC201300192
    
    Dim sMsg As String  'Add by Sam on 20140321 for ReqNo:JC201400084
    
    'add by Ernest on 2018/9/3 for BE#201600344/BE#201800415-----start
    Dim oGenChildLot            As FwLot
    Dim sUnpack                 As String
    Dim sRemain_unpack_TAT      As String
    'add by Ernest on 2018/9/3 for BE#201600344/BE#201800415-----end
    
    Dim sHoldCode As String  'add by Ernest on 2020/03/12 for 組織改組
    Dim sHoldCodePE As String  'Added by Jack on 2025/06/04 for #212677_BE#202400097 MES FVI Lot complete時, 如有: FVI Test type Error , 加Hold MK330
    
    Dim lDefectQty As Long 'Add by Sam on 20200904 for ReqNO:202000140
    
    '20060601, vencent M200605029, add label print function
    Call UtShowMsgBox("Do you want to Print Barcode Label?" & vbCrLf & _
                      "是否列印條碼?", vbYesNo, , , iResult)
    If iResult = vbOK Then
        bIsPrintLabel = True
        If Len(Me.cboPrintServer.Text) = 0 Then
            Call UtShowMsgBox("Please select Printer Server!")
            GoTo ExitHandler
        End If
        If Len(Me.cboLabelSpec.Text) = 0 Then
            Call UtShowMsgBox("Please select Label Type!")
            GoTo ExitHandler
        End If
    End If

   
    sTxnTime = GetSystemTime(moAppLog, moFwWIP, moFwWF, moCwMbx)
    
    
    Dim sSplitGroupHistory As String 'Add by Sam on 20110308 for ReqNo:JC201100064
    sSplitGroupHistory = "SplitLot" & "-" & GetTxnSeq(moProRawSql, moAppLog)
    
    '------------------------------------------------------------------
    Dim sGroupHistory As String
    sGroupHistory = msRuleName & "-" & GetTxnSeq(moProRawSql, moAppLog)
    '------------------------------------------------------------------
    
    Dim sSplitGroupHistory2 As String 'Add by Sam on 20110308 for ReqNo:JC201100064
    sSplitGroupHistory2 = "SplitLot" & "-" & GetTxnSeq(moProRawSql, moAppLog)
        
    Set oLot = FwuRetrieveLot(moFwWIP, msLotId, moAppLog)
    
    'Add by Jeff, 2002/08/18, for timestamp check
    If TimeStampChange(oLot.Id, oLot.TimeStamp, Me.spdLotList, miSpdFieldPos_LotId) = True Then
        GoTo ExitHandler
    End If
    'End of add
    
    'Add by Sam start on 20140321 for ReqNo:JC201400084'
    If CheckFtSelfCheck(oLot.Id, sMsg, moFwWIP, moFwMDL, moAppLog, moProRawSql) = False Then
        UtShowMsgBox "LotID:" & oLot.Id & vbNewLine & " TA自主檢查記錄不完整 / Unfinished TA Self-Check Record" & vbNewLine & _
                     "未完成項目 / Unfinished Item：" & vbNewLine & sMsg
        GoTo ExitHandler
    End If
    'Add by Sam end on 20140321 for ReqNo:JC201400084'
    
    Set oEqp = FwuRetrieveEqp(moFwMDL, Me.cboEqpID)
    sOPName = moFwOPR.ActiveUser.UserName
    
    'Add by Sam Start on 20200904 for ReqNO:202000140
    If CheckFviMergeLotID(Trim(oLot.Id), moProRawSql, moAppLog) Then
        'defect
        lDefectQty = 0
        If Me.txtFail <> "0" And IsNumeric(Me.txtFail) Then
            With Me.spdLotList
                .Row = 1
                .Col = miSpdFieldPos_Check
                Do While .Row <= .MaxRows
                    If .Text = "1" Then
                        .Col = miSpdFieldPos_ReasonCode
                        sReasonCode = Split(.Text, ";")
                        .Col = miSpdFieldPos_DefectQty
                        sDefectQty = Split(.Text, ";")
                        Exit Do
                    End If
                    .Row = .Row + 1
                Loop
                For iIdx = LBound(sReasonCode) To UBound(sReasonCode)
                    lDefectQty = lDefectQty + Val(sDefectQty(iIdx))
                Next
            End With
        End If
        
        'remain - split lot
        If Me.txtRemain <> "0" Then
            lDefectQty = lDefectQty + Val(txtRemain.Text)
        End If
        
        Set colResult = modCustomUpdate.GetCarrierMergeList(oLot.Id, moFwWIP, moFwWF, moCwMbx, oLot, moProRawSql, moAppLog)
        If Not colResult Is Nothing Then
            If colResult.Count > 0 Then
                Set frmTmp = New frmMergeCarrierChipQtyModify
                Load frmTmp
                With frmTmp
                    Set .CwMbxControl = moCwMbx
                    Set .FwMDLControl = moFwMDL
                    Set .FwWIPControl = moFwWIP
                    Set .FwOPRControl = moFwOPR
                    Set .FwPRPControl = moFwPRP
                    Set .FwWFControl = moFwWF
                    Set .FwCATControl = moFwCAT
                    Set .MainTraceLog = moAppLog
                    Set .CwMbxControl = moCwMbx
                    Set .ProRawSqlControl = moProRawSql
                    oLot.Refresh
                    .RuleName = msRuleName
                    .LotID = oLot.Id
                    .Qty = oLot.ComponentQuantity - lDefectQty
                    .CallByWhichForm = "frmFtTRLotComplete"
                                        
                    .Init
                    .Show vbModal
                    If .Result <> vbOK Then
                        GoTo ExitHandler
                    End If
                    Screen.MousePointer = vbDefault
                End With
                Unload frmTmp
            Else
                Call UtShowMsgBox("Lot '" & oLot.Id & " 無併批明細!! [併批捲/箱 批號管制專案]" & vbNewLine & _
                                  "Lot '" & oLot.Id & " no Merge List!!")
            End If
        End If
    End If
    'Add by Sam End on 20200904 for ReqNO:202000140
    
    'Add by Sam start on 20101028 for ReqNo:JC201000238
    With Me.spdLotList
        .Row = 1
        .Col = miSpdFieldPos_Check
        Do While .Row <= .MaxRows
            If .Text = "1" Then
                .Col = miSpdFieldPos_SumHold
                 vSumHold = .Text
                 
                 'Add by Sam start on 20101215 for ReqNo:JC201000277
                 .Col = miSpdFieldPos_CrackChoose
                 vCrackChoose = .Text
                 .Col = miSpdFieldPos_CrackComment
                 vCrackComment = .Text
                 .Col = miSpdFieldPos_CrackHold
                 vCrackHold = .Text
                 'Add by Sam end on 20101215 for ReqNo:JC201000277
            End If
            .Row = .Row + 1
        Loop
    End With
    'Add by Sam End on 20101028 for ReqNo:JC201000238
    
    
    'Add by Sam start on 20101215 for ReqNo:JC201000277
    If Trim(vCrackComment) <> "" Then
        oLot.Refresh
        Set oComment = moFwWIP.CreateFwComment
        oComment.Initialize msRuleName, msRuleName, "Choose " & vCrackChoose & ": " & vCrackComment
        oLot.AddComment oComment, moFwOPR.ActiveUser.UserName, sGroupHistory
        oLot.Refresh
    End If
    'Add by Sam End on 20101215 for ReqNo:JC201000277
    
    
    'defect
    If Me.txtFail <> "0" And IsNumeric(Me.txtFail) Then
    
        'Add by Tony Start on 2013/07/19 for Req.JC201300192
        sSQL = "select eqtype2 from view_b2b_fweqarea t WHERE EQID ='" & oEqp.Id & "' "
        Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
        If colSQLResult.Count > 0 Then
            sEqType2 = colSQLResult.Item(1).Item(1)
        End If
        'Add by Tony End on 2013/07/19 for Req.JC201300192
    
        With Me.spdLotList
            .Row = 1
            .Col = miSpdFieldPos_Check
            Do While .Row <= .MaxRows
                If .Text = "1" Then
                    .Col = miSpdFieldPos_ReasonCode
                    sReasonCode = Split(.Text, ";")
                    .Col = miSpdFieldPos_DefectQty
                    sDefectQty = Split(.Text, ";")
                    Exit Do
                End If
                
                .Row = .Row + 1
            Loop
            Set oFwNullStr = moFwWIP.CreateFwStrings
            For iIdx = LBound(sReasonCode) To UBound(sReasonCode)
                Set oReasonCode = moFwWIP.CreateFwStrings
                oReasonCode.Add sReasonCode(iIdx)
                'Modify by Sam on 20090717 for BBD ENABLE
'                oLot.Scrap oReasonCode, sDefectQty(iIdx), oFwNullStr, _
'                    oLot.CurrentStep.Steps.Item(1).Id, sOPName, sGroupHistory, , , , , , , False, sTxnTime
                oLot.Scrap oReasonCode, sDefectQty(iIdx), oFwNullStr, _
                    oLot.CurrentStep.Steps.Item(1).Id, sOPName, sGroupHistory, , , , , , , , sTxnTime
                    
                oLot.Refresh
                
                'Add by Tony on 2013/07/17 for Req.JC201300192
                Call Insert_Tbl_FTInhouse_Loss_Rec(moAppLog, moProRawSql, oLot, CLng(sDefectQty(iIdx)), _
                                           oLot.CurrentStep.Steps.Item(1).Id _
                                           , sReasonCode(iIdx), sOPName, _
                                           oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID), sEqType2)
            Next
        End With
    End If
    
    'remain - split lot
    If Me.txtRemain <> "0" Then
        Set oFwNullStr = moFwWIP.CreateFwStrings
        oLot.Refresh
        
        'Add by Sam start on20120523 for ReqNo:JC201200149
        Call modCustom.SplitLot(moFwWIP, moAppLog, moFwWF, moCwMbx, moProRawSql, _
                                oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), oLot.Id, lblChildLot.Caption, Val(txtRemain.Text), _
                                sOPName, sSplitGroupHistory)
        'Add by Sam end on20120523 for ReqNo:JC201200149
        'Mark by Sam start on 20120523 for ReqNo:JC201200149
'
'        'Modify by Sam on 20110308 ,將Split的Grouphistkey分開
'       'Modify by Sam on 20090717 for BBD ENABLE
''        oLot.Split lblChildLot, CLng(txtRemain), oFwNullStr, sOPName, , , , sGroupHistory, , , , , , , False, sTxnTime
'        'oLot.Split lblChildLot, CLng(txtRemain), oFwNullStr, sOPName, , , , sGroupHistory, , , , , , , , sTxnTime
'        oLot.Split lblChildLot, CLng(txtRemain), oFwNullStr, sOPName, , , , sSplitGroupHistory, , , , , , , , sTxnTime
'        oLot.Refresh
'
'        sAttrName = gsLOT_CUSTOMATTR_CHIP_QTY
'        sAttrValue = CStr(oLot.ComponentQuantity)
'       'Modify by Sam on 20090717 for BBD ENABLE
''        Call FwuModifyLotCustomAttribute(oLot, sAttrName, sAttrValue, sOPName, sGroupHistory, moAppLog, sTxnTime, False)
''        Call FwuModifyLotCustomAttribute(oLot, sAttrName, sAttrValue, sOPName, sGroupHistory, moAppLog, sTxnTime, True)
'        Call FwuModifyLotCustomAttribute(oLot, sAttrName, sAttrValue, sOPName, sSplitGroupHistory, moAppLog, sTxnTime, True)
        'Mark by Sam End on 20120523 for ReqNo:JC201200149

        Set oChildLot = moFwWIP.LotById(lblChildLot)
        
        'Mark by Sam start on 20120523 for ReqNo:JC201200149
'        'Add by Sam start on 20110308 ,Split分開處理,單獨處理數量變更
'        sAttrName = gsLOT_CUSTOMATTR_CHIP_QTY
'        sAttrValue = CStr(oChildLot.ComponentQuantity)
'        Call FwuModifyLotCustomAttribute(oChildLot, sAttrName, sAttrValue, sOPName, sSplitGroupHistory, moAppLog, sTxnTime)
'        'Add by Sam End on 20110308
        'Mark by Sam end on 20120523 for ReqNo:JC201200149
        
        '2009/05/06 update lot attributes in one transaction(only not null value)
        Set oAttrs = moFwWIP.CreateFwAttributes
        
        sAttrName = gsLOT_CUSTOMATTR_STATUS
        sAttrValue = gsLOTSTATUS_WAITD
        'Call FwuModifyLotCustomAttribute(oChildLot, sAttrName, sAttrValue, sOPName, sGroupHistory, moAppLog, sTxnTime)
        oAttrs.Add sAttrName, sAttrValue, fwAttrString
        
        'Mark by SAm start on 20110308 for Split分開處理
'        sAttrName = gsLOT_CUSTOMATTR_CHIP_QTY
'        sAttrValue = CStr(oChildLot.ComponentQuantity)
'        'Call FwuModifyLotCustomAttribute(oChildLot, sAttrName, sAttrValue, sOPName, sGroupHistory, moAppLog, sTxnTime)
'        oAttrs.Add sAttrName, sAttrValue, fwAttrInteger
        'Mark by Sam end on 20110308
        
        sAttrName = gsLOT_CUSTOMATTR_LAST_EQP_ID
        sAttrValue = oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
        If sAttrValue = "" Then
        'Modify by Sam on 20090717 for BBD ENABLE
'            Call FwuModifyLotCustomAttribute(oChildLot, sAttrName, sAttrValue, sOPName, sGroupHistory, moAppLog, sTxnTime, False)
            Call FwuModifyLotCustomAttribute(oChildLot, sAttrName, sAttrValue, sOPName, sGroupHistory, moAppLog, sTxnTime, True)
        Else
            oAttrs.Add sAttrName, sAttrValue, fwAttrString
        End If
        
        'Modify by Sam on 20090717 for BBD ENABLE
'        Call FwuModifyLotCustomAttribute(oChildLot, gsLOT_CUSTOMATTR_CUR_EQP_ID, "", sOPName, sGroupHistory, moAppLog, sTxnTime, False)
'        Call FwuModifyLotCustomAttribute(oChildLot, gsLOT_CUSTOMATTR_PGNAME, "", sOPName, sGroupHistory, moAppLog, sTxnTime, False)
'        Call FwuModifyLotCustomAttribute(oChildLot, gsLOT_CUSTOMATTR_PROCESS_HOUR, "", sOPName, sGroupHistory, moAppLog, sTxnTime, False)
        Call FwuModifyLotCustomAttribute(oChildLot, gsLOT_CUSTOMATTR_CUR_EQP_ID, "", sOPName, sGroupHistory, moAppLog, sTxnTime, True)
        Call FwuModifyLotCustomAttribute(oChildLot, gsLOT_CUSTOMATTR_PGNAME, "", sOPName, sGroupHistory, moAppLog, sTxnTime, True)
        Call FwuModifyLotCustomAttribute(oChildLot, gsLOT_CUSTOMATTR_PROCESS_HOUR, "", sOPName, sGroupHistory, moAppLog, sTxnTime, True)
        
        'Call FwuModifyLotCustomAttribute(oChildLot, gsLOT_CUSTOMATTR_PROCESSTIMEUNIT, " ", sOPName, sGroupHistory, moAppLog, sTxnTime)
        'Call FwuModifyLotCustomAttribute(oChildLot, gsLOT_CUSTOMATTR_TEMPERATURE, "", sOPName, sGroupHistory, moAppLog, sTxnTime)
        
        oAttrs.Add gsLOT_CUSTOMATTR_PROCESSTIMEUNIT, " ", fwAttrString
        Set oAttr = oChildLot.CustomAttributes(gsLOT_CUSTOMATTR_TEMPERATURE)
        oAttr.Value = ""
        'Modify by Sam on 20090717 for BBD ENABLE
'        oChildLot.ModifyAttribute oAttr, sOPName, sGroupHistory, , oAttrs, , , , False, sTxnTime
        oChildLot.ModifyAttribute oAttr, sOPName, sGroupHistory, , oAttrs, , , , , sTxnTime
        oChildLot.Refresh

        'RepositionRule to Last rule
        Call RepositionToLastRule(moFwPRP, moFwWF, oChildLot, _
                oChildLot.CurrentStep.Steps.Item(1).CurrentRule.Id, sOPName, sGroupHistory, _
                oChildLot.CurrentStep.Steps.Item(1).ResourceType, moAppLog)
        oChildLot.Refresh
    
        
'        sSql = "Insert into " & gsCAT_TBL_LOT_INFO _
'             & " ( " & gsCAT_TLI_LOT_ID & ", " & gsCAT_TLI_WAFER_IPN & ", " _
'                     & gsCAT_TLI_FG_IPN & ", " _
'                     & gsCAT_TLI_LOT_ORDER_QTY & ", " _
'                     & gsCAT_TLI_SPLIT_FLAG & ", " _
'                     & gsCAT_TLI_CUR_CARRIER_TYPE & ", " & gsCAT_TLI_SOURCE_LOT_ID & ", " _
'                     & gsCAT_TLI_VEND_LOT & ", " & gsCAT_TLI_MCP_FLAG & ", " _
'                     & gsCAT_TLI_ORDER_TYPE & ", " & gsCAT_TLI_CHECK_MARK & ", " _
'                     & gsCAT_TLI_STAGE_DUE & ", " & gsCAT_TLI_VENDOR & ", " _
'                     & gsCAT_TLI_CREATE_USER_ID & ", " & gsCAT_TLI_CREATE_TIME & ", " _
'                     & gsCAT_TLI_SPLIT_SEQ & ", " & gsCAT_TLI_BOXQTY & ", " _
'                     & gsCAT_TLI_ORDER_ID & ", " & gsCAT_TLI_INTERFACEFLAG
'        sSql = sSql & " ) select " _
'               & "'" & oChildLot.Id & "', " & gsCAT_TLI_WAFER_IPN & ", " _
'                     & gsCAT_TLI_FG_IPN & ", " _
'                     & gsCAT_TLI_LOT_ORDER_QTY & ", " _
'                     & gsCAT_TLI_SPLIT_FLAG & ", " _
'                     & gsCAT_TLI_CUR_CARRIER_TYPE & ", '" & olot.Id & "', " _
'                     & gsCAT_TLI_VEND_LOT & ", " & gsCAT_TLI_MCP_FLAG & ", " _
'                     & gsCAT_TLI_ORDER_TYPE & ", " & gsCAT_TLI_CHECK_MARK & ", " _
'                     & gsCAT_TLI_STAGE_DUE & ", " & gsCAT_TLI_VENDOR & ", " _
'                     & "'" & sOPName & "', to_char(sysdate,'YYYYMMDD HH24MISS')||'000', " _
'                     & "'A1'" & ", " & gsCAT_TLI_BOXQTY & ", " _
'                     & gsCAT_TLI_ORDER_ID & ", " & gsCAT_TLI_INTERFACEFLAG _
'                & " from " & gsCAT_TBL_LOT_INFO _
'                & " where " & gsCAT_TLI_LOT_ID & " = '" & olot.Id & "'"
'        Set colSQLResult = moProRawSql.QueryDatabase(sSql)

        'Mark by Sam start on 20120523 for ReqNo:JC201200149
'        '*********
'        '* Insert Child Lot InOutTime into tbl_lot_inout_time
'        '* add by Nelson on 2006/11/03 for ReqNo:M200611004 , OBS
'        '*********
'        Call modCustom.Clone_LotInOutTime(oLot.Id, oChildLot.Id, moFwOPR.ActiveUser.UserName, moProRawSql, moAppLog)
'
'        '*********
'        '* Insert Child Lot TestRec into tbl_lot_test_rec
'        '* add by Nelson on 2008/05/19 for Testing Flow Gating Project.
'        '*********
'        Call modCustom.Clone_LotTestRec(oLot.Id, oChildLot.Id, moFwOPR.ActiveUser.UserName, moProRawSql, moAppLog)
'
'        Call modCustom.Clone_LotInfo(oLot.Id, oChildLot.Id, moFwOPR.ActiveUser.UserName, moProRawSql, moAppLog)
        'Mark by Sam end on 20120523 for ReqNo:JC201200149
      
        'Add by Sam start on 20110308 for ReqNo:JC201100064
        If lblGenChildLot.Caption <> lblChildLot.Caption Then
            sMergeTime = GetSystemTime(moAppLog, moFwWIP, moFwWF, moCwMbx)
                        
            '將User指定的Lot Unterminate
            Set oParentLot = moFwWIP.LotById(cboMergeLots.Text)
            
            Call FwuModifyLotCustomAttribute(oParentLot, gsLOT_CUSTOMATTR_STATUS, _
                     gsLOTSTATUS_WAITD, _
                     moFwOPR.ActiveUser.UserName, sGroupHistory, moAppLog, sMergeTime)
                     
            oParentLot.Refresh
            oParentLot.Unterminate moFwOPR.ActiveUser.UserName, sGroupHistory
            oParentLot.Refresh
            
            '將原子批併給User指定的Lot
            Set oChildLotIds = moFwWIP.CreateFwStrings
            oChildLotIds.Add oChildLot.Id
                                
            oParentLot.Merge oChildLotIds, moFwOPR.ActiveUser.UserName, sGroupHistory
            oParentLot.Refresh
            oChildLot.Refresh
            
            Call Merge_Clone(moAppLog, moProRawSql, oParentLot.Id, oChildLotIds, msRuleName, moFwOPR.ActiveUser.UserName) 'Add by Tony on 2013/05/02 for Lot健康管理PhaseII project.
             
            oParentLot.Refresh
            Call FwuModifyLotCustomAttribute(oParentLot, gsLOT_CUSTOMATTR_CHIP_QTY, oParentLot.ComponentQuantity, sOPName, sGroupHistory, moAppLog)
            oParentLot.Refresh
            
           
            Call FwuModifyLotCustomAttribute(oChildLot, gsLOT_CUSTOMATTR_CHIP_QTY, 0, sOPName, sGroupHistory, moAppLog)
            Call FwuModifyLotCustomAttribute(oChildLot, gsLOT_CUSTOMATTR_WAFER_QTY, 0, sOPName, sGroupHistory, moAppLog)
            Call FwuModifyLotCustomAttribute(oChildLot, gsLOT_CUSTOMATTR_STATUS, gsLOTSTATUS_TERMINATE, sOPName, sGroupHistory, moAppLog)
            oChildLot.Refresh
                        
            '將指定Lot全部分批給新的子批
            
            sMergeTime = GetSystemTime(moAppLog, moFwWIP, moFwWF, moCwMbx)
            
            'Add by Sam start on 20120523 for ReqNo:JC201200149
            Call modCustom.SplitLot(moFwWIP, moAppLog, moFwWF, moCwMbx, moProRawSql, _
                                    oParentLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), oParentLot.Id, lblGenChildLot.Caption, oParentLot.ComponentQuantity, _
                                    sOPName, sSplitGroupHistory2, , , , sMergeTime)
            
            'add by Ernest on 2018/9/3 for BE#201600344/BE#201800415-----start
            sSQL = "select " & gsCAT_TLI_UNPACKTIME & "," & gsCAT_TLI_REMAIN_UNPACK_TAT & " from " & gsCAT_TBL_LOT_INFO & _
                    " where " & gsCAT_TLI_LOT_ID & " ='" & oLot.Id & "'"
            Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
            If colSQLResult.Count > 0 Then
                sUnpack = colSQLResult.Item(1).Item(gsCAT_TLI_UNPACKTIME)
                sRemain_unpack_TAT = colSQLResult.Item(1).Item(gsCAT_TLI_REMAIN_UNPACK_TAT)
                
                sSQL = " UPDATE " & gsCAT_TBL_LOT_INFO & " SET " & _
                   gsCAT_TLI_REMAIN_UNPACK_TAT & "='" & sRemain_unpack_TAT & "' ," & _
                   gsCAT_TLI_UNPACKTIME & "='" & sUnpack & "' ," & _
                   gsCAT_TLI_GROUPHISTKEY & "='" & sGroupHistory & "' ," & _
                   gsCAT_TLI_UPDATE_USER_ID & "='" & moFwOPR.ActiveUser.UserName & "' ," & _
                   gsCAT_TLI_UPDATE_TIME & "='" & sMergeTime & "' " & _
                   " WHERE " & gsCAT_TLI_LOT_ID & " ='" & lblGenChildLot.Caption & "' "
                   
                Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
                
                Set oGenChildLot = moFwWIP.LotById(lblGenChildLot.Caption)
                
                oGenChildLot.Refresh
                                
                Set oComment = moFwWIP.CreateFwComment
                oComment.Initialize msRuleName, _
                                    msRuleName, _
                                    "Remain_Unpack_Tat = " & sRemain_unpack_TAT & " hour." & " unpacktime = " & sUnpack
                oGenChildLot.AddComment oComment, sOPName, sGroupHistory
                oGenChildLot.Refresh
                
            End If
            
            'add by Ernest on 2018/9/3 for BE#201600344/BE#201800415-----end
            'Add by Sam end on 20120523 for ReqNo:JC201200149
            
            'Mark by Sam start on 20120523 for ReqNo:JC201200149
'            oParentLot.Split lblGenChildLot.Caption, oParentLot.ComponentQuantity, oFwNullStr, sOPName, , , , sSplitGroupHistory2, , , , , , , , sMergeTime
'            oParentLot.Refresh
'
'            sAttrName = gsLOT_CUSTOMATTR_CHIP_QTY
'            sAttrValue = 0
'            Call FwuModifyLotCustomAttribute(oParentLot, sAttrName, sAttrValue, sOPName, sSplitGroupHistory2, moAppLog, sMergeTime, True)
'            Call FwuModifyLotCustomAttribute(oParentLot, gsLOT_CUSTOMATTR_STATUS, _
'                     gsLOTSTATUS_TERMINATE, _
'                     moFwOPR.ActiveUser.UserName, sSplitGroupHistory2, moAppLog, sMergeTime)
'
'            Set oChildLot = moFwWIP.LotById(lblGenChildLot.Caption)
'            sAttrName = gsLOT_CUSTOMATTR_CHIP_QTY
'            sAttrValue = CStr(oChildLot.ComponentQuantity)
'            Call FwuModifyLotCustomAttribute(oChildLot, sAttrName, sAttrValue, sOPName, sSplitGroupHistory2, moAppLog, sMergeTime, True)
'
'            '*********
'            '* Insert Child Lot InOutTime into tbl_lot_inout_time
'            '*********
'            Call modCustom.Clone_LotInOutTime(oParentLot.Id, oChildLot.Id, moFwOPR.ActiveUser.UserName, moProRawSql, moAppLog)
'            '*********
'            '* Insert Child Lot TestRec into tbl_lot_test_rec
'            '*********
'            Call modCustom.Clone_LotTestRec(oParentLot.Id, oChildLot.Id, moFwOPR.ActiveUser.UserName, moProRawSql, moAppLog)
'            Call modCustom.Clone_LotInfo(oParentLot.Id, oChildLot.Id, moFwOPR.ActiveUser.UserName, moProRawSql, moAppLog)
            'Mark by Sam end on 20120523 for ReqNo:JC201200149
        End If
        'Add by Sam end on 20110308 for ReqNo:JC201100064
        
        'Add by Sam start on 20110317 FOR ReqNo:JC201100064
        If Trim(lblGenChildLot.Caption) <> "" Then
            oLot.Refresh
            Set oComment = moFwWIP.CreateFwComment
            oComment.Initialize msRuleName, msRuleName, "TR Remain LotID: " & lblGenChildLot.Caption & " "
            oLot.AddComment oComment, moFwOPR.ActiveUser.UserName, sGroupHistory
            oLot.Refresh
        End If
        'Add by Sam End on 20110317 FOR ReqNo:JC201100064

        
        Call modPrint.PrintBarcodeLabel(oChildLot.Id, _
                                        oChildLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_WAFER_QTY).Value, _
                                        oChildLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_CHIP_QTY).Value, _
                                        oChildLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value, _
                                        oChildLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_LOT_OWNER).Value, _
                                        oChildLot.PlanId, _
                                        Me.cboPrintServer.Text, _
                                        Me.cboLabelSpec.Text, _
                                        moFwOPR.ActiveUser.UserName, moAppLog, moFwWIP, moFwWF, _
                                        moCwMbx, moProRawSql)
        
    Else
    '不需分批，但做過Scrap，改ChipQty
        If Val(oLot.ComponentQuantity) <> Val(oLot.CustomAttributes(gsLOT_CUSTOMATTR_CHIP_QTY)) Then
            sAttrName = gsLOT_CUSTOMATTR_CHIP_QTY
            sAttrValue = CStr(oLot.ComponentQuantity)
        'Modify by Sam on 20090717 for BBD ENABLE
'            Call FwuModifyLotCustomAttribute(oLot, sAttrName, sAttrValue, sOPName, sGroupHistory, moAppLog, sTxnTime, False)
            Call FwuModifyLotCustomAttribute(oLot, sAttrName, sAttrValue, sOPName, sGroupHistory, moAppLog, sTxnTime, True)
        End If
    End If

    Call FTTrackOut(oLot, oEqp, sOPName, sGroupHistory, sTxnTime, moFwWIP, moFwPRP, moProRawSql, moAppLog, True)
    
    Call UpdateFVIAccHr(oEqp.Id, sOPName, moProRawSql, moAppLog) 'Add by Tony on 2013/04/02 for FVI SETUP RECIPE自動化專案
    
'    If olot.CurrentStep.Steps.Item(1).Description Like "EXCHG*" Then
'
'    End If
    
    '2009/06/11 Vencent, for Testing Flow Gating II Project
    modTestFlowGating.ModifyLotFlowRec oLot.Id, "FVI", sTxnTime, sOPName, moProRawSql, moAppLog
                
    'add by Ernest on 2020/03/12 for 組織改組---start
    sSQL = "select " & gsCAT_TRCO_REASON_CODE & " from " & gsCAT_TBL_REASON_CODE & " where " & gsCAT_TRCO_CATEGORY & " ='Department' and " & gsCAT_TRCO_GROUP1 & "='HW' "
    
    Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
    
    If colSQLResult.Count > 0 Then
        sHoldCode = colSQLResult.Item(1).Item(1)
    End If
    'add by Ernest on 2020/03/12 for 組織改組---end
                
                
    'Modify by Sam start on 20101215 for ReqNo:JC201000277,增加Crack Hold
    'Add by Sam start on 20101028 for ReqNo:JC201000238
'    If vSumHold = "Y"  Then
    If vSumHold = "Y" Or vCrackHold = "Y" Then
        Set oNextRule = FwuRetrieveNextRule(moFwPRP, _
                                        oLot.CurrentStep.Steps.Item(1).CurrentRule.Id, _
                                        oLot.CurrentStep.Steps.Item(1), _
                                        moAppLog)
                                        
        oLot.RepositionRule oLot.CurrentStep.Steps.Item(1), oNextRule.Id, sOPName, sGroupHistory
        oLot.Refresh
    
        If vSumHold = "Y" Then
            'Modify by Sam on 20131031 for Hold Lot模組化
'            FtHoldLot oLot, msHoldCode, msHoldReason, sOPName, _
'               moProRawSql, sGroupHistory, moAppLog, , gsHOLD_TYPE_LOT_HOLD, msHoldComment
            'modify by Ernest on 2020/03/12 for 組織改組
            'Call modCustom.HoldLot(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot, msHoldCode, msHoldReason, _
            '                            gsHOLD_TYPE_LOT_HOLD, sOPName, sGroupHistory, msHoldComment)
            Call modCustom.HoldLot(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot, sHoldCode, msHoldReason, _
                                        gsHOLD_TYPE_LOT_HOLD, sOPName, sGroupHistory, msHoldComment)
            
            'Added by Jack on 2025/06/04 for #212677_BE#202400097 MES FVI Lot complete時, 如有: FVI Test type Error , 加Hold MK330 <Start>
            If msHoldCommentPE = "FVI Test type Error" Then
                sSQL = "select " & gsCAT_TRCO_REASON_CODE & " from " & gsCAT_TBL_REASON_CODE & " where " & gsCAT_TRCO_CATEGORY & " ='Department' and " & gsCAT_TRCO_GROUP1 & "='PE' "
                
                Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
                
                If colSQLResult.Count > 0 Then
                    sHoldCodePE = colSQLResult.Item(1).Item(1)
                End If
                
                Call modCustom.HoldLot(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot, sHoldCodePE, msHoldCommentPE, _
                                            gsHOLD_TYPE_LOT_HOLD, sOPName, sGroupHistory, msHoldCommentPE)
            End If
            'Added by Jack on 2025/06/04 for #212677_BE#202400097 MES FVI Lot complete時, 如有: FVI Test type Error , 加Hold MK330 <End>
        End If
        
        If vCrackHold = "Y" Then
             'Modify by Sam on 20131031 for Hold Lot模組化
'           FtHoldLot oLot, msHoldCode, "Hold Crack", sOPName, _
'               moProRawSql, sGroupHistory, moAppLog, , gsHOLD_TYPE_LOT_HOLD, CStr(vCrackComment)
            'modify by Ernest on 2020/03/12 for 組織改組
            'Call modCustom.HoldLot(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot, msHoldCode, "Hold Crack", _
            '                            gsHOLD_TYPE_LOT_HOLD, sOPName, sGroupHistory, CStr(vCrackComment))
            Call modCustom.HoldLot(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot, sHoldCode, "Hold Crack", _
                                        gsHOLD_TYPE_LOT_HOLD, sOPName, sGroupHistory, CStr(vCrackComment))
        End If
        
   'Add by Sam end on 20101028 for ReqNo:JC201000238
   'Modify by Sam End on 20101215 for ReqNo:JC201000277
    Else
    
        Call RepositionToNextRule(moFwPRP, moFwWF, oLot, _
                                  oLot.CurrentStep.Steps.Item(1).CurrentRule.Id, _
                                  sOPName, sGroupHistory, Me.cboEqpID, moAppLog)
    End If

    Call modPrint.PrintBarcodeLabel(oLot.Id, _
                                    oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_WAFER_QTY).Value, _
                                    oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_CHIP_QTY).Value, _
                                    oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value, _
                                    oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_LOT_OWNER).Value, _
                                    oLot.PlanId, _
                                    Me.cboPrintServer.Text, _
                                    Me.cboLabelSpec.Text, _
                                    moFwOPR.ActiveUser.UserName, moAppLog, moFwWIP, moFwWF, _
                                    moCwMbx, moProRawSql)

    'Added by Jack on 2012/12/17 for 併批捲/箱 批號管制專案 Project <Start>
    oLot.Refresh
    
    'Mark by Sam start on 20200904 for ReqNo:202000140
'    'Modify by Weilun on 20171013 for Remark 限制系統建立
'    '原先寫死第9碼為M, 改由利用模組到Table中搜尋
'    'If Mid(CStr(Trim(olot.Id)), 9, 1) = "M" And Len(Trim(olot.Id)) = 10 Then
'    If CheckFviMergeLotID(Trim(oLot.Id), moProRawSql, moAppLog) Then
'
'        Set colResult = modCustomUpdate.GetCarrierMergeList(oLot.Id, moFwWIP, moFwWF, moCwMbx, oLot, moProRawSql, moAppLog)
'        If Not colResult Is Nothing Then
'            If colResult.Count > 0 Then
'                Set frmTmp = New frmMergeCarrierChipQtyModify
'                Load frmTmp
'                With frmTmp
'                    Set .CwMbxControl = moCwMbx
'                    Set .FwMDLControl = moFwMDL
'                    Set .FwWIPControl = moFwWIP
'                    Set .FwOPRControl = moFwOPR
'                    Set .FwPRPControl = moFwPRP
'                    Set .FwWFControl = moFwWF
'                    Set .FwCATControl = moFwCAT
'                    Set .MainTraceLog = moAppLog
'                    Set .CwMbxControl = moCwMbx
'                    Set .ProRawSqlControl = moProRawSql
'                    oLot.Refresh
'                    .RuleName = msRuleName
'                    .LotID = oLot.Id
'                    .CallByWhichForm = "frmFtTRLotComplete"
'
'            '        If moEQP.Type = "FTPACKING" Then
'            '            .FTPacking = True
'            '        Else
'            '            .FTPacking = False
'            '        End If
'
'                    .Init
'                    .Show vbModal
'                    If .Result = vbOK Then
'                    End If
'                    Screen.MousePointer = vbDefault
'                End With
'                Unload frmTmp
'            Else
'                Call UtShowMsgBox("Lot '" & oLot.Id & " 無併批明細!! [併批捲/箱 批號管制專案]" & vbNewLine & _
'                                  "Lot '" & oLot.Id & " no Merge List!!")
'            End If
'        End If
'    End If
'    'Added by Jack on 2012/12/17 for 併批捲/箱 批號管制專案 Project <Start>
    'Mark by Sam END on 20200904 for ReqNo:202000140
    miResult = vbOK
    Me.Hide
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    Screen.MousePointer = vbDefault
    ' Cleaning up
    Set oLot = Nothing
    Set oEqp = Nothing
    Set oChildLot = Nothing
    
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
           Case glERR_INVALIDOBJECT
               ' Retry code goes here...
           Case Else
                 typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
        End Select
        '---- Start of Select Case Block ----
        Screen.MousePointer = vbDefault
        cmdOk.Enabled = True
        miResult = vbCancel
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Sub

Private Sub cmdSelfCheck_Click()
'Add by Sam on 20140321 for ReqNo:JC201400084
Dim ofrmMain As Object
Dim vChk As Variant
Dim vLotID As Variant
Dim lIdx As Long
Dim iFlag As Integer

With Me.spdLotList
    For lIdx = 1 To .MaxRows
        .GetText miSpdFieldPos_Check, lIdx, vChk
        If vChk = "1" Then
            .GetText miSpdFieldPos_LotId, lIdx, vLotID
            iFlag = iFlag + 1
            If iFlag > 1 Then
                vLotID = ""
                Exit For
            End If
        End If
    Next
End With

Set ofrmMain = New frmFTTaSelfCheckOperation
Load ofrmMain
With ofrmMain
   Set .CwMbxControl = moCwMbx
   Set .FwMDLControl = moFwMDL
   Set .FwWIPControl = moFwWIP
   Set .FwOPRControl = moFwOPR
   Set .FwPRPControl = moFwPRP
   Set .FwWFControl = moFwWF
   Set .FwCATControl = moFwCAT
   Set .MainTraceLog = moAppLog
   Set .CwMbxControl = moCwMbx
   Set .ProRawSqlControl = moProRawSql
       .RuleName = "FTTaSelfCheckOperation"
       .LotID = CStr(vLotID)
       .Init
       .Show vbModal
End With

End Sub

Private Sub cmdSendVersa3_Click()
'================================================================================
' Function: cmdSendVersa3_Click()
'--------------------------------------------------------------------------------
' Description:  For Req.JC201300055
'--------------------------------------------------------------------------------
' Author:       Tony Chang , 2013/03/08
'--------------------------------------------------------------------------------
'================================================================================
On Error GoTo ExitHandler:
Dim sProcID     As String
Dim typErrInfo  As tErrInfo

Dim ofrmSendVersa3  As frmSendVersa3

Dim vIPN As Variant
Dim vLotID As Variant

'----
' Init
'----
    sProcID = "cmdSendVersa3_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    
'----
' Action
'----
    ' <Put your Action codes here>...
    Me.spdLotList.Row = 1
    Do While Me.spdLotList.Row <= Me.spdLotList.MaxRows
        Me.spdLotList.Col = miSpdFieldPos_Check
        If Me.spdLotList.Text = "1" Then
            Me.spdLotList.GetText miSpdFieldPos_LotId, Me.spdLotList.Row, vLotID
            Me.spdLotList.GetText miSpdFieldPos_IPN, Me.spdLotList.Row, vIPN
            Exit Do
        End If
        Me.spdLotList.Row = Me.spdLotList.Row + 1
    Loop

    If vLotID = "" Then GoTo ExitHandler
    
    Set ofrmSendVersa3 = New frmSendVersa3
    Load ofrmSendVersa3
    With ofrmSendVersa3
        Set .CwMbxControl = moCwMbx
        Set .FwMDLControl = moFwMDL
        Set .FwWIPControl = moFwWIP
        Set .FwOPRControl = moFwOPR
        Set .FwPRPControl = moFwPRP
        Set .FwWFControl = moFwWF
        Set .FwCATControl = moFwCAT
        Set .MainTraceLog = moAppLog
        Set .CwMbxControl = moCwMbx
        Set .ProRawSqlControl = moProRawSql
            .RuleName = msRuleName
            .LotID = vLotID
            .EqpId = Me.cboEqpID.Text
            .txtEQId = Me.cboEqpID.Text
            .IPN = vIPN
            .Init
            .Show vbModal
        
        
    End With
    Unload ofrmSendVersa3
    

'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
           Case glERR_INVALIDOBJECT
               ' Retry code goes here...
           Case Else
                 typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
        End Select
        '---- Start of Select Case Block ----
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Sub

Private Sub Form_Load()
'**************************************************
'**************************************************
    On Error Resume Next
    miResult = vbCancel
    Call ResetFwControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call ResetFwControls
End Sub




'================================================================================
' Function: Init()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent Wei, CIT 2002/05/07
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
' RETURN TYPE
'   Boolean         (R) True/False
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
' [REV 02] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
'================================================================================
Public Function Init() As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim lCol As Integer
'----
' Init
'----
    sProcID = "Init"
    Init = True
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    Dim iIdx As Integer
    Dim oEqp As FwEquipment
    
    With Me.spdLotList
        .MaxCols = 14 'Modify by Sam on 20101026 for ReqNo:JC201000238,10->11, 20101215, 11 -> 14
        .MaxRows = 0
        .Protect = True
        For iIdx = 3 To .MaxCols
            .ColUserSortIndicator(iIdx) = ColUserSortIndicatorAscending
            .UserColAction = UserColActionSort
        Next iIdx
        .Col = miSpdFieldPos_ReasonCode
        .colHidden = True
        .Col = miSpdFieldPos_DefectQty
        .colHidden = True
        
        'Add by Sam start on 20101026 for ReqNo:JC201000238
        .SetText miSpdFieldPos_SumHold, 0, msSpdFieldPos_SumHold
        .Col = miSpdFieldPos_SumHold
        .colHidden = True
        'Add by Sam end on 20101026 for ReqNo:JC201000238
        
        'Add by Sam start on 20101026 for ReqNo:JC201000277
        .SetText miSpdFieldPos_CrackChoose, 0, msSpdFieldPos_CrackChoose
        .Col = miSpdFieldPos_CrackChoose
        .TypeMaxEditLen = 1024
        .colHidden = True

        .SetText miSpdFieldPos_CrackComment, 0, msSpdFieldPos_CrackComment
        .Col = miSpdFieldPos_CrackComment
        .TypeMaxEditLen = 1024
        .colHidden = True

        .SetText miSpdFieldPos_CrackHold, 0, msSpdFieldPos_CrackHold
        .Col = miSpdFieldPos_CrackHold
        .colHidden = True
        'Add by Sam end on 20101026 for ReqNo:JC201000277
                
    End With
    
    'Add by Jeff, 2002/08/18; For TimeStamp check
    Call SetTimeStampColumn(Me.spdLotList, True)
    'End of add

    Me.cmdOk.Enabled = False
    Me.cmdLotCancel.Enabled = False
    
    If msEqpId = "" Then
        Call AddEqpIDtoCombo(moFwMDL, moFwWIP, moFwWF, cboEqpID, msEqpType, msEqpId)
    Else
        Set oEqp = FwuRetrieveEqp(moFwMDL, msEqpId)
        If oEqp.Type = msEqpType Or oEqp.Type = msEqpType2 Or oEqp.Type = msEqpType3 Then
            Me.Caption = oEqp.Type & " LotComplete"
            Call AddEqpIDtoCombo(moFwMDL, moFwWIP, moFwWF, cboEqpID, oEqp.Type, msEqpId)
        Else
            Call AddEqpIDtoCombo(moFwMDL, moFwWIP, moFwWF, cboEqpID, msEqpType, msEqpId)
        End If
        Set oEqp = Nothing
    End If
    
    '20060601, vencent M200605029, add label print function
    Call GetPrinterServers(moAppLog, moFwWIP, moFwWF, moCwMbx, Me.cboPrintServer)
    'set default value
    For iIdx = 0 To Me.cboPrintServer.ListCount - 1
        If Me.cboPrintServer.List(iIdx) = "FT-TR" Then
            Me.cboPrintServer.ListIndex = iIdx
        End If
    Next iIdx
    
    'Lable spec cbmbobox
    Dim sSQL As String
    Dim colRS As Collection, oItem As Object
    sSQL = "select " & gsCAT_TLS_LABEL_SPECNO & _
           " from " & gsCAT_TBL_LABEL_SPEC & _
           " where " & gsCAT_TLS_STAGE & " = '" & gsSTAGE_FT & "'" & _
           " and " & gsCAT_TLS_DELETE_FLAG & " <> 'Y'"
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    For Each oItem In colRS
        Me.cboLabelSpec.AddItem oItem.Item(1)
        
        If oItem.Item(1) = gsLABEL_FT_SMALL_LABEL Then
            Me.cboLabelSpec.ListIndex = Me.cboLabelSpec.ListCount - 1
        End If
    Next
    
    lblGenChildLot.Caption = "" 'Add by Sam on 20110308 for ReqNo:JC201100064
    
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>

ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
            Case glERR_INVALIDOBJECT
                ' Retry code goes here...
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
        End Select
        '---- Start of Select Case Block ----
        Init = False
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function


'================================================================================
' Sub: DispLotData()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent Wei, CIT 2002/04/10
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a Object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
' [REV 02] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
'================================================================================
Private Sub DispLotData()
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim typErrInfo As tErrInfo
    '----
    ' Init
    '----
    sProcID = "DispLotData"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

    '----
    ' Condition Checking
    '----
    ' <Put your condition checking codes here>...

    '----
    ' Action
    '----
    ' <Put your Action codes here>...
    Dim colLotInfo As clsRawSqlRecordset
    Dim oEqp As FwEquipment
    Dim iIdx As Integer
    Dim sResourceType As String
    Dim sGroup As Variant
    Dim sEBoardComment As String 'add by Sam on 20100517  for Project e-board
    
    Dim bFutActWithNoHold As Boolean  'Add by Ernest on 2018/05/21 for BE工業 3.5 Phase 18 - 紅燈減量

    spdLotList.MaxRows = 0
    
    If cboEqpID.Text = "" Then GoTo ExitHandler
    
    Set oEqp = FwuRetrieveEqp(moFwMDL, Me.cboEqpID.Text, moAppLog)
    
    If oEqp Is Nothing Then GoTo ExitHandler
    
    sResourceType = oEqp.Type
    For Each sGroup In oEqp.groups
        If sGroup <> oEqp.Type Then sResourceType = sResourceType & "','" & sGroup
    Next sGroup
    
    Set colLotInfo = RetrieveFTLotInfo(moCwMbx, moAppLog, moFwWF.ClientID, sResourceType, _
                     gsLOTSTATUS_BREAK & "','" & gsLOTSTATUS_ERUN & "','" & _
                     gsLOTSTATUS_RUN, msRuleName, True, oEqp.Id)

    With Me.spdLotList
        Do While colLotInfo.EOF = False
            .MaxRows = .MaxRows + 1
            
            .SetText miSpdFieldPos_LotId, .MaxRows, colLotInfo.Value("LOTID")
            .SetText miSpdFieldPos_IPN, .MaxRows, colLotInfo.Value(gsLOT_CUSTOMATTR_IPN)
            'add by Nelson start on 2003/12/31 for ReqNo:M200312033
            .SetText miSpdFieldPos_CustomerName, .MaxRows, GetCustomerName(colLotInfo.Value(gsLOT_CUSTOMATTR_IPN), moAppLog, moProRawSql)
            'add by Nelson end on 2003/12/31 for ReqNo:M200312033
            .SetText miSpdFieldPos_LotQty, .MaxRows, colLotInfo.Value(gsLOT_CUSTOMATTR_CHIP_QTY)
            
            
            spdLotList.Col = miSpdFieldPos_Reason
            spdLotList.CellType = CellTypeEdit
            spdLotList.Lock = True
            
            .SetText .MaxCols, .MaxRows, colLotInfo.Value("FWTIMESTAMP")


            'Add by Sam on 20100517 for Project E-BOARD
            sEBoardComment = modCustom.GetLotEBoardInfo(moAppLog, moFwWIP, moFwWF, moCwMbx, moFwMDL, colLotInfo.Value("LOTID"), oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2), "", oEqp.Id)
        
            'Add by Ernest on 2018/05/21 for BE工業 3.5 Phase 18 - 紅燈減量------start
            '原先三種FutAct沒有區分出Hold, 呼叫模組進行區分
            bFutActWithNoHold = False
            If colLotInfo.Value("LOTFA") <> "" Or _
                colLotInfo.Value("IPNFA") = "Y" Or _
                colLotInfo.Value("PRODGROUPFA") = "Y" Then
                bFutActWithNoHold = modCustomUpdate.CheckFutActWithNoHold(moAppLog, moProRawSql, colLotInfo.Value("LOTID"))
            End If
            'Add by Ernest on 2018/05/21 for BE工業 3.5 Phase 18 - 紅燈減量-------end
            'Modify by Sam on 20100518 for Project E-BOARD ,Add sEBoardComment <> ""
            'modify by Ernest on 10150713 for ReqNo:JC201500222
          '  If colLotInfo.Value("TECN") <> "" Or _
               colLotInfo.Value("LOTFA") <> "" Or _
                colLotInfo.Value("IPNFA") = "Y" Or _
                colLotInfo.Value("StepCom") <> "" Or _
                colLotInfo.Value("EqCom") <> "" Or _
                colLotInfo.Value("ERUNTICNO") <> "" Or _
                colLotInfo.Value("SAPRWNO") <> "" Or _
                sEBoardComment <> "" Or _
                colLotInfo.Value("PRODGROUPFA") = "Y" Then
                
            'Modify by Ernest on 2018/05/21 for BE工業 3.5 Phase 18 - 紅燈減量-------start
            '移除sTecn以及將三種FutAct改為bFutActWithNoHold
            'If colLotInfo.Value("LOTFA") <> "" Or _
            '    colLotInfo.Value("IPNFA") = "Y" Or _
            '    colLotInfo.Value("StepCom") <> "" Or _
            '    colLotInfo.Value("EqCom") <> "" Or _
            '    colLotInfo.Value("ERUNTICNO") <> "" Or _
            '    colLotInfo.Value("SAPRWNO") <> "" Or _
            '    sEBoardComment <> "" Or _
            '    colLotInfo.Value("PRODGROUPFA") = "Y" Then
                
            '    If colLotInfo.Value("LOTFA") <> "" Or _
            '     colLotInfo.Value("IPNFA") = "Y" Or _
            '     colLotInfo.Value("StepCom") <> "" Or _
            '     colLotInfo.Value("EqCom") <> "" Or _
            '     sEBoardComment <> "" Or _
            '     colLotInfo.Value("PRODGROUPFA") = "Y" Then
                 
            If bFutActWithNoHold = True Or _
                colLotInfo.Value("StepCom") <> "" Or _
                colLotInfo.Value("EqCom") <> "" Or _
                colLotInfo.Value("ERUNTICNO") <> "" Or _
                colLotInfo.Value("SAPRWNO") <> "" Or _
                sEBoardComment <> "" Then
                
                If bFutActWithNoHold = True Or _
                 colLotInfo.Value("StepCom") <> "" Or _
                 colLotInfo.Value("EqCom") <> "" Or _
                 sEBoardComment <> "" Then
                 
            'Modify by Ernest on 2018/05/21 for BE工業 3.5 Phase 18 - 紅燈減量-------end
                 
                 .Col = miSpdFieldPos_TECN
                 .Row = .MaxRows
                 .Text = "red light"
                 .CellType = CellTypeButton
                 .TypeButtonPicture = frmImageList.imglistLot.ListImages.Item("TECN").Picture
                 Else
                 .Col = miSpdFieldPos_TECN
                 .Row = .MaxRows
                 .CellType = CellTypeButton
                 .TypeButtonPicture = frmImageList.imglistLot.ListImages.Item("SpecialTicNO").Picture
                End If
            End If
            
            colLotInfo.MoveNext
        Loop
    End With
    Me.spdLotList.Sort 1, 1, Me.spdLotList.MaxCols, Me.spdLotList.MaxRows, SortByRow, miSpdFieldPos_LotId

    Call SetSpdColWidth(spdLotList, 1, , moAppLog)
    
    'disable button
    Me.cmdOk.Enabled = False
    
    Me.txtFail = ""
    Me.txtRemain = ""
    '----
    ' Done
    '----
    
    
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>
    Set colLotInfo = Nothing
    Set oEqp = Nothing
    
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
            Case glERR_INVALIDOBJECT
                ' Retry code goes here...
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                        "程式執行失敗, 請洽IT人員處理"
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Sub

'================================================================================
' Sub: spdLotList_ButtonClicked()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent Wei, CIT 2002/04/12
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   Argument1           (I) <Description goes here...>
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
' [REV 02] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
'================================================================================
Private Sub spdLotList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sLotID As Variant
    Dim oLot As FwLot
    Dim sSQL As String
    Dim colSQLResult As Collection
    Dim iIdx As Integer
    Dim sCurLotId As String
    Dim sLotQty As String
    
    
    'Add by Sam start on 20101215 for ReqNo:JC201000277
    Dim sCrackFlag  As String
    Dim sCrackComment As String
    Dim sCrackChoose  As String
    Dim frmCrack      As New frmFtVMCrackConfirm
    'Add by Sam end on 20101215 for ReqNo:JC201000277
    
    Dim sChildStepID As String
    
    Dim vMerge As Variant  'add  by Ernest on 20150714 for ReqNo:201500222
    Dim frmTmp As Variant  'add  by Ernest on 20150714 for ReqNo:201500222
    
    '----
    ' Init
    '----
    sProcID = "spdLotList_ButtonClicked"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    Me.spdLotList.Row = Row
    Call Me.spdLotList.GetText(miSpdFieldPos_LotId, Row, sLotID)
    Set oLot = FwuRetrieveLot(moFwWIP, CStr(sLotID))
    sCurLotId = oLot.Id
    sChildStepID = oLot.CurrentStep.Steps.Item(1).Id 'Add by Sam on 20150622 for ReqNo:JC201500175
    '----
    ' Action
    '----
    
    Select Case Col
        Case miSpdFieldPos_Check:
        
            Me.spdLotList.Col = Col
        
            If ButtonDown = 1 Then   'checked
                If Len(oLot.CustomAttributes(gsLOT_CUSTOMATTR_LOC_ID)) > 0 Then
                    UtShowMsgBox "此Lot在儲位內，請先執行下架作業", , , True
                    Me.spdLotList.SetInteger Col, Row, 0
                Else
                    mblotSelected = True
                    msLotId = sCurLotId
                    Set moDefectQty = New Collection
                    Set moReasonCode = New Collection
                    
                    ' lock other check box
                    For iIdx = 1 To Me.spdLotList.MaxRows
                        If iIdx <> Row Then
                            Me.spdLotList.Col = miSpdFieldPos_Check
                            Me.spdLotList.Row = iIdx
                            Me.spdLotList.Lock = True
                        End If
                    Next iIdx
                    
                    Me.spdLotList.Row = Row
                    Me.spdLotList.Col = miSpdFieldPos_LotQty
                    sLotQty = Me.spdLotList.Text
                    
                    Me.spdLotList.Col = miSpdFieldPos_FailQty
                    Me.txtFail = Me.spdLotList.Text
                    
                    
                    Me.spdLotList.Col = miSpdFieldPos_Reason
                    Me.spdLotList.CellType = CellTypeButton
                    Me.spdLotList.TypeButtonText = "Fail Reason"
                    Me.spdLotList.Lock = False
                    
                    sSQL = " select " & gsCAT_TPBS_CARRIER_QTY & _
                           " from " & gsCAT_TBL_PRM_BE_SPEC & _
                           " where " & gsCAT_TPBS_IPN & " = '" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "' and " & _
                                       gsCAT_TPBS_DEFAULTS & " = 'Y' and " & _
                                       gsCAT_TPBS_DELETE_FLAG & " <> 'Y'"
                    Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
                    If colSQLResult.Count > 0 Then
                        Me.txtTRQty = colSQLResult.Item(1).Item(1)
                        If Val(Me.txtTRQty) <> 0 Then
                            Me.txtReelCount = CStr((Val(sLotQty) - Val(Me.txtFail)) \ Val(Me.txtTRQty))
                        End If
                    End If
                    'add by Ernest on 20150710 for ReqNo:JC201500222-------------------------<Start>
            If Me.spdLotList.GetText(miSpdFieldPos_TECN, Row, vMerge) = True Then

            Set frmTmp = New frmTECN
            Load frmTmp
            With frmTmp
                Set .CwMbxControl = moCwMbx
                Set .FwMDLControl = moFwMDL
                Set .FwWIPControl = moFwWIP
                Set .FwOPRControl = moFwOPR
                Set .FwPRPControl = moFwPRP
                Set .FwWFControl = moFwWF
                Set .FwCATControl = moFwCAT
                Set .MainTraceLog = moAppLog
                Set .CwMbxControl = moCwMbx
                Set .ProRawSqlControl = moProRawSql
                    .RuleName = msRuleName
                    .LotID = sCurLotId
                    .EqpId = Me.cboEqpID.Text

                    .Init
                    .Show vbModal
            End With
            
            Unload frmTmp
            End If
        '------------------------------------------------------------------------------<End>



                    If Me.txtFail = "" Then Me.txtFail = 0
                    
                    
                    Me.txtTotalReeledCount = CStr(Val(Me.txtReelCount) * Val(Me.txtTRQty))
                    Me.txtRemain = CStr(Val(sLotQty) - Val(Me.txtTotalReeledCount) - Val(Me.txtFail))
                    
    '                Me.lblChildLot.Visible = True
    '                Me.lblChildLotLabel.Visible = True
    
                    Me.cmdLotCancel.Enabled = True
                    
                    cboMergeLots.Clear 'Add by Sam on 20150525 for ReqNo:JC201500175 ,將原本的Clear從GetMergeLots移出來
                    
                    'Added by Jack on 2012/12/17 for 併批捲/箱 批號管制專案 Project <Start>
                    '若 Lotid 為 *M* , 則需先串出其母批, 再以母批的Lotid 去串當站的Merge List
                    'else 以畫面上的Lotid 串出當站的Merge List
                    
                    'Modify by Weilun on 20171013 for Remark 限制系統建立
                    '原先寫死第9碼為M, 改由利用模組到Table中搜尋
                    'If sCurLotId Like "*M?" Then
                    If CheckFviMergeLotID(sCurLotId, moProRawSql, moAppLog) Then
                        sSQL = "select FSL.WIPID PARENTLOT," & vbNewLine & _
                                    "      FSL.ACTIVITY," & vbNewLine & _
                                    "      FSL.TXNTIMESTAMP TXNTIME," & vbNewLine & _
                                    "      FS.CHILDLOTID CHILDLOT," & vbNewLine & _
                                    "      FSLM.VALDATA STEPID," & vbNewLine & _
                                    "      C.DESCRIPTION STEPNAME," & vbNewLine & _
                                    "      D.VALDATA as IPN," & vbNewLine & _
                                    "      FS.SPLITQTY QTY," & vbNewLine & _
                                    "      FSL.SYSID" & vbNewLine & _
                                    "  from FWSPLITLOT      FSL," & vbNewLine & _
                                    "      FWSPLITLOT_N2M  FSLN," & vbNewLine & _
                                    "      FWSPLIT         FS," & vbNewLine & _
                                    "      FWSPLITLOT_PN2M FSLM," & vbNewLine & _
                                    "      (select STEPNAME, DESCRIPTION from FWSTEPVERSION where REVSTATE = 'Active') C," & vbNewLine & _
                                    "      (select FROMID, VALDATA from FWLOT_PN2M where KEYDATA = 'IPN') D" & vbNewLine & _
                                    " where FSL.SYSID = FSLN.FROMID and" & vbNewLine & _
                                    "       FSLN.TOID = FS.SYSID and" & vbNewLine & _
                                    "      FSL.SYSID = FSLM.FROMID and" & vbNewLine & _
                                    "       FSLM.KEYDATA = 'stepId' and" & vbNewLine
                        sSQL = sSQL & "  substr(FSLM.VALDATA, 1, 5) = C.STEPNAME and" & vbNewLine & _
                                      "  FSL.LOTOBJECT = D.FROMID "
                        sSQL = sSQL & _
                                " and FSL.SYSID in (" & _
                                "select A.PARENTTXN " & vbNewLine & _
                            "from FWSPLITLOT A " & vbNewLine & _
                            "where A.WIPID = '" & Trim(sCurLotId) & "' " & vbNewLine & _
                            "and A.ACTIVITY = 'Split'" & vbNewLine & _
                            "and A.PARENTTXN is not null" & vbNewLine & _
                                 ") " & vbNewLine & _
                                " and FS.CHILDLOTID = '" & Trim(sCurLotId) & "'"
                        
                        sSQL = sSQL & " and substr(FSLM.VALDATA,1,5) = '" & sChildStepID & "' "  'Add by Sam on 20150622 for ReqNo:JC201500175,限制StepID需與子批相同
                        
                        Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
                        If colSQLResult.Count > 0 Then
                            sCurLotId = colSQLResult.Item(1).Item(1)
                            'Add by Sam Start on 20150525 for ReqNo:JC201500175 ,將MX Lot取得的母批帶入清單
                            cboMergeLots.AddItem sCurLotId
                            Call GetMergeLots(sCurLotId)
                            'Add by Sam End on 20150525 for ReqNo:JC201500175
                        End If
                    Else
                        cboMergeLots.AddItem sCurLotId 'Add by Sam  on 20150703 for ReqNo:JC201500219 ,非M* LOT , 將原批號加入下拉選單
                    End If
                    'Added by Jack on 2012/12/17 for 併批捲/箱 批號管制專案 Project <End>
                                                                                                                                                
                    sCurLotId = oLot.Id 'Add by Sam Start on 20150525 for ReqNo:JC201500175 , 以原本的Lot取清單
                    Call GetMergeLots(sCurLotId) 'Add by Sam on 20110308 for ReqNo:JC201100064

                End If
                
            Else    'unehecked
                mblotSelected = False
                Set moDefectQty = Nothing
                Set moReasonCode = Nothing
                    
                For iIdx = 1 To Me.spdLotList.MaxRows
                    If iIdx <> Row Then
                        Me.spdLotList.Row = iIdx
                        Me.spdLotList.Lock = False
                    End If
                Next iIdx
                Me.txtRemain = ""
                Me.txtFail = ""
                Me.txtTotalReeledCount = ""
'                Me.txtRemain.BackColor = &H8000000B
'                Me.txtRemain.Enabled = False

                Me.spdLotList.Row = Row
                Me.spdLotList.Col = miSpdFieldPos_Reason
                Me.spdLotList.CellType = CellTypeEdit
                Me.spdLotList.Lock = True
                Me.spdLotList.Col = miSpdFieldPos_FailQty
                Me.spdLotList.Text = ""

                Me.cmdOk.Enabled = False
                Me.cmdLotCancel.Enabled = False
                Me.lblChildLot.Visible = False
                Me.lblChildLotLabel.Visible = False
                
                Me.cboMergeLots.Visible = lblChildLotLabel.Visible  'Add by Sam on 20110308 for ReqNo:JC201100064
                
            End If
        
        Case miSpdFieldPos_TECN:
            'Dim frmTmp As New frmTECN  modify by Ernest on 20150714 for ReqNo:201500222
            Set frmTmp = New frmTECN
            Load frmTmp
            With frmTmp
                Set .CwMbxControl = moCwMbx
                Set .FwMDLControl = moFwMDL
                Set .FwWIPControl = moFwWIP
                Set .FwOPRControl = moFwOPR
                Set .FwPRPControl = moFwPRP
                Set .FwWFControl = moFwWF
                Set .FwCATControl = moFwCAT
                Set .MainTraceLog = moAppLog
                Set .CwMbxControl = moCwMbx
                Set .ProRawSqlControl = moProRawSql
                    .RuleName = msRuleName
                    .LotID = sCurLotId
                    .EqpId = Me.cboEqpID.Text

                    .Init
                    .Show vbModal
            End With
            
            Unload frmTmp
        
        Case miSpdFieldPos_Reason
            Dim frmLoss As New frmDefect
            Load frmLoss
            With frmLoss
                Set .CwMbxControl = moCwMbx
                Set .FwMDLControl = moFwMDL
                Set .FwWIPControl = moFwWIP
                Set .FwOPRControl = moFwOPR
                Set .FwPRPControl = moFwPRP
                Set .FwWFControl = moFwWF
                Set .FwCATControl = moFwCAT
                Set .MainTraceLog = moAppLog
                Set .CwMbxControl = moCwMbx
                Set .ProRawSqlControl = moProRawSql
                    .RuleName = msRuleName
                    .LotID = sCurLotId
                    .EqpId = Me.cboEqpID.Text
                    
                    Me.spdLotList.Row = Row
                    Me.spdLotList.Col = miSpdFieldPos_FailQty
                    .txtLossQty = Me.spdLotList.Text
                    .lblLossQty.Visible = False
                    .txtLossQty.Visible = False
                    
                    'Mark by Sam start on 20101028 for ReqNo:JC201000238
'                    .txtGoodQty = "0"
'                    .txtGoodQty.Visible = False
'                    .lblGoodQty.Visible = False
                    'Mark by Sam End on 20101028 for ReqNo:JC201000238
                    
                    .txtFailQty = "0"
                    .txtFailQty.Visible = False
                    .lblFailQty.Visible = False

                    .Init
                    .InitOldReason moReasonCode, moDefectQty
                    
                     'Add by Sam start on 20101028 for ReqNo:JC201000238
                    .txtGoodQty = Val(txtTRQty.Text) * Val(txtReelCount.Text)
                    .cmdOk.Enabled = True
                    .txtGoodQty.Locked = True
                     'Add by Sam End on 20101028 for ReqNo:JC201000238
                     
                    .Show vbModal
                If .Result = vbOK Then
                    'Add by Sam start on  20110419 for ReqNo:JC201100108
                    Me.txtGoodQty.Text = Val(.txtGoodQty.Text)
                    'Add by Sam end on  20110419 for ReqNo:JC201100108
                    
                    'Add by Sam start on 20101215 for ReqNo:JC201000277
                    For iIdx = 1 To .ReasonCode.Count
                        If .ReasonCode.Item(iIdx) = "Crack" Or _
                           .ReasonCode.Item(iIdx) = "Package break" Or _
                           .ReasonCode.Item(iIdx) = "Chip out" Then
                                sCrackFlag = "True"
                                Exit For
                        End If
                    Next iIdx
                    
                    If sCrackFlag = "True" Then
                    
                        Load frmCrack
                            frmCrack.Init
                            frmCrack.Show vbModal
                            If .Result <> vbOK Then
                                GoTo ExitHandler
                            End If
                            
                            sCrackComment = frmCrack.Comment
                            sCrackChoose = frmCrack.Choose
                            
                            Me.spdLotList.SetText miSpdFieldPos_CrackChoose, Row, sCrackChoose
                            Me.spdLotList.SetText miSpdFieldPos_CrackComment, Row, sCrackComment
                            
                            Unload frmCrack
                                            
                        If (InStr("ABC", sCrackChoose)) Then
                            Me.spdLotList.SetText miSpdFieldPos_CrackHold, Row, "Y"
                        End If
                    End If
                    'Add by Sam End on 20101215 for ReqNo:JC201000277
                
                
                    Dim iDefectQty As Long
                    Dim sDefectQty As String, sReasonCode As String
                    
                    If Me.txtFail = "" Then Me.txtFail = "0"
                    
                    Set moDefectQty = .DefectQty
                    Set moReasonCode = .ReasonCode
                    
                    For iIdx = 1 To moDefectQty.Count
                        sDefectQty = sDefectQty & moDefectQty.Item(iIdx) & ";"
                        sReasonCode = sReasonCode & moReasonCode.Item(iIdx) & ";"
                    Next
                    If Right(sDefectQty, 1) = ";" Then
                        sDefectQty = Left(sDefectQty, Len(sDefectQty) - 1)
                        sReasonCode = Left(sReasonCode, Len(sReasonCode) - 1)
                    End If
                    
                    Me.spdLotList.SetText miSpdFieldPos_ReasonCode, Row, sReasonCode
                    Me.spdLotList.SetText miSpdFieldPos_DefectQty, Row, sDefectQty
                                                                
                    msHoldCommentPE = "" 'Added by Jack on 2025/06/04 for #212677_BE#202400097 MES FVI Lot complete時, 如有: FVI Test type Error , 加Hold MK330
                    
                    'Add by Sam start on 20101026 for ReqNo:JC201000238
                    If .txtHoldFlag.Text = "True" Then
                        Me.spdLotList.SetText miSpdFieldPos_SumHold, Row, "Y"
                        msHoldComment = .txtHoldComment
                        'Added by Jack on 2025/06/04 for #212677_BE#202400097 MES FVI Lot complete時, 如有: FVI Test type Error , 加Hold MK330 <Start>
                        If .txtpeisstatus.Text = "FVI Test type Error" Then
                            msHoldCommentPE = "FVI Test type Error"
                        End If
                        'Added by Jack on 2025/06/04 for #212677_BE#202400097 MES FVI Lot complete時, 如有: FVI Test type Error , 加Hold MK330 <End>
                    Else
                        Me.spdLotList.SetText miSpdFieldPos_SumHold, Row, "N"
                        msHoldComment = ""
                    End If
                    'Add by Sam End on 20101026 for ReqNo:JC201000238
                                        
                    For iIdx = 1 To moDefectQty.Count
                        iDefectQty = iDefectQty + moDefectQty.Item(iIdx)
                    Next iIdx
                    
                    Me.spdLotList.SetText miSpdFieldPos_FailQty, Row, iDefectQty
                    
                    iDefectQty = 0
                    Me.spdLotList.Row = 1
                    Do While Me.spdLotList.Row <= Me.spdLotList.MaxRows
                        Me.spdLotList.Col = miSpdFieldPos_Check
                        If Me.spdLotList.Text = "1" Then
                            Me.spdLotList.Col = miSpdFieldPos_FailQty
                            If IsNumeric(Me.spdLotList.Text) = True Then
                               'Modify by Janus, 20060808, change integer to long
                                'iDefectQty = iDefectQty + CInt(Me.spdLotList.Text)
                                iDefectQty = iDefectQty + CLng(Me.spdLotList.Text)
                               'Modify end by Janus, 20060808, change integer to long
                            End If
                        End If
                        Me.spdLotList.Row = Me.spdLotList.Row + 1
                    Loop
                    
                    Me.txtFail = iDefectQty
                    
                    Call Qty_Change 'Add by Sam  on 20101026 for ReqNo:JC201000238

                End If
            End With
            
            Unload frmLoss
    End Select
    
    '----
    ' Done
    '----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' Cleaning up
    Set oLot = Nothing

ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
            Case glERR_INVALIDOBJECT
                ' Retry code goes here...
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                        "程式執行失敗, 請洽IT人員處理"
        End Select
        '---- Start of Select Case Block ----
        Screen.MousePointer = vbDefault

        miResult = vbCancel
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Sub




Private Sub Qty_Change()
    On Error Resume Next
    Dim sLotQty As String
'    Dim sOldSplitLotID As String, sNewSPlitLotID As String
    
    Dim vSumHold As Variant 'Add by Sam on 20101028 for ReqNo:JC201000238
    
    
    If mblotSelected = False Or Val(Me.txtTRQty) = 0 Then
        Me.cmdOk.Enabled = False
        Exit Sub
    End If
    
    With Me.spdLotList
        .Row = 1
        Do While .Row <= .MaxRows
            .Col = miSpdFieldPos_Check
            If .Text = "1" Then
                 'Add by Sam Start on 20101028 for ReqNo:JC201000238
                 .Col = miSpdFieldPos_SumHold
                 vSumHold = .Text
                 'Add by Sam End on 20101028 for ReqNo:JC201000238
                .Col = miSpdFieldPos_LotQty
                sLotQty = .Text
                Exit Do
            End If
            .Row = .Row + 1
        Loop
    End With
    
    If Val(Me.txtReelCount) * Val(Me.txtTRQty) + Val(Me.txtFail) > Val(sLotQty) Or _
    Val(Me.txtReelCount) = 0 Then
        Me.txtReelCount = (Val(sLotQty) - Val(Me.txtFail)) \ Val(Me.txtTRQty)
    End If
    
    Me.txtTotalReeledCount = CStr(Val(Me.txtReelCount) * Val(Me.txtTRQty))
    Me.txtRemain = CStr(Val(sLotQty) - Val(Me.txtTotalReeledCount) - Val(Me.txtFail))
    If Val(Me.txtRemain) > 0 And Val(Me.txtTotalReeledCount.Text) > 0 Then   'Modify by Sam on 20110317 for ReqNo:JC201100064,增加判斷要有實箱數
        If lblChildLot.Visible = False Then
            If Not lblChildLot.Caption Like Left(msLotId, 8) & "*" Then
                lblChildLot.Caption = GetChildLotId(moAppLog, moFwWIP, moFwWF, moCwMbx, msLotId, gsSTAGE_FT)
            End If
            lblChildLot.Visible = True
            lblChildLotLabel.Visible = True
        End If
    Else
            lblChildLot.Visible = False
            lblChildLotLabel.Visible = False
    End If
    'Modify by Sam on 20101028 for ReqNo:JC201000238
    'If Val(Me.txtTotalReeledCount) > 0 Then
    If Val(Me.txtTotalReeledCount) > 0 And vSumHold <> "" Then
        Me.cmdOk.Enabled = True
    Else
        Me.cmdOk.Enabled = False
    End If
    
    'Add by Sam start on  20110419 for ReqNo:JC201100108
    If Val(txtTRQty.Text) * Val(txtReelCount) <> Val(Me.txtGoodQty) Then
        Me.cmdOk.Enabled = False
    End If
    'Add by Sam end on  20110419 for ReqNo:JC201100108
    
    'Add by Sam Start on 20110308 for ReqNo:JC201100064
    If Val(txtRemain.Text) > 0 And Trim(lblGenChildLot.Caption) = "" Then
         Me.cmdOk.Enabled = False
    End If
    
    cboMergeLots.Visible = lblChildLotLabel.Visible
    cmdGenLotID.Visible = lblChildLotLabel.Visible
    lblGenChildLot.Visible = lblChildLotLabel.Visible
    
    'Add by Sam End on 20110308 for ReqNo:JC201100064
End Sub

Private Sub txtFail_Change()
    Call Qty_Change
End Sub

Private Sub txtReelCount_Change()
    Call Qty_Change
    cmdOk.Enabled = False 'Add by Sam on 20101108 for ReqNo:JC201000238
End Sub

Private Sub txtTRQty_Change()
    Call Qty_Change
    cmdOk.Enabled = False 'Add by Sam on 20101108 for ReqNo:JC201000238
End Sub


'================================================================================
' Sub: GetMergeLots()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen, CIT 2011/03/08
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
' RETURN TYPE
'   Boolean         (R) True/False
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
' [REV 02] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
'================================================================================
Public Sub GetMergeLots(ByVal sLotID As String)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

Dim oLot As FwLot
Dim sSQL As String
Dim colRS As Collection

Dim iIdx As Integer
Dim sMergeGrouphistkey As String

'----
' GetMergeLots
'----
    sProcID = "GetMergeLots"

    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)

    Set oLot = FwuRetrieveLot(moFwWIP, sLotID, moAppLog)
    
'    cboMergeLots.Clear 'Mark by Sam on 20150525 for ReqNo:JC201500175
    
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----

'    cboMergeLots.AddItem sLotID 'Mark by Sam on 20150525 for ReqNo:JC201500175
    
   ' <Put your Action codes here>...
   'Mark by Sam Start on 20150525 for ReqNo:JC201500175,取消Grouphistkey的條件
'    '取得最後一次Merge的Grouphistkey
'   sSQL = " select A.WIPID PARENTLOT," & vbNewLine & _
'          " A.ACTIVITY," & vbNewLine & _
'          " A.TXNTIMESTAMP TXNTIME," & vbNewLine & _
'          " B.VALDATA CHILDLOT," & vbNewLine & _
'          " A.MERGESTEPID STEPID," & vbNewLine & _
'          " C.GROUPHISTKEY " & _
'            " from FWMERGE A," & vbNewLine & _
'                 " FWMERGE_PN2M B," & vbNewLine & _
'                 " Fwwiphistory C" & vbNewLine & _
'            " where A.SYSID = B.FROMID" & vbNewLine & _
'              " and A.SYSID = C.WIPTXN " & vbNewLine & _
'              " and B.LINKNAME = 'childLotCollection'" & vbNewLine & _
'              " and SUBSTR(A.MERGESTEPID, 1, 5) = '" & oLot.CurrentStep.Steps.Item(1).Id & "'" & vbNewLine & _
'              " AND A.WIPID = '" & oLot.Id & "'" & vbNewLine & _
'            " order by PARENTLOT, TXNTIME DESC "
'    Set colRS = moProRawSql.QueryDatabase(sSQL)
'    If colRS.Count > 0 Then
'        sMergeGrouphistkey = colRS.Item(1).Item("GROUPHISTKEY")
'    End If
    'Mark by Sam End on 20150525 for ReqNo:JC201500175,取消Grouphistkey的條件
    
    'Modify by Sam on 20150525 for ReqNo:JC201500175,取消Grouphistkey的條件
     '只取最後一次Merge的Grouphistkey的Lot
   sSQL = " select A.WIPID PARENTLOT," & vbNewLine & _
          " A.ACTIVITY," & vbNewLine & _
          " A.TXNTIMESTAMP TXNTIME," & vbNewLine & _
          " B.VALDATA CHILDLOT," & vbNewLine & _
          " A.MERGESTEPID STEPID," & vbNewLine & _
          " C.GROUPHISTKEY " & _
            " from FWMERGE A," & vbNewLine & _
                 " FWMERGE_PN2M B," & vbNewLine & _
                 " Fwwiphistory C" & vbNewLine & _
            " where A.SYSID = B.FROMID" & vbNewLine & _
              " and A.SYSID = C.WIPTXN " & vbNewLine & _
              " and B.LINKNAME = 'childLotCollection'" & vbNewLine & _
              " and SUBSTR(A.MERGESTEPID, 1, 5) = '" & oLot.CurrentStep.Steps.Item(1).Id & "'" & vbNewLine & _
              " AND A.WIPID = '" & oLot.Id & "'" & vbNewLine & _
            " order by PARENTLOT, TXNTIME DESC "
            
'              " AND C.GROUPHISTKEY = '" & sMergeGrouphistkey & "'" & vbNewLine & _

    Set colRS = moProRawSql.QueryDatabase(sSQL)
    For iIdx = 1 To colRS.Count
        cboMergeLots.AddItem colRS.Item(iIdx).Item("CHILDLOT")
    Next
    
    
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>

ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
            Case glERR_INVALIDOBJECT
                ' Retry code goes here...
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Sub



