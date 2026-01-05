Attribute VB_Name = "modCustomUpdate"
Option Explicit

'*********************************************************************
'*                                                                   *
'*  Create by Jack on 2012/12/07 for 併批捲/箱 批號管制專案 Project  *
'*                                                                   *
'*********************************************************************
'
' Revision History:
'................................................................................
' [REV 01] Jack, MXIC, 2012/12/07 for 併批捲/箱 批號管制專案 Project.
' 1) 新增 Function : GetCarrierMergeList
'                    傳入 ParentLotID, 抓出 Tbl_Merge_List 最後一組資料.
' 2) 新增 Function : InsertCarrierMergeList
'
' [REV 02] Weilun, MXIC, 2017/10/13 & 18 for Remark 限制系統建立
' 1) 調整 Function : GetCarrierMergeList
'                    目標從TBL_MERGE_LIST改為TBL_FVI_MERGE_LIST
' 2) 調整 Function : InsertCarrierMergeList
'                    目標從TBL_MERGE_LIST改為TBL_FVI_MERGE_LIST,
'                    並且因為欄位與規則需大幅修改, ChildLot中不該有後端併批,
'                    交由外部Rule卡關
' 3) 新增 Function : CheckFviMergeLotID
'                    傳入 LotID, moProRawSql, oLogCtrl.
'                    檢查LotID的9/10碼是否在TBL_FVI_MERGELOTNO_CONTRL的組合中
' 4) 新增 Function : CheckAssyVendorCodeForMerge
'                    傳入 ParentLotID, sChildLotID, moProRawSql, oLogCtrl.
'                    檢查ipn.brand <> 'KH' 時, 子母批AssyVendorCode是否一樣

                    

Private Const msMODULE_ID As String = "modCustomUpdate"

'================================================================================
' Sub: GetCarrierMergeList()
'--------------------------------------------------------------------------------
' Description:  for 併批捲/箱 批號管制專案 Project, 取得 MergeCarrier 的併批明細.
'--------------------------------------------------------------------------------
' Author:       Create by Jack on 2012/12/07 for 併批捲/箱 批號管制專案 Project
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a Object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] Jack, MXIC, 2012/12/07 for 併批捲/箱 批號管制專案 Project.
' 1) 新增 Function : GetCarrierMergeList
'                    傳入 ParentLotID, 抓出 Tbl_Merge_List 最後一組資料.
' [REV 02] Weilun, MXIC, 2017/10/13 for Remark 限制系統建立
' 1) 調整 Function : GetCarrierMergeList
'                    目標從TBL_MERGE_LIST改為TBL_FVI_MERGE_LIST
'
'================================================================================
Public Function GetCarrierMergeList(ByVal sParentLotID As String, ByRef oFwWIP As Object, _
                               ByRef oFwWF As Object, ByRef oCwMbx As Object, _
                               ByRef oLot As FwLot, ByRef moProRawSql As Object, _
                               Optional ByRef oLogCtrl As Object) As Collection
On Error GoTo ExitHandler:

Dim sProcID     As String
Dim typErrInfo  As tErrInfo
Dim colRaws     As Collection
Dim oRaws       As FwStrings
Dim sTable      As String
Dim sColumn     As String
Dim sWhere      As String

Dim sStage      As String
Dim sSQL        As String
Dim colRS       As Collection
Dim iIndex      As Integer

'----
' Init
'----
    sProcID = "GetCarrierMergeList"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    
'----
' Action
'----

    'Mark by Weilun on 20171013 for Remark 限制系統建立 <start>
'    sStage = "CARRIER"
'
'    sSQL = "select a." & gsCAT_TML_PARENTLOTID & ", a." & gsCAT_TML_CHILDLOTID & ", a." & gsCAT_TML_CHIPQTY & " " & vbNewLine & _
'             "from " & gsCAT_TBL_MERGE_LIST & " a " & _
'            "where a." & gsCAT_TML_PARENTLOTID & " = '" & sParentLotID & "' " & _
'              "and a.createtime = (select max(a." & gsCAT_TML_CREATETIME & ") " & _
'                                    "from " & gsCAT_TBL_MERGE_LIST & " a " & _
'                                   "where a." & gsCAT_TML_PARENTLOTID & " = '" & sParentLotID & "' " & _
'                                     "and a." & gsCAT_TML_STAGE & " = '" & sStage & "') " & _
'              "and a." & gsCAT_TML_PARENTLOTID & " <> a." & gsCAT_TML_CHILDLOTID & vbNewLine & _
'              "and a." & gsCAT_TML_STAGE & " = '" & sStage & "' "
    'Mark by Weilun on 20171013 for Remark 限制系統建立 <end>
    
    'Add by Weilun on 20171013 for Remark 限制系統建立 <start>
    sSQL = "select a." & gsCAT_TFML_PARENTLOTID & ", " & _
                 " a." & gsCAT_TFML_CHILDLOTID & ", " & _
                 " a." & gsCAT_TFML_CQTY & " " & _
             "from " & gsCAT_TBL_FVI_MERGE_LIST & " a " & _
            "where a." & gsCAT_TFML_PARENTLOTID & " = '" & sParentLotID & "' " & _
              "and a." & gsCAT_TFML_CREATETIME & " = (" & _
                    "select max(aa." & gsCAT_TFML_CREATETIME & ") " & _
                      "from " & gsCAT_TBL_FVI_MERGE_LIST & " aa " & _
                     "where aa." & gsCAT_TFML_PARENTLOTID & " = '" & sParentLotID & "') " & _
              "and a." & gsCAT_TFML_PARENTLOTID & " <> a." & gsCAT_TFML_CHILDLOTID
    'Add by Weilun on 20171013 for Remark 限制系統建立 <end>
    
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
        
    Set GetCarrierMergeList = colRaws
    
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Function

'================================================================================
' Function: InsertCarrierMergeList()
'--------------------------------------------------------------------------------
' Description:  for 併批捲/箱 批號管制專案 Project, 取得 MergeCarrier 的併批明細.
'--------------------------------------------------------------------------------
' Author:       Create by Jack on 2012/12/07 for 併批捲/箱 批號管制專案 Project
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a clsLogTraceMsg object
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
'                **************
'                *            *
' ****************  Case簡介  ****************
'                *            *
'                **************
' 在MES Phase 16 - Remark 限制功能建立後只剩Case-1, Comment by Weilun on 20171013 for Remark 限制系統建立
' <Case-1> Merge 輸入 A, 併批 B, C
'          B, C --> A (OriParent) --> AM2 (NewParent)
'          Tbl_Merge_List 寫入 : AM2, A, A (ParentLot為NewParent, ChildLot, MiddleLot為OriParent)
'                                AM2, B, A
'                                AM2, C, A
'Mark by Weilun on 20171013 for Remark 限制系統建立 <start>
'''' <Case-2> Merge 輸入 AM2, 併批 D, E
''''          D, E --> AM2 (OriParent) --> AM2 (NewParent)
''''          Tbl_Merge_List 寫入 : AM2, D, AM2 (ParentLot為NewParent, ChildLot, MiddleLot為OriParent)
''''                                AM2, E, AM2
''''          再解開 AM2 --> A,B,C
''''          Tbl_Merge_List 寫入 : AM2, A, null (ParentLot為NewParent, ChildLot, MiddleLot為null/因為OriParent=NewParent)
''''                                AM2, B, null
''''                                AM2, C, null
'''' <Case-3> Merge 輸入 F, 併批 G, AM2
''''          G, AM2 --> F (OriParent) --> FM2 (NewParent)
''''          Tbl_Merge_List 寫入 : FM2, G, F (ParentLot為NewParent, ChildLot, MiddleLot為OriParent)
''''          再解開 AM2 --> A,B,C,D,E
''''          Tbl_Merge_List 寫入 : FM2, A, F (ParentLot為NewParent, ChildLot, MiddleLot為OriParent/因為OriParent<>NewParent)
''''                                FM2, B, F
''''                                FM2, C, F
''''                                FM2, D, F
''''                                FM2, E, F
'Mark by Weilun on 20171013 for Remark 限制系統建立 <end>
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
' [REV 02] Weilun, MXIC, 2017/10/13 for Remark 限制系統建立
' 1) 調整 Function : InsertCarrierMergeList
'                    目標從TBL_MERGE_LIST改為TBL_FVI_MERGE_LIST,
'                    並且因為欄位與規則需大幅修改, ChildLot中不該有後端併批,
'                    交由外部Rule卡關
'
'Comment by Weilun on 20171013 for Remark 限制系統建立 <start>
'sSource固定為'BE', 取代stage, 由Trigger寫入時為'SUBCON'
'傳入 : <1> parentlotid, childlotid, chipqty, 其他則利用childLotId到Tbl_Lot_Info和Tbl_Lot_Attribute取得,
'       <2> CreateUserID, CreateTime (相同一組寫入相同時間)'
'Comment by Weilun on 20171013 for Remark 限制系統建立 <end>
'
'Mark by Weilun on 20171013 for Remark 限制系統建立 <start>
''''stage固定為 'CARRIER'
''''傳入 : <1> Lot相關 : parentlotid, childlotid, middlelotid, chipqty, fgipn, datecode
''''       <2> 最後一筆傳入 infoflag = 'Y', 其他傳入 'N'
''''       <3> CreateUserID, CreateTime (相同一組寫入相同時間)
'Mark by Weilun on 20171013 for Remark 限制系統建立 <end>
'
'================================================================================
Public Function InsertCarrierMergeList(ByVal sOriParentLotID As String, ByVal sNewParentLotID As String, _
                                       ByVal sChildLotIDList As String, ByVal sChildCQtyList As String, _
                                       ByVal sUserID As String, ByVal sMergeQty As String, _
                                       ByRef oFwWIP As Object, ByRef oFwWF As Object, ByRef oCwMbx As Object, _
                                       ByRef oLot As FwLot, ByRef moProRawSql As Object, _
                                       Optional ByRef oLogCtrl As Object) As Boolean
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim sSQL            As String
Dim sWaferID        As String
Dim iIdx            As Integer
Dim iIdx2           As Integer
Dim colResult       As Collection

Dim sStage             As String
Dim sInfoFlag          As String
Dim sSysTime           As String
Dim sChildLotIDArray() As String
Dim sChildCQtyArray()  As String

'Mark by Weilun on 20171013 for Remark 限制系統建立 <start>
'Dim bHasCarrierParent  As Boolean
'Dim iTotalDataCount    As Integer
'Dim iProcessDataCount  As Integer
'Dim sMiddleLotID       As String
'Dim sLastParentLotID   As String
'Dim sLastChildLotID    As String
'Dim sLsatCQty          As String
'Mark by Weilun on 20171013 for Remark 限制系統建立 <start>

'Add by Weilun on 20171013 for Remark 限制系統建立 <start>
Dim sSQLInsertColumn   As String
Dim sSQLInsertValue    As String
Dim sSource            As String
Dim sOldCreateTime     As String
'Add by Weilun on 20171013 for Remark 限制系統建立 <end>

'----
' Init
'----
    sProcID = "InsertCarrierMergeList"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)
    InsertCarrierMergeList = True
    'Mark by Weilun on 20171013 for Remark 限制系統建立 <start>
'    bHasCarrierParent = False
'    iTotalDataCount = 0
'    iProcessDataCount = 0
'    sStage = "CARRIER"
'    sInfoFlag = "N"

'    '若 AM1 為 A,B,C組成; AM1又併D,E;
'    '則 D,E  的MiddleLotID為AM1;
'    '   A,B,C的MiddleLotID為null.
'    sMiddleLotID = sOriParentLotID
    'Mark by Weilun on 20171013 for Remark 限制系統建立 <end>
    
    sSysTime = modRawSQL.GetSystemTime(oLogCtrl, oFwWIP, oFwWF, oCwMbx)
    sSource = "BE"
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    'Set colResult = moProRawSql.QueryDatabase(sSql)
    
    'Step-1 : 先將 ChildLotIDList 以 "," 切割,
    
    'Mark by Weilun on 20171013 for Remark 限制系統建立 <start>
    '子批不應存在後端併批批號中, 舊規則是'*M?', 新規則是看Table
'''    'Step-2 : 判斷 第9碼 是否有 %M_ (VB: *M?) 的 Lot?
'''    '         有 --> 再找出 Tbl_Merge_List 最新的一組資料的筆數.
'''    'Step-3 : 計算總筆數, 最後一筆 InfoFlag 寫入 "Y".
    'Mark by Weilun on 20171013 for Remark 限制系統建立 <end>
    
    sChildLotIDArray = Split(sChildLotIDList, ",")
    sChildCQtyArray = Split(sChildCQtyList, ",")
    
    'Add by Weilun on 20171013 for Remark 限制系統建立 <start>
    
    '先準備Insert欄位, 後面用併的
    sSQLInsertColumn = "insert into " & gsCAT_TBL_FVI_MERGE_LIST & " " & _
                                 "( " & gsCAT_TFML_PARENTLOTID & ", " & _
                                        gsCAT_TFML_CHILDLOTID & ", " & _
                                        gsCAT_TFML_IPN & ", " & _
                                        gsCAT_TFML_ASSYVENDORCODE & ", " & _
                                        gsCAT_TFML_DATECODE & ", " & _
                                        gsCAT_TFML_DATECODESEQ & ", " & _
                                        gsCAT_TFML_SAPOITEM & ", " & _
                                        gsCAT_TFML_SAPONO & ", " & _
                                        gsCAT_TFML_SASM_POITEM & ", " & _
                                        gsCAT_TFML_SASM_PONO & ", " & _
                                        gsCAT_TFML_FABCREATETIME & ", " & _
                                        gsCAT_TFML_REMARKCOUNT & ", " & _
                                        gsCAT_TFML_CQTY & ", " & _
                                        gsCAT_TFML_PARENTQTY & ", " & _
                                        gsCAT_TFML_SOURCE & ", " & _
                                        gsCAT_TFML_CREATEUSERID & ", " & _
                                        gsCAT_TFML_CREATETIME & ") "
    
    '取得舊資料時間
    sSQL = "select max(AA." & gsCAT_TFML_CREATETIME & ") As " & gsCAT_TFML_CREATETIME & " " & _
             "from " & gsCAT_TBL_FVI_MERGE_LIST & " AA " & _
            "where AA." & gsCAT_TFML_PARENTLOTID & " = '" & sNewParentLotID & "' "
            
    Set colResult = moProRawSql.QueryDatabase(sSQL)
            
    '取得時間, 沒有時間或資料的話會取得空字串
    If colResult.Count > 0 Then
        sOldCreateTime = colResult.Item(1).Item(gsCAT_TFML_CREATETIME)
    End If
      
    For iIdx = LBound(sChildLotIDArray) To UBound(sChildLotIDArray)
       '取得舊資料清單
        sSQL = "select a." & gsCAT_TFML_PARENTLOTID & ", " & _
                     " a." & gsCAT_TFML_CHILDLOTID & ", " & _
                     " a." & gsCAT_TFML_CREATETIME & " " & _
                 "from " & gsCAT_TBL_FVI_MERGE_LIST & " a " & _
                "where a." & gsCAT_TFML_PARENTLOTID & " = '" & sNewParentLotID & "' " & _
                  "and a." & gsCAT_TFML_CHILDLOTID & " = '" & sChildLotIDArray(iIdx) & "' " & _
                  "and a." & gsCAT_TFML_CREATETIME & " = '" & sOldCreateTime & "' "
        Set colResult = moProRawSql.QueryDatabase(sSQL)
        
        '判定是否有舊資料來決定資料輸入來源
        '利用時間只有取得一筆舊資料, 且有CreateTime也才是正常舊資料
        If colResult.Count = 1 And sOldCreateTime <> "" Then
            '以Tbl_Fvi_Merge_List舊資料更新
            sSQLInsertValue = "select " & "'" & sNewParentLotID & "', " & _
                                      "'" & sChildLotIDArray(iIdx) & "', " & _
                                      "A." & gsCAT_TFML_IPN & ", " & _
                                      "A." & gsCAT_TFML_ASSYVENDORCODE & ", " & _
                                      "A." & gsCAT_TFML_DATECODE & ", " & _
                                      "A." & gsCAT_TFML_DATECODESEQ & ", " & _
                                      "A." & gsCAT_TFML_SAPOITEM & ", " & _
                                      "A." & gsCAT_TFML_SASM_PONO & ", " & _
                                      "A." & gsCAT_TFML_SASM_POITEM & ", " & _
                                      "A." & gsCAT_TFML_SASM_PONO & ", " & _
                                      "A." & gsCAT_TFML_FABCREATETIME & ", " & _
                                      "A." & gsCAT_TFML_REMARKCOUNT & ", " & _
                                      "'" & sChildCQtyArray(iIdx) & "', " & _
                                      "'" & sMergeQty & "', " & _
                                      "'" & sSource & "', " & _
                                      "'" & sUserID & "', " & _
                                      "'" & sSysTime & "' " & _
                            "from " & gsCAT_TBL_FVI_MERGE_LIST & " A " & _
                           "where A." & gsCAT_TFML_PARENTLOTID & " = '" & sNewParentLotID & "' " & _
                             "and A." & gsCAT_TFML_CHILDLOTID & " = '" & sChildLotIDArray(iIdx) & "' " & _
                             "and A." & gsCAT_TFML_CREATETIME & " = '" & sOldCreateTime & "' "
        Else
            '無舊資料, 從Tbl_Lot_info和Tbl_Lot_Attribute更新
            sSQLInsertValue = "select " & "'" & sNewParentLotID & "', " & _
                                      "'" & sChildLotIDArray(iIdx) & "', " & _
                                      "B." & gsCAT_TLATT_IPN & ", " & _
                                      "A." & gsCAT_TLI_VENDORCODE & ", " & _
                                      "B." & gsCAT_TLATT_DATECODE & ", " & _
                                      "A." & gsCAT_TLI_DATECODE_SEQ & ", " & _
                                      "A." & gsCAT_TLI_SAPOITEM & ", " & _
                                      "A." & gsCAT_TLI_SAPONO & ", " & _
                                      "A." & gsCAT_TLI_SASM_POITEM & ", " & _
                                      "A." & gsCAT_TLI_SASM_PONO & ", " & _
                                      "A." & gsCAT_TLI_FABCREATETIME & ", " & _
                                      "A." & gsCAT_TLI_REMARKCOUNT & ", " & _
                                      "'" & sChildCQtyArray(iIdx) & "', " & _
                                      "'" & sMergeQty & "', " & _
                                      "'" & sSource & "', " & _
                                      "'" & sUserID & "', " & _
                                      "'" & sSysTime & "' " & _
                            "from " & gsCAT_TBL_LOT_INFO & " A, " & _
                                      gsCAT_TBL_LOT_ATTRIBUTE & " B " & _
                           "where A." & gsCAT_TLI_LOT_ID & " = '" & sChildLotIDArray(iIdx) & "' " & _
                             "and A." & gsCAT_TLI_LOT_ID & " = B." & gsCAT_TLATT_LOTID & "(+) "
        End If
                         
        Call moProRawSql.QueryDatabase(sSQLInsertColumn & sSQLInsertValue)
    Next iIdx
    'Add by Weilun on 20171013 for Remark 限制系統建立 <end>

    'Mark by Weilun on 20171013 for Remark 限制系統建立 <start>
'    For iIdx = LBound(sChildLotIDArray) To UBound(sChildLotIDArray)
'        If sChildLotIDArray(iIdx) Like "*M?" And Len(sChildLotIDArray(iIdx)) >= 10 Then
'            Set colResult = GetCarrierMergeList(sChildLotIDArray(iIdx), oFwWIP, oFwWF, oCwMbx, oLot, moProRawSql, oLogCtrl)
'            If Not colResult Is Nothing Then
'                bHasCarrierParent = True
'                iTotalDataCount = iTotalDataCount + colResult.Count
'            End If
'        Else
'            iTotalDataCount = iTotalDataCount + 1
'        End If
'    Next iIdx
'
'    'Step-4 : 先將 ChildLotIDList  寫入 Tbl_Merge_List.
'    'Step-4-1 ChildLotIDList中 沒有 %M_ (VB: *M?) 的 Lot.
'    '         最後一筆 InfoFlag 寫入 "Y".
'    If Not bHasCarrierParent Then
'        For iIdx = LBound(sChildLotIDArray) To UBound(sChildLotIDArray)
'            sMiddleLotID = sOriParentLotID
'            iProcessDataCount = iProcessDataCount + 1
'            If iProcessDataCount = iTotalDataCount Then
'                sInfoFlag = "Y" '最後一筆 InfoFlag 寫入 "Y".
'            End If
'            'sMergeQty gsCAT_TML_MERGEQTY
'            sSQL = "insert into " & gsCAT_TBL_MERGE_LIST & " " _
'                  & "( " & gsCAT_TML_STAGE & "," & gsCAT_TML_PARENTLOTID & "," & gsCAT_TML_CHILDLOTID & "," & _
'                           gsCAT_TML_CHIPQTY & "," & gsCAT_TML_INFOFLAG & "," & _
'                           gsCAT_TML_MERGEQTY & "," & _
'                           gsCAT_TML_CREATEUSERID & "," & gsCAT_TML_CREATETIME & "," & _
'                           gsCAT_TML_FGIPN & "," & gsCAT_TML_MIDDLELOTID & " ) "
'            sSQL = sSQL & "values('" & sStage & "','" & sNewParentLotID & "','" & sChildLotIDArray(iIdx) & "'," & _
'                                  "'" & sChildCQtyArray(iIdx) & "','" & sInfoFlag & "'," & _
'                                  "'" & sMergeQty & "'," & _
'                                  "'" & sUserID & "','" & sSysTime & "'," & _
'                                  "(select " & gsCAT_TLATT_IPN & " from " & gsCAT_TBL_LOT_ATTRIBUTE & _
'                                  "  where " & gsCAT_TLATT_LOTID & "='" & sChildLotIDArray(iIdx) & "') , " & _
'                                  "'" & sMiddleLotID & "')"
'            Call moProRawSql.QueryDatabase(sSQL)
'        Next iIdx
'    Else
'    'Step-4-2 ChildLotIDList中 有 %M_ (VB: *M?) 的 Lot.
'        For iIdx = LBound(sChildLotIDArray) To UBound(sChildLotIDArray)
'            '若 AM1 為 A,B,C組成; AM1又併D,E;
'            '則 D,E  的MiddleLotID為AM1;
'            '   A,B,C的MiddleLotID為null.
'            sMiddleLotID = sOriParentLotID
'            'Step-4-2-1 Mx LotID --> 將 %M_ (VB: *M?) 的 Lot 找出 Tbl_Merge_List 最新的一組資料, 再重新寫入.
'            '                        最後一筆 InfoFlag 寫入 "Y".
'            If sChildLotIDArray(iIdx) Like "*M?" And Len(sChildLotIDArray(iIdx)) >= 10 Then
'                sMiddleLotID = sOriParentLotID
'                Set colResult = GetCarrierMergeList(sChildLotIDArray(iIdx), oFwWIP, oFwWF, oCwMbx, oLot, moProRawSql, oLogCtrl)
'                If Not colResult Is Nothing Then
'                    For iIdx2 = 1 To colResult.Count
'
'                        sLastParentLotID = colResult.Item(iIdx2).Item(gsCAT_TML_PARENTLOTID)
'                        sLastChildLotID = colResult.Item(iIdx2).Item(gsCAT_TML_CHILDLOTID)
'                        sLsatCQty = colResult.Item(iIdx2).Item(gsCAT_TML_CHIPQTY)
'                        sMiddleLotID = sOriParentLotID
'
'                        iProcessDataCount = iProcessDataCount + 1
'                        If iProcessDataCount = iTotalDataCount Then
'                            sInfoFlag = "Y" '最後一筆 InfoFlag 寫入 "Y".
'                        End If
'                        If sOriParentLotID = sNewParentLotID Then
'                            sMiddleLotID = "" 'null
'                        End If
'
'                        sSQL = "insert into " & gsCAT_TBL_MERGE_LIST & " " _
'                              & "( " & gsCAT_TML_STAGE & "," & gsCAT_TML_PARENTLOTID & "," & gsCAT_TML_CHILDLOTID & "," & _
'                                       gsCAT_TML_CHIPQTY & "," & gsCAT_TML_INFOFLAG & "," & _
'                                       gsCAT_TML_MERGEQTY & "," & _
'                                       gsCAT_TML_CREATEUSERID & "," & gsCAT_TML_CREATETIME & "," & _
'                                       gsCAT_TML_FGIPN & "," & gsCAT_TML_MIDDLELOTID & " ) "
'                        sSQL = sSQL & "values('" & sStage & "','" & sNewParentLotID & "','" & sLastChildLotID & "'," & _
'                                              "'" & sLsatCQty & "','" & sInfoFlag & "'," & _
'                                              "'" & sMergeQty & "'," & _
'                                              "'" & sUserID & "','" & sSysTime & "'," & _
'                                              "(select " & gsCAT_TLATT_IPN & " from " & gsCAT_TBL_LOT_ATTRIBUTE & _
'                                              "  where " & gsCAT_TLATT_LOTID & "='" & sLastChildLotID & "') , " & _
'                                              "'" & sMiddleLotID & "')"
'                        Call moProRawSql.QueryDatabase(sSQL)
'                    Next iIdx2
'                End If
'            'Step-4-2-2 非 Mx LotID --> MiddleLotID = sOriParentLotID
'            '                           最後一筆 InfoFlag 寫入 "Y".
'            Else
'                sMiddleLotID = sOriParentLotID
'                iProcessDataCount = iProcessDataCount + 1
'                If iProcessDataCount = iTotalDataCount Then
'                    sInfoFlag = "Y" '最後一筆 InfoFlag 寫入 "Y".
'                End If
'                sSQL = "insert into " & gsCAT_TBL_MERGE_LIST & " " _
'                      & "( " & gsCAT_TML_STAGE & "," & gsCAT_TML_PARENTLOTID & "," & gsCAT_TML_CHILDLOTID & "," & _
'                               gsCAT_TML_CHIPQTY & "," & gsCAT_TML_INFOFLAG & "," & _
'                               gsCAT_TML_MERGEQTY & "," & _
'                               gsCAT_TML_CREATEUSERID & "," & gsCAT_TML_CREATETIME & "," & _
'                               gsCAT_TML_FGIPN & "," & gsCAT_TML_MIDDLELOTID & " ) "
'                sSQL = sSQL & "values('" & sStage & "','" & sNewParentLotID & "','" & sChildLotIDArray(iIdx) & "'," & _
'                                      "'" & sChildCQtyArray(iIdx) & "','" & sInfoFlag & "'," & _
'                                      "'" & sMergeQty & "'," & _
'                                      "'" & sUserID & "','" & sSysTime & "'," & _
'                                      "(select " & gsCAT_TLATT_IPN & " from " & gsCAT_TBL_LOT_ATTRIBUTE & _
'                                      "  where " & gsCAT_TLATT_LOTID & "='" & sChildLotIDArray(iIdx) & "') , " & _
'                                      "'" & sMiddleLotID & "')"
'                Call moProRawSql.QueryDatabase(sSQL)
'            End If
'        Next iIdx
'    End If
    'Mark by Weilun on 20171013 for Remark 限制系統建立 <end>

'----
' Done
'----
    InsertCarrierMergeList = True
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
            Case glERR_INVALIDOBJECT, glERR_FAILTOUPDATE
                typErrInfo.sUserText = typErrInfo.sErrDescription
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
            End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        InsertCarrierMergeList = False
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Function

'================================================================================
' Sub: updLotInfo_OutTime()
'--------------------------------------------------------------------------------
' Description:  for JC201300077 update Tbl_Lot_Info.OutTime欄位.
'--------------------------------------------------------------------------------
' Author:       Create by Jack on 2013/03/25
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a Object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] Jack, MXIC, 2013/03/25 for JC201300077 update Tbl_Lot_Info.OutTime欄位.
' 1) 新增 Function : updLotInfo_OutTime
'                    傳入 LotID, GroupHistKey, UpdateUser.
' 2) Used in WS.frmSortLotEnd/frmSortLotComplete; Gen.frmWsShelfDown
'================================================================================
Public Function updLotInfo_OutTime(ByVal sLotID As String, ByRef moProRawSql As Object, _
                                   ByVal sGroupHistkey, ByRef sUpdateUser, ByRef oLogCtrl As Object) As Collection
On Error GoTo ExitHandler:

Dim sProcID     As String
Dim typErrInfo  As tErrInfo
Dim colRaws     As Collection
Dim oRaws       As FwStrings
Dim sTable      As String
Dim sColumn     As String
Dim sWhere      As String

Dim sStage      As String
Dim sSQL        As String
Dim colRS       As Collection
Dim iIndex      As Integer

'----
' Init
'----
    sProcID = "updLotInfo_OutTime"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    
'----
' Action
'----
    sSQL = "update " & gsCAT_TBL_LOT_INFO _
         & " set " & " outtime = to_char(sysdate, 'YYYYMMDD HH24MISS') || '000', " _
         & gsCAT_TLI_UPDATE_USER_ID & " = '" & sUpdateUser & "', " _
         & gsCAT_TLI_GROUPHISTKEY & " = '" & sGroupHistkey & "'," _
         & gsCAT_TLI_UPDATE_TIME & " = to_char(sysdate, 'YYYYMMDD HH24MISS')||'000' " _
         & " where " & gsCAT_TLI_LOT_ID & " = '" & sLotID & "' "
    
    moProRawSql.QueryDatabase (sSQL)
    
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Function

'================================================================================
' Sub: CheckFviMergeLotID()
'--------------------------------------------------------------------------------
' Description:  for Remark 限制系統建立, 檢查Lot是否為後端併批
'--------------------------------------------------------------------------------
' Author:       Create by Weilun on 20171013
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a Object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] Weilun, MXIC, 20171013 for Remark 限制系統建立
' 1) 新增 Function : CheckFviMergeLotID
'                    傳入 LotID, moProRawSql, oLogCtrl.
'                    檢查LotID的9/10碼是否在TBL_FVI_MERGELOTNO_CONTRL的組合中
'================================================================================
Public Function CheckFviMergeLotID(ByVal sLotID As String, _
                                   ByRef moProRawSql As Object, _
                                   ByRef oLogCtrl As Object) As Boolean
On Error GoTo ExitHandler:

Dim sProcID     As String
Dim typErrInfo  As tErrInfo
Dim colRaws     As Collection
Dim sLotSeq     As String
Dim sSQL        As String


'----
' Init
'----
    sProcID = "CheckFviMergeLotID"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    
    CheckFviMergeLotID = False
    
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    
'----
' Action
'----
    If Len(sLotID) = 10 Then
        
        sLotSeq = Right(sLotID, 2)
        sSQL = "select a." & gsCAT_TFMCT_VERDERNAME & " " & _
                 "from " & gsCAT_TBL_FVI_MERGELOTNO_CONTRL & " a " & _
                "where a." & gsCAT_TFMCT_LOTNO9TH & " || a." & gsCAT_TFMCT_LOTNO10TH & " = '" & sLotSeq & "' "
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        
        If colRaws.Count > 0 Then
            CheckFviMergeLotID = True
        End If
    End If
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Function

'================================================================================
' Sub: CheckAssyVendorCodeForMerge()
'--------------------------------------------------------------------------------
' Description:  for Remark 限制系統建立, 檢查子母批AssyVendorCode是否相同,
'               IPN.Brand = 'KH' 例外
'--------------------------------------------------------------------------------
' Author:       Create by Weilun on 20171018
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a Object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] Weilun, MXIC, 20171017 for Remark 限制系統建立
' 1) 新增 Function : CheckAssyVendorCodeForMerge
'                    傳入 ParentLotID, sChildLotID, moProRawSql, oLogCtrl.
'                    檢查ipn.brand <> 'KH' 時, 子母批AssyVendorCode是否一樣
'================================================================================
Public Function CheckAssyVendorCodeForMerge(ByVal sParentLotID As String, _
                                            ByVal sChildLotID As String, _
                                            ByRef moProRawSql As Object, _
                                            ByRef oLogCtrl As Object) As Boolean
On Error GoTo ExitHandler:

Dim sProcID     As String
Dim typErrInfo  As tErrInfo
Dim colRaws     As Collection
Dim sLotSeq     As String
Dim sSQL        As String
'----
' Init
'----
    sProcID = "CheckFviMergeLotID"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    
    CheckAssyVendorCodeForMerge = False
    
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    
'----
' Action
'----
    If sParentLotID <> "" And sChildLotID <> "" Then
        
        sSQL = "select b." & gsCAT_TIM_BRAND & " " & _
         "from " & gsCAT_TBL_LOT_ATTRIBUTE & " a, " & _
                   gsCAT_TBL_IPN_MASTER & " b " & _
        "where a." & gsCAT_TLATT_LOTID & " in '" & sParentLotID & "' " & _
          "and a." & gsCAT_TLATT_IPN & " = b." & gsCAT_TIM_IPN & " (+) "
          
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        
        If colRaws.Count = 1 Then
            'KH不用檢查, 回傳True
            If colRaws.Item(1).Item(gsCAT_TIM_BRAND) <> "KH" Then
                '子母批一起查詢
                sSQL = "select a." & gsCAT_TLI_LOT_ID & ", " & _
                              "a." & gsCAT_TLI_VENDORCODE & " " & _
                         "from " & gsCAT_TBL_LOT_INFO & " a " & _
                        "where a." & gsCAT_TLI_LOT_ID & " in ('" & sParentLotID & "', '" & sChildLotID & "')"
                        
                Set colRaws = moProRawSql.QueryDatabase(sSQL)
                
                '兩筆資料進行條件檢查
                If colRaws.Count = 2 Then
                    If colRaws.Item(1).Item(gsCAT_TLI_VENDORCODE) = colRaws.Item(2).Item(gsCAT_TLI_VENDORCODE) Then
                        'IPN.BRAND <> 'KH', 子母批的IPN和AssyVendorCode一樣
                        CheckAssyVendorCodeForMerge = True
                    End If
                End If
            Else
                'KH不用檢查, 回傳True
                CheckAssyVendorCodeForMerge = True
            End If
        End If
    End If
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Function

'================================================================================
' Sub: CheckFutActWithNoHold()
' Description:  以LotId為索引, 取得ByLot,ByIpn,ByProdGroup的FutureAction資料.
'               並不包含Future Hold的資料.
'               部分Rule在FutActByLot需要EqType2作為條件.
' Author: Weilun Huang, MXIC on 2018/03/08 for BE 工業 3.5 Phase 16 - CP Auto Lot Complete
' sRuleName Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案
'================================================================================
Public Function CheckFutActWithNoHold(ByRef oLogCtrl As Object, _
                                      ByRef oProRawSQL As Object, _
                                      ByVal sLotID As String, _
                                      Optional ByVal sEqType2 As String = "", _
                                      Optional ByVal colInput As Collection, _
                                      Optional ByVal sRuleName As String = "") As Boolean
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim sSQL            As String
Dim sSubSQL         As String
Dim colRaws         As Collection
Dim colFutAct       As Collection

Dim sEng            As String
Dim sNormal         As String
Dim sRework         As String

Dim sIPN            As String
Dim sProdgroup      As String
Dim sStepNo         As String
Dim sStepName       As String   'Add by Weilun on 20180524 for for BE 工業 3.5 Phase 16 - CP Auto Lot Complete
Dim sPath           As String
Dim sLotOwner       As String
Dim sLotType        As String

Dim lCnt            As Long
Dim lEngCnt         As Long
Dim lRwCnt          As Long
Dim bFutAct         As Long

'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案<start>
Dim bActionOwner_SQL        As Boolean
Dim bActionOwner_PROD       As Boolean
Dim bActionOwner_HW         As Boolean
Dim bActionOwner_PE         As Boolean
Dim lCommCnt                As Long

Dim lProdCommCnt            As Long
Dim lHWCommCnt              As Long
Dim lPECommCnt              As Long
'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案<end>
'----
' Init
'----
    sProcID = "CheckFutActWithNoHold"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)
    CheckFutActWithNoHold = False   '預設值
    
    sEng = "Eng"
    sNormal = "Normal"
    sRework = "Rework"
    
    'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
    bActionOwner_SQL = False
    bActionOwner_PROD = False
    bActionOwner_HW = False
    bActionOwner_PE = False
    'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
        '取得Action Owner (使用oProRawSql查詢)
    Call GetActionOwner(oLogCtrl, sRuleName, oProRawSQL, Nothing, Nothing, Nothing, Nothing, bActionOwner_SQL, bActionOwner_PROD, bActionOwner_HW, bActionOwner_PE) 'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案

    '說明:
    'Hold判定主要為HoldCode以及HoldReason
    '1. 兩個都無值: Future Action, 模組回傳True (紅燈)
    '2. 其中一個無值, 另一個有值: 異常, 模組回傳True (SI希望紅燈)
    '3. 兩個都有值: Future Hold, 模組維持預設值False, 並繼續檢查下一項

    If Not (colInput Is Nothing) Then GoTo ForSpecQuery  'Add by Sam on 20180615 for 紅燈減量 ,當有另外指定資料時直接跳過Lot Base

    '**********
    'Step1:基礎資料
    '**********
    'Add StepName, by Weilun on 20180524 for for BE 工業 3.5 Phase 16 - CP Auto Lot Complete
    sSQL = " select H." & gsCAT_TLATT_IPN & ", " & _
                  " H." & gsCAT_TLATT_PRODGROUP & ", " & _
                  " H." & gsCAT_TLATT_STEPID & ", " & _
                  " H." & gsCAT_TLATT_STEPNAME & ", " & _
                  " H." & gsCAT_TLATT_ROUTE & ", " & _
                  " H." & gsCAT_TLATT_LOTOWNER & ", " & _
                  " Decode(F." & gsCAT_TLI_ERUNTICNO & ", " & _
                         " NULL, " & _
                         " Decode(F." & gsCAT_TLI_SAPRWNO & ", NULL, '" & sNormal & "', '" & sRework & "')," & _
                         " '" & sEng & "') as LotType " & _
             " from " & gsCAT_TBL_LOT_INFO & " F, " & gsCAT_TBL_LOT_ATTRIBUTE & " H " & _
            " where H." & gsCAT_TLATT_LOTID & " = F." & gsCAT_TLI_LOT_ID & " " & _
              " and H." & gsCAT_TLATT_LOTID & " = '" & sLotID & "' "
              
    Set colRaws = oProRawSQL.QueryDatabase(sSQL)
    If colRaws.Count = 1 Then
        sIPN = colRaws.Item(1).Item(gsCAT_TLATT_IPN)
        sProdgroup = colRaws.Item(1).Item(gsCAT_TLATT_PRODGROUP)
        sStepNo = colRaws.Item(1).Item(gsCAT_TLATT_STEPID) 'StepID -> StepNo
        sStepName = colRaws.Item(1).Item(gsCAT_TLATT_STEPNAME) 'Add by Weilun on 20180524 for for BE 工業 3.5 Phase 16 - CP Auto Lot Complete
        sPath = colRaws.Item(1).Item(gsCAT_TLATT_ROUTE) 'Route -> Path
        sLotOwner = colRaws.Item(1).Item(gsCAT_TLATT_LOTOWNER)
        sLotType = colRaws.Item(1).Item("LotType")
        
        '**********
        'Step2:Check FutAct By IPN
        '**********
        'Modify from DBFunction.Fun_ChkFutActByIPN,
        '除了排除Hold的狀況, 也多加了select的nvl以避免空值與數字的比較
        'prodcnt,hwcnt,pecnt,Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案
        bFutAct = False
        sSQL = " select count(a." & gsCAT_TIFA_IPN & ") as cnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TIFA_PRODCOMMENTS & ", null, 0, 1)),0) as prodcnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TIFA_HWCOMMENTS & ", null, 0, 1)),0) as hwcnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TIFA_PECOMMENTS & ", null, 0, 1)),0) as pecnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TIFA_INCLUDEENGLOT & ", 'Y', 1, 0)),0) as engcnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TIFA_INCLUDEREWORKLOT & ", 'Y', 1, 0)),0) as rwcnt " & _
                 " from " & gsCAT_TBL_IPN_FUTACT & " a " & _
                " Where a." & gsCAT_TIFA_IPN & " = '" & sIPN & "' " & _
                  " and a." & gsCAT_TIFA_STEP_NO & " = '" & sStepNo & "' " & _
                  " and a." & gsCAT_TIFA_PATH & " = '" & sPath & "' " & _
                  " and (nvl(a." & gsCAT_TIFA_INCLUDELOTOWNER & ", 'All') = 'All' or " & _
                      " a." & gsCAT_TIFA_INCLUDELOTOWNER & " = '" & sLotOwner & "' or " & _
                      " a." & gsCAT_TIFA_INCLUDELOTOWNER & " like '" & sLotOwner & "' || ',%' or " & _
                      " a." & gsCAT_TIFA_INCLUDELOTOWNER & " like '%,' || '" & sLotOwner & "' or " & _
                      " a." & gsCAT_TIFA_INCLUDELOTOWNER & " like '%,' || '" & sLotOwner & "' || ',%') " & _
                  " and (a." & gsCAT_TIFA_HOLD_CODE & " is null " & _
                    " or a." & gsCAT_TIFA_HOLD_REASON & " is null )" & _
                  " and a." & gsCAT_TIFA_DELETE_FLAG & " = 'N' "
                  
        Set colFutAct = oProRawSQL.QueryDatabase(sSQL)
        If colFutAct.Count = 1 Then '使用SQL的運算, 所以固定一筆
            lCnt = colFutAct.Item(1).Item("cnt")
            lEngCnt = colFutAct.Item(1).Item("engcnt")
            lRwCnt = colFutAct.Item(1).Item("rwcnt")
            
            If lCnt = 0 Then
                bFutAct = False
            ElseIf sLotType = sNormal Then
                bFutAct = True
            ElseIf sLotType = sEng And lEngCnt > 0 Then
                bFutAct = True
            ElseIf sLotType = sRework And lRwCnt > 0 Then
                bFutAct = True
            Else
                bFutAct = False
            End If
            
            If bFutAct = True Then
                'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                lCommCnt = 0
'                If bActionOwner_SQL = False Then '沒指定ActionOwner舊有規則 (暫留)'20211005確定拔掉
'                    CheckFutActWithNoHold = True
'                    GoTo ExitHandler
'                End If
            
                '大於1表示有指定的ActionOwner的內容
                '相容性: 非指定Rule時, 除了bActionOwner_SQL都為True, 新規則有設定ActionOwner將吃的到
                If bActionOwner_PROD = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("prodcnt")
                End If
                If bActionOwner_HW = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("hwcnt")
                End If
                If bActionOwner_PE = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("pecnt")
                End If
                If lCommCnt > 0 Then
                    CheckFutActWithNoHold = True
                    GoTo ExitHandler
                End If
                'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>

                'Mark by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                'CheckFutActWithNoHold = True
                'GoTo ExitHandler
                'Mark by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案  <end>
            End If
        'Mark by Sam on 20180522,無資料不需要產生error exception
'        Else
'            Call RaiseError(glERR_FAILTOQUERY, FormatErrorText(gsETX_FAILTOQUERY, sLotId))
        End If
        'End of check FutAct By IPN
        
        '**********
        'Step3:Check FutAct By ProdGroup
        '**********
        'Modify from DBFunction.Fun_ChkFutActByProdGroup,
        '除了排除Hold的狀況, 也多加了select的nvl以避免空值與數字的比較
        'prodcnt,hwcnt,pecnt,Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案
        bFutAct = False
        sSQL = " select count(a." & gsCAT_TPGFA_PROD_GROUP & ") as cnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TPGFA_PRODCOMMENTS & ", null, 0, 1)),0) as prodcnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TPGFA_HWCOMMENTS & ", null, 0, 1)),0) as hwcnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TPGFA_PECOMMENTS & ", null, 0, 1)),0) as pecnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TPGFA_INCLUDEENGLOT & ", 'Y', 1, 0)),0) as engcnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TPGFA_INCLUDEREWORKLOT & ", 'Y', 1, 0)),0) as rwcnt " & _
                 " from " & gsCAT_TBL_PRODGROUP_FUTACT & " a " & _
                " Where a." & gsCAT_TPGFA_PROD_GROUP & " = '" & sProdgroup & "' " & _
                  " and a." & gsCAT_TPGFA_STEP_NO & " = '" & sStepNo & "' " & _
                  " and a." & gsCAT_TPGFA_PATH & " = '" & sPath & "' " & _
                  " and (nvl(a." & gsCAT_TIFA_INCLUDELOTOWNER & ", 'All') = 'All' or " & _
                      " a." & gsCAT_TPGFA_INCLUDELOTOWNER & " = '" & sLotOwner & "' or " & _
                      " a." & gsCAT_TPGFA_INCLUDELOTOWNER & " like '" & sLotOwner & "' || ',%' or " & _
                      " a." & gsCAT_TPGFA_INCLUDELOTOWNER & " like '%,' || '" & sLotOwner & "' or " & _
                      " a." & gsCAT_TPGFA_INCLUDELOTOWNER & " like '%,' || '" & sLotOwner & "' || ',%') " & _
                  " and (a." & gsCAT_TPGFA_HOLD_CODE & " is null " & _
                    " or a." & gsCAT_TPGFA_HOLD_REASON & " is null )" & _
                  " and a." & gsCAT_TPGFA_DELETE_FLAG & " = 'N' "
                  
        Set colFutAct = oProRawSQL.QueryDatabase(sSQL)
        If colFutAct.Count = 1 Then '使用SQL的運算, 所以固定一筆
            lCnt = colFutAct.Item(1).Item("cnt")
            lEngCnt = colFutAct.Item(1).Item("engcnt")
            lRwCnt = colFutAct.Item(1).Item("rwcnt")
            
            If lCnt = 0 Then
                bFutAct = False
            ElseIf sLotType = sNormal Then
                bFutAct = True
            ElseIf sLotType = sEng And lEngCnt > 0 Then
                bFutAct = True
            ElseIf sLotType = sRework And lRwCnt > 0 Then
                bFutAct = True
            Else
                bFutAct = False
            End If
            
            If bFutAct = True Then
                'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                lCommCnt = 0
'                If bActionOwner_SQL = False Then '沒指定ActionOwner舊有規則 (暫留)'20211005確定拔掉
'                    CheckFutActWithNoHold = True
'                    GoTo ExitHandler
'                End If
            
                '大於1表示有指定的ActionOwner的內容
                '相容性: 非指定Rule時, 除了bActionOwner_SQL都為True, 新規則有設定ActionOwner將吃的到
                If bActionOwner_PROD = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("prodcnt")
                End If
                If bActionOwner_HW = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("hwcnt")
                End If
                If bActionOwner_PE = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("pecnt")
                End If
                If lCommCnt > 0 Then
                    CheckFutActWithNoHold = True
                    GoTo ExitHandler
                End If
                'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>

                'Mark by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                'CheckFutActWithNoHold = True
                'GoTo ExitHandler
                'Mark by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案  <end>
            End If
        'Mark by Sam on 20180522,無資料不需要產生error exception
'        Else
'            Call RaiseError(glERR_FAILTOQUERY, FormatErrorText(gsETX_FAILTOQUERY, sLotId))
        End If
        'End of check FutAct By IPN
        
        '**********
        'Step4:Check FutAct By Lot
        '**********
        'prodcnt,hwcnt,pecnt,Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案
        sSubSQL = " select " & gsCAT_TLFA_LOT_ID & ", " & _
                               gsCAT_TLFA_STEP_NO & ", " & _
                               gsCAT_TLFA_PATH & ", " & _
                             " nvl(sum(decode(" & gsCAT_TLFA_PRODCOMMENTS & ", null, 0, 1)),0) as prodcnt, " & _
                             " nvl(sum(decode(" & gsCAT_TLFA_HWCOMMENTS & ", null, 0, 1)),0) as hwcnt, " & _
                             " nvl(sum(decode(" & gsCAT_TLFA_PECOMMENTS & ", null, 0, 1)),0) as pecnt " & _
                    " from " & gsCAT_TBL_LOT_FUTACT & " " & _
                   " where (" & gsCAT_TLFA_HOLD_CODE & " is null " & _
                       " or " & gsCAT_TLFA_HOLD_REASON & " is null )" & _
                     " and " & gsCAT_TLFA_DELETE_FLAG & " = 'N' "
        
        '加額外條件eqtype2
        If sEqType2 <> "" Then
            sSubSQL = sSubSQL & _
                         " and nvl(" & gsCAT_TLFA_EQTYPE2 & ", '" & sEqType2 & "') = '" & sEqType2 & "' "
        End If
        
        sSubSQL = sSubSQL & _
                  " group by " & gsCAT_TLFA_LOT_ID & ", " & _
                                 gsCAT_TLFA_STEP_NO & ", " & _
                                 gsCAT_TLFA_PATH & " "
        
        'prodcnt,hwcnt,pecnt,Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案
        sSQL = " select decode(N." & gsCAT_TLFA_LOT_ID & ", NULL, 'N', 'Y') LotFutAct, " & _
                      " nvl(N.prodcnt, 0) as prodcnt, " & _
                      " nvl(N.hwcnt, 0) as hwcnt, " & _
                      " nvl(N.pecnt, 0) As pecnt " & _
                 " from (" & sSubSQL & ")N, " & _
                        gsCAT_TBL_LOT_ATTRIBUTE & " tla " & _
                " where tla." & gsCAT_TLATT_LOTID & " = N." & gsCAT_TLFA_LOT_ID & " (+) " & _
                  " and tla." & gsCAT_TLATT_STEPID & " = N." & gsCAT_TLFA_STEP_NO & " (+) " & _
                  " and tla." & gsCAT_TLATT_ROUTE & " = N." & gsCAT_TLFA_PATH & " (+) " & _
                  " and tla." & gsCAT_TLATT_LOTID & " = '" & sLotID & "' "
        
        Set colFutAct = oProRawSQL.QueryDatabase(sSQL)
        If colFutAct.Count > 0 Then '基本上一筆
            If colFutAct.Item(1).Item("LotFutAct") = "Y" Then
                'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                lCommCnt = 0
'                If bActionOwner_SQL = False Then '沒指定ActionOwner舊有規則 (暫留)'20211005確定拔掉
'                    CheckFutActWithNoHold = True
'                    GoTo ExitHandler
'                End If
            
                '大於1表示有指定的ActionOwner的內容
                '相容性: 非指定Rule時, 除了bActionOwner_SQL都為True, 新規則有設定ActionOwner將吃的到
                If bActionOwner_PROD = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("prodcnt")
                End If
                If bActionOwner_HW = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("hwcnt")
                End If
                If bActionOwner_PE = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("pecnt")
                End If
                If lCommCnt > 0 Then
                    CheckFutActWithNoHold = True
                    GoTo ExitHandler
                End If
                'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>
                
                'Mark by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
'                CheckFutActWithNoHold = True
'                GoTo ExitHandler
                'Mark by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>
            End If
        'Mark by Sam on 20180522,無資料不需要產生error exception
'        Else
'            Call RaiseError(glERR_FAILTOQUERY, FormatErrorText(gsETX_FAILTOQUERY, sLotId))
        End If
        
        'Add by Weilun on 20180524 for for BE 工業 3.5 Phase 16 - CP Auto Lot Complete<Start>
        '**********
        'Step5:Check FutAct By IPN, StepName
        '**********
        'Modify from DBFunction.Fun_ChkFutActByIPN,
        '除了排除Hold的狀況, 也多加了select的nvl以避免空值與數字的比較
        'prodcnt,hwcnt,pecnt,Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案
        bFutAct = False
        sSQL = " select count(a." & gsCAT_TIFA_IPN & ") as cnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TIFA_PRODCOMMENTS & ", null, 0, 1)),0) as prodcnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TIFA_HWCOMMENTS & ", null, 0, 1)),0) as hwcnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TIFA_PECOMMENTS & ", null, 0, 1)),0) as pecnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TIFA_INCLUDEENGLOT & ", 'Y', 1, 0)),0) as engcnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TIFA_INCLUDEREWORKLOT & ", 'Y', 1, 0)),0) as rwcnt " & _
                 " from " & gsCAT_TBL_IPN_FUTACT & " a " & _
                " Where a." & gsCAT_TIFA_IPN & " = '" & sIPN & "' " & _
                  " and a." & gsCAT_TIFA_STEPNAME & " = '" & sStepName & "' " & _
                  " and (nvl(a." & gsCAT_TIFA_INCLUDELOTOWNER & ", 'All') = 'All' or " & _
                      " a." & gsCAT_TIFA_INCLUDELOTOWNER & " = '" & sLotOwner & "' or " & _
                      " a." & gsCAT_TIFA_INCLUDELOTOWNER & " like '" & sLotOwner & "' || ',%' or " & _
                      " a." & gsCAT_TIFA_INCLUDELOTOWNER & " like '%,' || '" & sLotOwner & "' or " & _
                      " a." & gsCAT_TIFA_INCLUDELOTOWNER & " like '%,' || '" & sLotOwner & "' || ',%') " & _
                  " and (a." & gsCAT_TIFA_HOLD_CODE & " is null " & _
                    " or a." & gsCAT_TIFA_HOLD_REASON & " is null )" & _
                  " and a." & gsCAT_TIFA_DELETE_FLAG & " = 'N' "
                  
        Set colFutAct = oProRawSQL.QueryDatabase(sSQL)
        If colFutAct.Count = 1 Then '使用SQL的運算, 所以固定一筆
            lCnt = colFutAct.Item(1).Item("cnt")
            lEngCnt = colFutAct.Item(1).Item("engcnt")
            lRwCnt = colFutAct.Item(1).Item("rwcnt")
            
            If lCnt = 0 Then
                bFutAct = False
            ElseIf sLotType = sNormal Then
                bFutAct = True
            ElseIf sLotType = sEng And lEngCnt > 0 Then
                bFutAct = True
            ElseIf sLotType = sRework And lRwCnt > 0 Then
                bFutAct = True
            Else
                bFutAct = False
            End If
            
            If bFutAct = True Then
                'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                lCommCnt = 0
'                If bActionOwner_SQL = False Then '沒指定ActionOwner舊有規則 (暫留)'20211005確定拔掉
'                    CheckFutActWithNoHold = True
'                    GoTo ExitHandler
'                End If
            
                '大於1表示有指定的ActionOwner的內容
                '相容性: 非指定Rule時, 除了bActionOwner_SQL都為True, 新規則有設定ActionOwner將吃的到
                If bActionOwner_PROD = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("prodcnt")
                End If
                If bActionOwner_HW = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("hwcnt")
                End If
                If bActionOwner_PE = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("pecnt")
                End If
                If lCommCnt > 0 Then
                    CheckFutActWithNoHold = True
                    GoTo ExitHandler
                End If
                'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>

                'Mark by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                'CheckFutActWithNoHold = True
                'GoTo ExitHandler
                'Mark by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案  <end>
            End If
        End If
        'End of check FutAct By IPN
        
        '**********
        'Step6:Check FutAct By ProdGroup, StepName
        '**********
        'Modify from DBFunction.Fun_ChkFutActByProdGroup,
        '除了排除Hold的狀況, 也多加了select的nvl以避免空值與數字的比較
        'prodcnt,hwcnt,pecnt,Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案
        bFutAct = False
        sSQL = " select count(a." & gsCAT_TPGFA_PROD_GROUP & ") as cnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TPGFA_PRODCOMMENTS & ", null, 0, 1)),0) as prodcnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TPGFA_HWCOMMENTS & ", null, 0, 1)),0) as hwcnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TPGFA_PECOMMENTS & ", null, 0, 1)),0) as pecnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TPGFA_INCLUDEENGLOT & ", 'Y', 1, 0)),0) as engcnt, " & _
                      " nvl(sum(decode(a." & gsCAT_TPGFA_INCLUDEREWORKLOT & ", 'Y', 1, 0)),0) as rwcnt " & _
                 " from " & gsCAT_TBL_PRODGROUP_FUTACT & " a " & _
                " Where a." & gsCAT_TPGFA_PROD_GROUP & " = '" & sProdgroup & "' " & _
                  " and a." & gsCAT_TPGFA_STEPNAME & " = '" & sStepName & "' " & _
                  " and (nvl(a." & gsCAT_TIFA_INCLUDELOTOWNER & ", 'All') = 'All' or " & _
                      " a." & gsCAT_TPGFA_INCLUDELOTOWNER & " = '" & sLotOwner & "' or " & _
                      " a." & gsCAT_TPGFA_INCLUDELOTOWNER & " like '" & sLotOwner & "' || ',%' or " & _
                      " a." & gsCAT_TPGFA_INCLUDELOTOWNER & " like '%,' || '" & sLotOwner & "' or " & _
                      " a." & gsCAT_TPGFA_INCLUDELOTOWNER & " like '%,' || '" & sLotOwner & "' || ',%') " & _
                  " and (a." & gsCAT_TPGFA_HOLD_CODE & " is null " & _
                    " or a." & gsCAT_TPGFA_HOLD_REASON & " is null )" & _
                  " and a." & gsCAT_TPGFA_DELETE_FLAG & " = 'N' "
                  
        Set colFutAct = oProRawSQL.QueryDatabase(sSQL)
        If colFutAct.Count = 1 Then '使用SQL的運算, 所以固定一筆
            lCnt = colFutAct.Item(1).Item("cnt")
            lEngCnt = colFutAct.Item(1).Item("engcnt")
            lRwCnt = colFutAct.Item(1).Item("rwcnt")
            
            If lCnt = 0 Then
                bFutAct = False
            ElseIf sLotType = sNormal Then
                bFutAct = True
            ElseIf sLotType = sEng And lEngCnt > 0 Then
                bFutAct = True
            ElseIf sLotType = sRework And lRwCnt > 0 Then
                bFutAct = True
            Else
                bFutAct = False
            End If
            
            If bFutAct = True Then
                'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                lCommCnt = 0
'                If bActionOwner_SQL = False Then '沒指定ActionOwner舊有規則 (暫留)'20211005確定拔掉
'                    CheckFutActWithNoHold = True
'                    GoTo ExitHandler
'                End If
            
                '大於1表示有指定的ActionOwner的內容
                '相容性: 非指定Rule時, 除了bActionOwner_SQL都為True, 新規則有設定ActionOwner將吃的到
                If bActionOwner_PROD = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("prodcnt")
                End If
                If bActionOwner_HW = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("hwcnt")
                End If
                If bActionOwner_PE = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("pecnt")
                End If
                If lCommCnt > 0 Then
                    CheckFutActWithNoHold = True
                    GoTo ExitHandler
                End If
                'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>

                'Mark by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                'CheckFutActWithNoHold = True
                'GoTo ExitHandler
                'Mark by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案  <end>
            End If
        End If
        'End of check FutAct By IPN
        
        '**********
        'Step7:Check FutAct By Lot, StepName
        '**********
        sSubSQL = " select " & gsCAT_TLFA_LOT_ID & ", " & _
                               gsCAT_TLFA_STEPNAME & ", " & _
                             " nvl(sum(decode(" & gsCAT_TLFA_PRODCOMMENTS & ", null, 0, 1)),0) as prodcnt, " & _
                             " nvl(sum(decode(" & gsCAT_TLFA_HWCOMMENTS & ", null, 0, 1)),0) as hwcnt, " & _
                             " nvl(sum(decode(" & gsCAT_TLFA_PECOMMENTS & ", null, 0, 1)),0) as pecnt " & _
                    " from " & gsCAT_TBL_LOT_FUTACT & " " & _
                   " where (" & gsCAT_TLFA_HOLD_CODE & " is null " & _
                       " or " & gsCAT_TLFA_HOLD_REASON & " is null )" & _
                     " and " & gsCAT_TLFA_DELETE_FLAG & " = 'N' "
        
        '加額外條件eqtype2
        If sEqType2 <> "" Then
            sSubSQL = sSubSQL & _
                         " and nvl(" & gsCAT_TLFA_EQTYPE2 & ", '" & sEqType2 & "') = '" & sEqType2 & "' "
        End If
        
        sSubSQL = sSubSQL & _
                  " group by " & gsCAT_TLFA_LOT_ID & ", " & _
                                 gsCAT_TLFA_STEPNAME & " "

        sSQL = " select decode(N." & gsCAT_TLFA_LOT_ID & ", NULL, 'N', 'Y') LotFutAct, " & _
                      " nvl(N.prodcnt, 0) as prodcnt, " & _
                      " nvl(N.hwcnt, 0) as hwcnt, " & _
                      " nvl(N.pecnt, 0) As pecnt " & _
                 " from (" & sSubSQL & ")N, " & _
                        gsCAT_TBL_LOT_ATTRIBUTE & " tla " & _
                " where tla." & gsCAT_TLATT_LOTID & " = N." & gsCAT_TLFA_LOT_ID & " (+) " & _
                  " and tla." & gsCAT_TLATT_STEPNAME & " = N." & gsCAT_TLFA_STEPNAME & " (+) " & _
                  " and tla." & gsCAT_TLATT_LOTID & " = '" & sLotID & "' "
        
        Set colFutAct = oProRawSQL.QueryDatabase(sSQL)
        If colFutAct.Count > 0 Then '基本上一筆
            If colFutAct.Item(1).Item("LotFutAct") = "Y" Then
                'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                lCommCnt = 0
'                If bActionOwner_SQL = False Then '沒指定ActionOwner舊有規則 (暫留)'20211005確定拔掉
'                    CheckFutActWithNoHold = True
'                    GoTo ExitHandler
'                End If
            
                '大於1表示有指定的ActionOwner的內容
                '相容性: 非指定Rule時, 除了bActionOwner_SQL都為True, 新規則有設定ActionOwner將吃的到
                If bActionOwner_PROD = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("prodcnt")
                End If
                If bActionOwner_HW = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("hwcnt")
                End If
                If bActionOwner_PE = True Then
                    lCommCnt = lCommCnt + colFutAct.Item(1).Item("pecnt")
                End If
                If lCommCnt > 0 Then
                    CheckFutActWithNoHold = True
                    GoTo ExitHandler
                End If
                'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>

                'Mark by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                'CheckFutActWithNoHold = True
                'GoTo ExitHandler
                'Mark by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案  <end>
            End If
        End If
        'End of check FutAct By Lot
        'Add by Weilun on 20180524 for for BE 工業 3.5 Phase 16 - CP Auto Lot Complete<End>
        
     'Mark by Sam on 20180522,無資料不需要產生error exception
'    Else
'        Call RaiseError(glERR_FAILTOQUERY, FormatErrorText(gsETX_FAILTOQUERY, sLotId))
    End If
    
ForSpecQuery:
    If Not colInput Is Nothing Then
        If colInput.Count >= 5 Then  '無Lot僅看IPN及Prodgroup FutAct,5個參數需湊足
            lCnt = 0
            lProdCommCnt = 0
            lHWCommCnt = 0
            lPECommCnt = 0
        
            bFutAct = False
            sIPN = colInput("ipn")
            sProdgroup = colInput("prodgroup")
            sStepNo = colInput("stepno")
            sStepName = colInput("stepname")
            sPath = colInput("path")
            
            'Lot
'            sSQL = " select a." & gsCAT_TLFA_LOT_ID & " from " & gsCAT_TBL_LOT_FUTACT & " a " & _
'                    " Where a." & gsCAT_TLFA_LOT_ID & " = '" & sLotID & "' " & _
'                      " and a." & gsCAT_TLFA_STEP_NO & " = '" & sStepNo & "' " & _
'                      " and a." & gsCAT_TLFA_PATH & " = '" & sPath & "' " & _
'                      " and (a." & gsCAT_TLFA_HOLD_CODE & " is null " & _
'                        " or a." & gsCAT_TLFA_HOLD_REASON & " is null )" & _
'                      " and a." & gsCAT_TLFA_DELETE_FLAG & " = 'N' "
            sSQL = " select count(a." & gsCAT_TLFA_LOT_ID & ") as cnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TLFA_PRODCOMMENTS & ", null, 0, 1)),0) as prodcnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TLFA_HWCOMMENTS & ", null, 0, 1)),0) as hwcnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TLFA_PECOMMENTS & ", null, 0, 1)),0) as pecnt " & _
                     " from " & gsCAT_TBL_LOT_FUTACT & " a " & _
                    " Where a." & gsCAT_TLFA_LOT_ID & " = '" & sLotID & "' " & _
                      " and a." & gsCAT_TLFA_STEP_NO & " = '" & sStepNo & "' " & _
                      " and a." & gsCAT_TLFA_PATH & " = '" & sPath & "' " & _
                      " and (a." & gsCAT_TLFA_HOLD_CODE & " is null " & _
                        " or a." & gsCAT_TLFA_HOLD_REASON & " is null )" & _
                      " and a." & gsCAT_TLFA_DELETE_FLAG & " = 'N' "
            Set colFutAct = oProRawSQL.QueryDatabase(sSQL)
            If colFutAct.Count > 0 Then
'                bFutAct = True 'Mark by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案

                'Add by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                lCnt = lCnt + colFutAct.Item(1).Item("cnt")
                lProdCommCnt = lProdCommCnt + colFutAct.Item(1).Item("prodcnt")
                lHWCommCnt = lHWCommCnt + colFutAct.Item(1).Item("hwcnt")
                lPECommCnt = lPECommCnt + colFutAct.Item(1).Item("pecnt")
                'Add by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>
            End If
            
            'Lot byStepName
'            sSQL = " select a." & gsCAT_TLFA_LOT_ID & " from " & gsCAT_TBL_LOT_FUTACT & " a " & _
'                    " Where a." & gsCAT_TLFA_LOT_ID & " = '" & sLotID & "' " & _
'                      " and a." & gsCAT_TLFA_STEPNAME & " = '" & sStepName & "' " & _
'                      " and (a." & gsCAT_TLFA_HOLD_CODE & " is null " & _
'                        " or a." & gsCAT_TLFA_HOLD_REASON & " is null )" & _
'                      " and a." & gsCAT_TLFA_DELETE_FLAG & " = 'N' "
            sSQL = " select count(a." & gsCAT_TLFA_LOT_ID & ") as cnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TLFA_PRODCOMMENTS & ", null, 0, 1)),0) as prodcnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TLFA_HWCOMMENTS & ", null, 0, 1)),0) as hwcnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TLFA_PECOMMENTS & ", null, 0, 1)),0) as pecnt " & _
                     " from " & gsCAT_TBL_LOT_FUTACT & " a " & _
                    " Where a." & gsCAT_TLFA_LOT_ID & " = '" & sLotID & "' " & _
                      " and a." & gsCAT_TLFA_STEPNAME & " = '" & sStepName & "' " & _
                      " and (a." & gsCAT_TLFA_HOLD_CODE & " is null " & _
                        " or a." & gsCAT_TLFA_HOLD_REASON & " is null )" & _
                      " and a." & gsCAT_TLFA_DELETE_FLAG & " = 'N' "
            Set colFutAct = oProRawSQL.QueryDatabase(sSQL)
            If colFutAct.Count > 0 Then
'                bFutAct = True 'Mark by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案

                'Add by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                lCnt = lCnt + colFutAct.Item(1).Item("cnt")
                lProdCommCnt = lProdCommCnt + colFutAct.Item(1).Item("prodcnt")
                lHWCommCnt = lHWCommCnt + colFutAct.Item(1).Item("hwcnt")
                lPECommCnt = lPECommCnt + colFutAct.Item(1).Item("pecnt")
                'Add by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>
            End If
            
            'IPN
'            sSQL = " select a." & gsCAT_TIFA_IPN & " from " & gsCAT_TBL_IPN_FUTACT & " a " & _
'                    " Where a." & gsCAT_TIFA_IPN & " = '" & sIPN & "' " & _
'                      " and a." & gsCAT_TIFA_STEP_NO & " = '" & sStepNo & "' " & _
'                      " and a." & gsCAT_TIFA_PATH & " = '" & sPath & "' " & _
'                      " and (a." & gsCAT_TIFA_HOLD_CODE & " is null " & _
'                        " or a." & gsCAT_TIFA_HOLD_REASON & " is null )" & _
'                      " and a." & gsCAT_TIFA_DELETE_FLAG & " = 'N' "
            sSQL = " select count(a." & gsCAT_TIFA_IPN & ") as cnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TIFA_PRODCOMMENTS & ", null, 0, 1)),0) as prodcnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TIFA_HWCOMMENTS & ", null, 0, 1)),0) as hwcnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TIFA_PECOMMENTS & ", null, 0, 1)),0) as pecnt " & _
                     " from " & gsCAT_TBL_IPN_FUTACT & " a " & _
                    " Where a." & gsCAT_TIFA_IPN & " = '" & sIPN & "' " & _
                      " and a." & gsCAT_TIFA_STEP_NO & " = '" & sStepNo & "' " & _
                      " and a." & gsCAT_TIFA_PATH & " = '" & sPath & "' " & _
                      " and (a." & gsCAT_TIFA_HOLD_CODE & " is null " & _
                        " or a." & gsCAT_TIFA_HOLD_REASON & " is null )" & _
                      " and a." & gsCAT_TIFA_DELETE_FLAG & " = 'N' "
                      
            Set colFutAct = oProRawSQL.QueryDatabase(sSQL)
            If colFutAct.Count > 0 Then
'                bFutAct = True 'Mark by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案

                'Add by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                lCnt = lCnt + colFutAct.Item(1).Item("cnt")
                lProdCommCnt = lProdCommCnt + colFutAct.Item(1).Item("prodcnt")
                lHWCommCnt = lHWCommCnt + colFutAct.Item(1).Item("hwcnt")
                lPECommCnt = lPECommCnt + colFutAct.Item(1).Item("pecnt")
                'Add by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>
            End If
            
            'IPN STEPNAME
'            sSQL = " select a." & gsCAT_TIFA_IPN & " from " & gsCAT_TBL_IPN_FUTACT & " a " & _
'                    " Where a." & gsCAT_TIFA_IPN & " = '" & sIPN & "' " & _
'                      " and a." & gsCAT_TIFA_STEPNAME & " = '" & sStepName & "' " & _
'                      " and (a." & gsCAT_TIFA_HOLD_CODE & " is null " & _
'                        " or a." & gsCAT_TIFA_HOLD_REASON & " is null )" & _
'                      " and a." & gsCAT_TIFA_DELETE_FLAG & " = 'N' "
            sSQL = " select count(a." & gsCAT_TIFA_IPN & ") as cnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TIFA_PRODCOMMENTS & ", null, 0, 1)),0) as prodcnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TIFA_HWCOMMENTS & ", null, 0, 1)),0) as hwcnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TIFA_PECOMMENTS & ", null, 0, 1)),0) as pecnt " & _
                     " from " & gsCAT_TBL_IPN_FUTACT & " a " & _
                    " Where a." & gsCAT_TIFA_IPN & " = '" & sIPN & "' " & _
                      " and a." & gsCAT_TIFA_STEPNAME & " = '" & sStepName & "' " & _
                      " and (a." & gsCAT_TIFA_HOLD_CODE & " is null " & _
                        " or a." & gsCAT_TIFA_HOLD_REASON & " is null )" & _
                      " and a." & gsCAT_TIFA_DELETE_FLAG & " = 'N' "
                      
            Set colFutAct = oProRawSQL.QueryDatabase(sSQL)
            If colFutAct.Count > 0 Then
'                bFutAct = True 'Mark by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案

                'Add by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                lCnt = lCnt + colFutAct.Item(1).Item("cnt")
                lProdCommCnt = lProdCommCnt + colFutAct.Item(1).Item("prodcnt")
                lHWCommCnt = lHWCommCnt + colFutAct.Item(1).Item("hwcnt")
                lPECommCnt = lPECommCnt + colFutAct.Item(1).Item("pecnt")
                'Add by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>
            End If
            
'            count(a." & gsCAT_TPGFA_PROD_GROUP & ") as cnt, " & _
'                      " nvl(sum(decode(a." & gsCAT_TPGFA_PRODCOMMENTS & ", null, 0, 1)),0) as prodcnt, " & _
'                      " nvl(sum(decode(a." & gsCAT_TPGFA_HWCOMMENTS & ", null, 0, 1)),0) as hwcnt, " & _
'                      " nvl(sum(decode(a." & gsCAT_TPGFA_PECOMMENTS & ", null, 0, 1)),0) as pecnt, " & _

            
            'Prodgroup
'            sSQL = " select a." & gsCAT_TPGFA_PROD_GROUP & " from " & gsCAT_TBL_PRODGROUP_FUTACT & " a " & _
'                    " Where a." & gsCAT_TPGFA_PROD_GROUP & " = '" & sProdgroup & "' " & _
'                      " and a." & gsCAT_TPGFA_STEP_NO & " = '" & sStepNo & "' " & _
'                      " and a." & gsCAT_TPGFA_PATH & " = '" & sPath & "' " & _
'                      " and (a." & gsCAT_TPGFA_HOLD_CODE & " is null " & _
'                        " or a." & gsCAT_TPGFA_HOLD_REASON & " is null )" & _
'                      " and a." & gsCAT_TPGFA_DELETE_FLAG & " = 'N' "
            sSQL = " select count(a." & gsCAT_TPGFA_PROD_GROUP & ") as cnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TPGFA_PRODCOMMENTS & ", null, 0, 1)),0) as prodcnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TPGFA_HWCOMMENTS & ", null, 0, 1)),0) as hwcnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TPGFA_PECOMMENTS & ", null, 0, 1)),0) as pecnt " & _
                     " from " & gsCAT_TBL_PRODGROUP_FUTACT & " a " & _
                    " Where a." & gsCAT_TPGFA_PROD_GROUP & " = '" & sProdgroup & "' " & _
                      " and a." & gsCAT_TPGFA_STEP_NO & " = '" & sStepNo & "' " & _
                      " and a." & gsCAT_TPGFA_PATH & " = '" & sPath & "' " & _
                      " and (a." & gsCAT_TPGFA_HOLD_CODE & " is null " & _
                        " or a." & gsCAT_TPGFA_HOLD_REASON & " is null )" & _
                      " and a." & gsCAT_TPGFA_DELETE_FLAG & " = 'N' "
                      
            Set colFutAct = oProRawSQL.QueryDatabase(sSQL)
            If colFutAct.Count > 0 Then
'                bFutAct = True 'Mark by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案

                'Add by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                lCnt = lCnt + colFutAct.Item(1).Item("cnt")
                lProdCommCnt = lProdCommCnt + colFutAct.Item(1).Item("prodcnt")
                lHWCommCnt = lHWCommCnt + colFutAct.Item(1).Item("hwcnt")
                lPECommCnt = lPECommCnt + colFutAct.Item(1).Item("pecnt")
                'Add by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>
            End If
            
            'Prodgorup Stepname
'            sSQL = " select a." & gsCAT_TPGFA_PROD_GROUP & " from " & gsCAT_TBL_PRODGROUP_FUTACT & " a " & _
'                    " Where a." & gsCAT_TPGFA_PROD_GROUP & " = '" & sProdgroup & "' " & _
'                      " and a." & gsCAT_TPGFA_STEPNAME & " = '" & sStepName & "' " & _
'                      " and (a." & gsCAT_TPGFA_HOLD_CODE & " is null " & _
'                        " or a." & gsCAT_TPGFA_HOLD_REASON & " is null )" & _
'                      " and a." & gsCAT_TPGFA_DELETE_FLAG & " = 'N' "
            sSQL = " select count(a." & gsCAT_TPGFA_PROD_GROUP & ") as cnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TPGFA_PRODCOMMENTS & ", null, 0, 1)),0) as prodcnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TPGFA_HWCOMMENTS & ", null, 0, 1)),0) as hwcnt, " & _
                          " nvl(sum(decode(a." & gsCAT_TPGFA_PECOMMENTS & ", null, 0, 1)),0) as pecnt " & _
                     " from " & gsCAT_TBL_PRODGROUP_FUTACT & " a " & _
                    " Where a." & gsCAT_TPGFA_PROD_GROUP & " = '" & sProdgroup & "' " & _
                      " and a." & gsCAT_TPGFA_STEPNAME & " = '" & sStepName & "' " & _
                      " and (a." & gsCAT_TPGFA_HOLD_CODE & " is null " & _
                        " or a." & gsCAT_TPGFA_HOLD_REASON & " is null )" & _
                      " and a." & gsCAT_TPGFA_DELETE_FLAG & " = 'N' "
                      
            Set colFutAct = oProRawSQL.QueryDatabase(sSQL)
            If colFutAct.Count > 0 Then
'                bFutAct = True 'Mark by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案

                'Add by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
                lCnt = lCnt + colFutAct.Item(1).Item("cnt")
                lProdCommCnt = lProdCommCnt + colFutAct.Item(1).Item("prodcnt")
                lHWCommCnt = lHWCommCnt + colFutAct.Item(1).Item("hwcnt")
                lPECommCnt = lPECommCnt + colFutAct.Item(1).Item("pecnt")
                'Add by Weilun on 20210908 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>
            End If
            
            'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <start>
            If lCnt > 0 Then '相對於原本的bFutAct = True
                lCommCnt = 0
'                If bActionOwner_SQL = False Then '沒指定ActionOwner舊有規則 (暫留)'20211005確定拔掉
'                    CheckFutActWithNoHold = True
'                    GoTo ExitHandler
'                End If
            
                '大於1表示有指定的ActionOwner的內容
                '相容性: 非指定Rule時, 除了bActionOwner_SQL都為True, 新規則有設定ActionOwner將吃的到
                If bActionOwner_PROD = True Then
                    lCommCnt = lCommCnt + lProdCommCnt
                End If
                If bActionOwner_HW = True Then
                    lCommCnt = lCommCnt + lHWCommCnt
                End If
                If bActionOwner_PE = True Then
                    lCommCnt = lCommCnt + lPECommCnt
                End If
                If lCommCnt > 0 Then
                    CheckFutActWithNoHold = True
                    GoTo ExitHandler
                End If
            End If
            'Add by Weilun on 20210907 for Project.BE 工業 3.5 Phase 36 - CP AGV導入專案 <end>
    
            If bFutAct = True Then
                CheckFutActWithNoHold = True
                GoTo ExitHandler
            End If
        End If
    End If
'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
            Case glERR_FAILTOQUERY
                typErrInfo.sUserText = typErrInfo.sErrDescription
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
            End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        CheckFutActWithNoHold = False
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Function

'================================================================================
' Sub: Clone_FutHoldForVirtualChildLotMerge()
' Description:  針對CP/FT LotComplete時, 虛擬子批併回母批, 將獨立的FutHold繼承給母批
'               搭配虛擬子批併批的Merge使用(oLotIds)
'               開發時為OutStepCheck , SortRunLotComplete, CPAutoLotComplete在Merge後會用到
' Author: Weilun Huang, MXIC on 2019/03/13 for ReqNo.BE#201900XXX
'================================================================================
Public Function Clone_FutHoldForVirtualChildLotMerge(ByRef oLogCtrl As Object, _
                                                     ByRef oProRawSQL As Object, _
                                                     ByRef oFwWIP As Object, _
                                                     ByVal sParentLotID As String, _
                                                     ByVal sUserName As String, _
                                                     ByVal oLotIds As FwStrings) As Boolean
                                 
                                 

                                                                  '
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim sSQL            As String
Dim colRaws         As Collection

Dim oItem           As Object
Dim sColumns        As String
Dim sChildLotIds    As String
Dim lIdx            As Long
'Dim oLotIds         As FwStrings '暫放
'----
' Init
'----
    sProcID = "Clone_FutHoldForVirtualChildLotMerge"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)
    Clone_FutHoldForVirtualChildLotMerge = True
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    '改寫參照: modCustom.Clone_FutAct
    sColumns = ""
    
    sSQL = "select column_name" & _
        " from user_tab_columns" & _
        " where table_name = '" & UCase(gsCAT_TBL_LOT_FUTACT) & "'" & _
        " order by column_id"
    Set colRaws = oProRawSQL.QueryDatabase(sSQL)
    For Each oItem In colRaws
        If oItem.Item(1) <> UCase(gsCAT_TLFA_CREATE_USER_ID) And _
           oItem.Item(1) <> UCase(gsCAT_TLFA_DELETE_USER_ID) And _
           oItem.Item(1) <> UCase(gsCAT_TLFA_UPDATE_USER_ID) And _
           oItem.Item(1) <> UCase(gsCAT_TLFA_CREATE_TIME) And _
           oItem.Item(1) <> UCase(gsCAT_TLFA_DELETE_TIME) And _
           oItem.Item(1) <> UCase(gsCAT_TLFA_UPDATE_TIME) And _
           oItem.Item(1) <> UCase(gsCAT_TLFA_PARENTLOTID) And _
           oItem.Item(1) <> UCase(gsCAT_TLFA_LOT_ID) Then
            sColumns = sColumns & oItem.Item(1) & ","
        End If
    Next oItem
    
    '取得子批清單, 在這裡FT和CP有點不同, oLotIds為外部輸入
    sChildLotIds = ""
    For lIdx = 1 To oLotIds.Count
        sChildLotIds = sChildLotIds & "'" & oLotIds.Item(lIdx) & "',"
    Next lIdx
    
    If sColumns <> "" And sChildLotIds <> "" Then
        sColumns = Left(sColumns, Len(sColumns) - 1) '去掉最後的','
        sChildLotIds = Left(sChildLotIds, Len(sChildLotIds) - 1)

        
        sSQL = "select count(*)" & _
                " from " & gsCAT_TBL_LOT_FUTACT & " " & _
               " where " & gsCAT_TLFA_HOLD_POSITION & " = '" & gsHOLD_POSITION_BEFORE_STEPOUT & "' " & _
                 " and " & gsCAT_TLFA_PARENTLOTID & " is null " & _
                 " and " & gsCAT_TLFA_LOT_ID & " in (" & sChildLotIds & ") " & _
                 " and " & gsCAT_TLFA_HOLD_REASON & " <> 'WaitMerge' " & _
                 " and " & gsCAT_TLFA_DELETE_FLAG & " = 'N' "

        Set colRaws = oProRawSQL.QueryDatabase(sSQL)
        If colRaws.Item(1).Item(1) > 0 Then
           
            sSQL = "insert into " & gsCAT_TBL_LOT_FUTACT & _
                " ( " & sColumns & ", " & _
                        gsCAT_TLFA_LOT_ID & ", " & _
                        gsCAT_TLFA_CREATE_USER_ID & ", " & _
                        gsCAT_TLFA_CREATE_TIME & " " & _
                " ) "
            
            sSQL = sSQL & _
              " select " & sColumns & ", " & _
                       " '" & sParentLotID & "', " & _
                       " '" & sUserName & "', " & _
                       " to_char(sysdate+rownum/24/60/60, 'YYYYMMDD HH24MISS')||'000' " & _
                " from " & gsCAT_TBL_LOT_FUTACT & " " & _
               " where " & gsCAT_TLFA_HOLD_POSITION & " = '" & gsHOLD_POSITION_BEFORE_STEPOUT & "' " & _
                 " and " & gsCAT_TLFA_PARENTLOTID & " is null " & _
                 " and " & gsCAT_TLFA_LOT_ID & " in (" & sChildLotIds & ") " & _
                 " and " & gsCAT_TLFA_HOLD_REASON & " <> 'WaitMerge' " & _
                 " and " & gsCAT_TLFA_DELETE_FLAG & " = 'N' "
            Call oProRawSQL.QueryDatabase(sSQL)
        End If
    End If
    
'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
            Case glERR_INVALIDOBJECT, glERR_FAILTOUPDATE
                typErrInfo.sUserText = typErrInfo.sErrDescription
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
            End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Clone_FutHoldForVirtualChildLotMerge = False
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Function

'================================================================================
' Sub: GetIPNPackingSpecDesc()
' Description: 輸入IPN取得IPN Master 的Carrier以及Boxing SpecDesc
' Author: Weilun Huang, MXIC on 2019/06/25 for ReqNo.BE#201900436
'================================================================================
Public Function GetIPNPackingSpecDesc(ByRef oLogCtrl As Object, _
                                      ByRef oProRawSQL As Object, _
                                      ByVal sIPN As String, _
                                      ByRef sSmallWQty As String, _
                                      ByRef sCarrierSpecDesc As String, _
                                      ByRef sBoxingSpecDesc As String, _
                                      ByRef sSmallCarrierSpecDesc As String, _
                                      ByRef sSmallBoxingSpecDesc As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim sSQL            As String
Dim colResult       As Collection

'----
' Init
'----
    sProcID = "GetIPNPackingSpecDesc"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)
    GetIPNPackingSpecDesc = False
    
    sSmallWQty = ""
    sCarrierSpecDesc = ""
    sBoxingSpecDesc = ""
    sSmallCarrierSpecDesc = ""
    sSmallBoxingSpecDesc = ""
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    sSQL = " select tim." & gsCAT_TIM_SMALL_WQTY & ", " & _
                  " tim." & gsCAT_TIM_CARRIER_SPEC_DESC & ", " & _
                  " tim." & gsCAT_TIM_BOXING_SPEC_DESC & ", " & _
                  " tim." & gsCAT_TIM_SMALL_CARRIER_SPEC_DESC & ", " & _
                  " tim." & gsCAT_TIM_SMALL_BOXING_SPEC_DESC & " " & _
             " from " & gsCAT_TBL_IPN_MASTER & " tim " & _
            " where tim." & gsCAT_TIM_IPN & " = '" & sIPN & "' " & _
              " and tim." & gsCAT_TIM_DELETE_FLAG & " = 'N' "

    Set colResult = oProRawSQL.QueryDatabase(sSQL)
    If colResult.Count = 1 Then
        sSmallWQty = Trim(colResult.Item(1).Item(gsCAT_TIM_SMALL_WQTY))
        sCarrierSpecDesc = Trim(colResult.Item(1).Item(gsCAT_TIM_CARRIER_SPEC_DESC))
        sBoxingSpecDesc = Trim(colResult.Item(1).Item(gsCAT_TIM_BOXING_SPEC_DESC))
        sSmallCarrierSpecDesc = Trim(colResult.Item(1).Item(gsCAT_TIM_SMALL_CARRIER_SPEC_DESC))
        sSmallBoxingSpecDesc = Trim(colResult.Item(1).Item(gsCAT_TIM_SMALL_BOXING_SPEC_DESC))
        
        GetIPNPackingSpecDesc = True
    End If
    
'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
            Case glERR_INVALIDOBJECT, glERR_FAILTOUPDATE
                typErrInfo.sUserText = typErrInfo.sErrDescription
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
            End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        GetIPNPackingSpecDesc = False
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Function

'================================================================================
' Sub: GetIPNPackingSpecDescByLotID()
' Description: 輸入LotID取得IPN Master 的Carrier以及Boxing SpecDesc, 並判斷是否為小批量
' Author: Weilun Huang, MXIC on 2019/06/25 for ReqNo.BE#201900436
'================================================================================
Public Function GetIPNPackingSpecDescByLotID(ByRef oLogCtrl As Object, _
                                             ByRef oProRawSQL As Object, _
                                             ByVal sLotID As String, _
                                             ByRef sSmallWQty As String, _
                                             ByRef bIsSmallWQty As Boolean, _
                                             ByRef sCarrierSpecDesc As String, _
                                             ByRef sBoxingSpecDesc As String, _
                                             ByRef sSmallCarrierSpecDesc As String, _
                                             ByRef sSmallBoxingSpecDesc As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim sSQL            As String
Dim colResult       As Collection

Dim sLotWQty        As String

'----
' Init
'----
    sProcID = "GetIPNPackingSpecDescByLotID"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)
    GetIPNPackingSpecDescByLotID = False
    
    sSmallWQty = ""
    bIsSmallWQty = False
    sCarrierSpecDesc = ""
    sBoxingSpecDesc = ""
    sSmallCarrierSpecDesc = ""
    sSmallBoxingSpecDesc = ""
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    sSQL = " select tlatt." & gsCAT_TLATT_IPN & ", " & _
                  " tlatt." & gsCAT_TLATT_WAFERQTY & " " & _
             " from " & gsCAT_TBL_LOT_ATTRIBUTE & " tlatt " & _
            " where tlatt." & gsCAT_TLATT_LOTID & " = '" & sLotID & "' "
            
    Set colResult = oProRawSQL.QueryDatabase(sSQL)
    If colResult.Count = 1 Then
        sLotWQty = Trim(colResult.Item(1).Item(gsCAT_TLATT_WAFERQTY))
        
        If GetIPNPackingSpecDesc(oLogCtrl, _
                                 oProRawSQL, _
                                 colResult.Item(1).Item(gsCAT_TLATT_IPN), _
                                 sSmallWQty, _
                                 sCarrierSpecDesc, _
                                 sBoxingSpecDesc, _
                                 sSmallCarrierSpecDesc, _
                                 sSmallBoxingSpecDesc _
                                ) = True Then
                                
            '當Tbl_IPN_Master.SmallWQty空白，維持現行包裝資訊不變動。                        (bIsSmallWQty = false)
            '當Tbl_IPN_Master.SmallWQty有值，需增加判斷勾選Lot片數。                         (This function)
            'Lot 's WaferQty > SmallWQty ' 維持現行包裝資訊不變動。                          (bIsSmallWQty = false)
            'Lot 's WaferQty <=SmallWQty ' 改顯示SmallBoxingSpecDesc與SmallCarrierSpecDesc。 (bIsSmallWQty = true)
            '當SmallWQty不為空值, 取Lot的WQty並與IPN的SmallQty比對,LotWQty <= sSmallWQty即為小批量
            If sSmallWQty <> "" And IsNumeric(sSmallWQty) = True And _
               sLotWQty <> "" And IsNumeric(sLotWQty) = True Then
                If CInt(sLotWQty) <= CInt(sSmallWQty) Then
                    bIsSmallWQty = True
                End If
            End If
            
            GetIPNPackingSpecDescByLotID = True
        End If
    End If
'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
            Case glERR_INVALIDOBJECT, glERR_FAILTOUPDATE
                typErrInfo.sUserText = typErrInfo.sErrDescription
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
            End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        GetIPNPackingSpecDescByLotID = False
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Function

'================================================================================
' Sub: ChkInklessGradeQty()
' Author: Jack Hsieh, MXIC on 2020/04/30 for BE MES Phase 47 - N19 KH-KTD 產品新增lot級別專案.
'
' << 注意事項 >>
' <1> Called by : frmLineRecLotStart(Inkless) / frmFtShippingLotStart(WsShippingLotStart)
' <2> 記得同步修改 AutoDropShipment 的部份.
'================================================================================
Public Function ChkInklessGradeQty(ByVal oLot As Object, _
                                    ByVal vIPN As String, _
                                    ByVal vLotCQty As Long, _
                                    ByVal sRuleName As String, _
                                    ByRef sReturnComment As String, _
                                    ByRef moAppLog As Object, _
                                    ByRef moProRawSql As Object, _
                                    ByRef moFwWIP As Object, _
                                    ByRef moFwWF As Object, _
                                    ByRef moCwMbx As Object, _
                                    ByVal sUserID As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo

Dim sGroupHistory           As String
Dim sInklessCreateTime      As String
Dim sInklessSource          As String
Dim bSkipInsertInkGradeQty  As Boolean

Dim colRS                   As Collection
Dim sSQL                    As String
Dim sSQL_InklessGradeQty    As String
Dim colRS_InklessGradeQty   As Collection
Dim lIdx                    As Integer

Dim sInklessLotID           As String
Dim sWaferCount             As String
Dim bHasData                As Boolean '是否有 Tbl_Inkless_GradeQty 資料 ?? (避免0筆資料)
Dim bIsWafetCountMatch      As Boolean
Dim bNeedHoldBinIssue       As Boolean 'Bin1/Bin2/Bin3是否有任一為空值?
Dim bIsBinQtyMatchLotQty    As Boolean '所有 WaferNO 的 Bin1/Bin2/Bin3加總, 是否等於 Lot's CQty ?
Dim bNeedToHoldLot          As Boolean
Dim lTmpCQty                As Long

Dim lIdx_WvmFinalPass       As Integer
Dim sInkGrade_LotID         As String
Dim sInkGrade_WaferNO       As String
Dim sInkGrade_BIN1          As String
Dim sInkGrade_BIN2          As String
Dim sInkGrade_BIN3          As String

Dim sHoldReason             As String
Dim sHoldCode               As String
Dim sHoldComment            As String

'----
' Init
'----
    sProcID = "ChkInklessGradeQty"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", moAppLog, glLOG_PROC, msMODULE_ID, sProcID)
    
    ChkInklessGradeQty = False
    sGroupHistory = sRuleName & "-" & GetTxnSeq(moProRawSql, moAppLog)
    
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
'傳入 sLotID, vIPN, oUser, vLotCQty, sGroupHistory
                    
    lTmpCQty = 0
    bHasData = False        '是否有 Tbl_Inkless_GradeQty 資料 ?? (避免0筆資料)
    bIsBinQtyMatchLotQty = False  '所有 WaferNO 的 Bin1/Bin2/Bin3加總, 是否等於 Lot's CQty ?
    bIsWafetCountMatch = False
    bNeedHoldBinIssue = False
    bNeedToHoldLot = False
    sReturnComment = ""
    
    bSkipInsertInkGradeQty = False
    sSQL = ""
    Set colRS = Nothing
    
    If modQuery.GetLotInfo_MergeType(CStr(oLot.Id), moProRawSql, moAppLog) = "CPSMALLLOT" Then
        Call modQuery.GetInklessMergeSource(CStr(oLot.Id), sInklessCreateTime, sInklessSource, moProRawSql, moAppLog)
        
        If UCase(sInklessSource) <> UCase("Inhouse") And UCase(sInklessSource) <> UCase("SubCon") And _
           UCase(sInklessSource) <> UCase("KTDMerge") Then
           '若非上述Source, 則不需檢查tbl_inkless_grade. (2020/05/05 Hanno確認)
            bSkipInsertInkGradeQty = True
        Else
            sSQL = "select " & gsCAT_TIML_CHILDLOTID & ", " & gsCAT_TIML_WAFERID & " as waferid, " & _
                         " FUN_SplitLen(WaferID, ';') as WaferCount " & _
                    " from " & gsCAT_TBL_INKLESS_MERGE_LIST & _
                   " where " & gsCAT_TIML_PARENTLOTID & " ='" & CStr(oLot.Id) & "' "
            If UCase(sInklessSource) = UCase("KTDMerge") Then
                sSQL = sSQL & _
                   " and " & gsCAT_TIML_CREATETIME & " ='" & sInklessCreateTime & "' "
            End If
        End If
        
    Else
        sSQL = "select " & gsCAT_TLI_LOT_ID & " as ChildLotID, " & gsCAT_TLI_WAFERID & " as waferid, " & _
                     " FUN_SplitLen(" & gsCAT_TLI_WAFERID & ", ';') as WaferCount " & _
                " from " & gsCAT_TBL_LOT_INFO & _
               " where " & gsCAT_TLI_LOT_ID & " ='" & CStr(oLot.Id) & "' "
    End If
    
    '非 Inhouse / SubCon / KTDMerge 不執行
    If sSQL <> "" Then
        Set colRS = moProRawSql.QueryDatabase(sSQL)
    End If
    
    If Not colRS Is Nothing And colRS.Count > 0 Then
        For lIdx = 1 To colRS.Count
            sInklessLotID = Mid(colRS.Item(lIdx).Item("ChildLotID"), 1, 8) '記錄ChildLotID前8碼.
            sWaferCount = colRS.Item(lIdx).Item("WaferCount")
                    
            sSQL_InklessGradeQty = "select x.lotid, x.waferno, x.bin1, x.bin2, x.bin3 " & _
                                    " FROM Tbl_Inkless_GradeQty X " & _
                                   " where x.lotid = substr('" & sInklessLotID & "', 1, 8) "
            If oLot.ComponentUnits <> gsLOTUNIT_CHIP Then
                sSQL_InklessGradeQty = sSQL_InklessGradeQty & _
                                     " and instr('" & colRS.Item(lIdx).Item("waferid") & "', x.waferno) > 0 "
            Else
                sSQL_InklessGradeQty = sSQL_InklessGradeQty & _
                                     " and x.waferno is not null "
            End If
            
            sSQL_InklessGradeQty = sSQL_InklessGradeQty & _
                                     " AND (x.waferno, CREATETIME) IN " & _
                                         " (select x.waferno, MAX(X.CREATETIME) " & _
                                            " FROM Tbl_Inkless_GradeQty X " & _
                                           " where x.lotid = substr('" & sInklessLotID & "', 1, 8) "
                                           
            If oLot.ComponentUnits <> gsLOTUNIT_CHIP Then
                sSQL_InklessGradeQty = sSQL_InklessGradeQty & _
                                     " and instr('" & colRS.Item(lIdx).Item("waferid") & "', x.waferno) > 0 "
            Else
                'Modified by Jack on 2021/04/09 for BE#202100094
                '增加 WaferNo 比對, 以免抓到屬於8碼LotID的WaferNo, 但不是本批的資料.
                sSQL_InklessGradeQty = sSQL_InklessGradeQty & _
                                     " and instr('" & colRS.Item(lIdx).Item("waferid") & "', x.waferno) > 0 " & _
                                     " and x.waferno is not null "
            End If
            
            sSQL_InklessGradeQty = sSQL_InklessGradeQty & " GROUP BY x.waferno) "

            Set colRS_InklessGradeQty = moProRawSql.QueryDatabase(sSQL_InklessGradeQty)
            If colRS_InklessGradeQty Is Nothing Then
                bNeedHoldBinIssue = True
            Else
                If Val(sWaferCount) <> colRS_InklessGradeQty.Count Then
                    bIsWafetCountMatch = True
                    bNeedHoldBinIssue = True
                End If
                If colRS_InklessGradeQty.Count > 0 Then
                    For lIdx_WvmFinalPass = 1 To colRS_InklessGradeQty.Count
                        bHasData = True
                        
                        sInkGrade_LotID = colRS_InklessGradeQty.Item(lIdx_WvmFinalPass).Item("lotid")
                        sInkGrade_WaferNO = colRS_InklessGradeQty.Item(lIdx_WvmFinalPass).Item("waferno")
                        sInkGrade_BIN1 = colRS_InklessGradeQty.Item(lIdx_WvmFinalPass).Item("bin1")
                        sInkGrade_BIN2 = colRS_InklessGradeQty.Item(lIdx_WvmFinalPass).Item("bin2")
                        sInkGrade_BIN3 = colRS_InklessGradeQty.Item(lIdx_WvmFinalPass).Item("bin3")
                        
                        If Trim(sInkGrade_BIN1) = "" Or Trim(sInkGrade_BIN2) = "" Or Trim(sInkGrade_BIN3) = "" Then
                            bNeedHoldBinIssue = True
                        End If
    
                        lTmpCQty = lTmpCQty + Val(sInkGrade_BIN1) + Val(sInkGrade_BIN2) + Val(sInkGrade_BIN3)
                        
                    Next lIdx_WvmFinalPass
                Else '0筆
                    bNeedHoldBinIssue = True
                End If
            End If
        Next lIdx
    Else
        If sSQL <> "" Then
            bNeedHoldBinIssue = True
        End If
    End If
    
    '所有 WaferNO 的 Bin1/Bin2/Bin3加總, 是否等於 Lot's CQty ?
    If lTmpCQty = Val(vLotCQty) Then
        bIsBinQtyMatchLotQty = True
    End If
    
    If (Not bIsBinQtyMatchLotQty) Or bNeedHoldBinIssue Then
        
        sSQL = "SELECT REASONCODE FROM TBL_REASONCODE WHERE GROUP1='PE' and CATEGORY='Department'"
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        If colRS.Count > 0 Then
            sHoldCode = colRS.Item(1).Item(1)
        Else
            sHoldCode = "MK330"
        End If
        
        If bNeedHoldBinIssue Then
            bNeedToHoldLot = True
            
            sHoldReason = "無分Bin資料"
            sHoldComment = modCustom.SqlString("Lot : " & oLot.Id & " 為需分Bin產品, 無分Bin資料") '
            sReturnComment = sHoldComment
            
            oLot.Refresh
            Call modCustom.HoldLot(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot, sHoldCode, sHoldReason, _
                                   gsHOLD_TYPE_LOT_HOLD, sUserID, sGroupHistory, sHoldComment)
            oLot.Refresh
        End If
        
        '<<<<< Inkless(frmLineRecLotStart) 不需 Hold Qty不一致 >>>>>.
        If UCase(sRuleName) = "WSSHIPLOTSTART" Then
            'Added by Jack on 2024/12/26 for UAT bug fixed (KTW 不需檢查QTY) <Start>
            sSQL = "select fwadmin.FUN_CHK_KTW('" & CStr(oLot.Id) & "') as prod from dual "
            Set colRS = moProRawSql.QueryDatabase(sSQL)
            If (Not colRS Is Nothing And colRS.Count > 0) And colRS.Item(1).Item(1) = "KTW" Then
                'do nothing
            'Added by Jack on 2024/12/26 for UAT bug fixed (KTW 不需檢查QTY) <End>
            Else
                '已 Hold "無分Bin資料", 不需再 Hold Qty不一致 (2020/05/05 Hanno確認)
                If (Not bNeedHoldBinIssue) And (Not bIsBinQtyMatchLotQty) Then
                    bNeedToHoldLot = True
                    
                    sHoldReason = "分Bin數量<>ChipQty"
                    sHoldComment = modCustom.SqlString("Lot : " & oLot.Id & " 分Bin數量 <> ChipQty") '
                    sReturnComment = sHoldComment
                
                    oLot.Refresh
                    Call modCustom.HoldLot(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot, sHoldCode, sHoldReason, _
                                           gsHOLD_TYPE_LOT_HOLD, sUserID, sGroupHistory, sHoldComment)
                    oLot.Refresh
                End If
            End If
        End If
    End If
    
    '正常結束 ChkInklessGradeQty 會是 True;
    '有 Hold 或 執行異常(sReturnComment=""), 則維持 False.
    If Not bNeedToHoldLot Then
        ChkInklessGradeQty = True
    End If
    
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
            Case glERR_INVALIDOBJECT, glERR_FAILTOUPDATE
                typErrInfo.sUserText = typErrInfo.sErrDescription
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
            End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        ChkInklessGradeQty = False
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Function

'================================================================================
' Sub: GetIPNProdAndRouteInfo()
'--------------------------------------------------------------------------------
' Description:  取得CP退庫所需的IPN資訊
'--------------------------------------------------------------------------------
' Author:       Create by Weilun on 20200707, for ReqNo.BE#202000195
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a Object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01]
'================================================================================
Public Function GetIPNProdAndRouteInfo(ByVal sIPN As String, _
                                       ByRef oFwPRP As Object, _
                                       ByRef oProRawSQL As Object, _
                                       ByRef oLogCtrl As Object, _
                                       ByRef sProdgroup As String, _
                                       ByRef sProdgroup_Version As String, _
                                       ByRef sRoute As String, _
                                       ByRef sRoute_Version As String) As Boolean
    On Error GoTo ExitHandler:

    Dim sProcID     As String
    Dim typErrInfo  As tErrInfo
    Dim colRaws     As Collection
    Dim oRaws       As Collection
    Dim sSQL        As String

    Dim oProd    As FwProduct
    Dim oRoute   As FwProcessPlan

'----
' Init
'----
    sProcID = "GetIPNProdAndRouteInfo"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    
    GetIPNProdAndRouteInfo = False
    
    sProdgroup = ""
    sProdgroup_Version = ""
    sRoute = ""
    sRoute_Version = ""
    
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    
'----
' Action
'----

    sSQL = " select " & gsCAT_TIM_PROD_GROUP & " " & _
             " from " & gsCAT_TBL_IPN_MASTER & " " & _
            " where " & gsCAT_TIM_IPN & " = '" & sIPN & "' " & _
              " and " & gsCAT_TIM_DELETE_FLAG & " = 'N' "
    Set colRaws = oProRawSQL.QueryDatabase(sSQL)
    
    If colRaws.Count > 0 Then
        '取得Prodgroup
        sProdgroup = colRaws.Item(1).Item(gsCAT_TIM_PROD_GROUP)
        Set oProd = oFwPRP.ProductById(sProdgroup)
        If Not oProd Is Nothing Then
            sProdgroup_Version = oProd.Version
            
            '取得Route
            sRoute = oProd.Processes.Item(1).Id
            Set oRoute = oFwPRP.PlanById(sRoute)
            If Not oRoute Is Nothing Then
                sRoute_Version = oRoute.Version
            End If
        End If
    End If
    
    GetIPNProdAndRouteInfo = True

'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Function

'================================================================================
'Name       : CanSkipMaxMergeCheck
'Description: 是否跳過MaxMerge的檢查, 用於frmMerge以及frmSpecialMerge
'Author     : Weilun huang, MXIC 2020/09/29 for ReqNo. BE#201900942
'Parameters : N/A
'================================================================================
Public Function CanSkipMaxMergeCheck(ByRef oProRawSQL As Object, _
                                     ByRef oLogCtrl As Object, _
                                     ByRef oSpread As fpSpread, _
                                     ByVal sLotID As String, _
                                     ByVal iLotCol As Integer) As Boolean
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo
    
    Dim sSQL                        As String
    Dim colSQLResult                As Collection
    
    Dim sIPN                        As String
    Dim sCargradeFlag               As String
    Dim vCheck                      As Variant
    Dim vLotID                      As Variant
    Dim sPaerentSapono              As String
    Dim sReworkPurpose              As String
    Dim sLotIdList                  As String
    Dim sStepName                   As String
    
    Dim lCnt                        As Long
    Dim iIndex                      As Integer
    '-----
    ' Init
    '-----
    sProcID = "CanSkipMaxMergeCheck"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl) '
    
    CanSkipMaxMergeCheck = False
    
    '--------------------
    ' Condition Checking
    '--------------------
    
    
    '-------
    ' Action
    '-------
    
    '關於SAPRWNO
    '主要是由frmReworkPrepare選擇併批後按cmdOK後寫入勾選Lot的LotInfo.Saprwno, 並於重工資訊Spread按ReworkIPN啟動frmMerge
    'SI表示有一種狀況是上面只作一半(比如失敗), 就可以從F1按進來
    
    sStepName = ""
    sPaerentSapono = ""
    sCargradeFlag = ""
    sReworkPurpose = ""
    
    '可能會有多張重工單(IPN等資料不一樣), 基本上是只有一個重工目的, 兩個以上則不Skip MaxMerge檢查
    sSQL = " select distinct " & _
                  " tlatt." & gsCAT_TLATT_STEPNAME & ", " & _
                  " tli." & gsCAT_TLI_SAPRWNO & ", " & _
                  " tim." & gsCAT_TIM_CARGRADEFLAG & ", " & _
                  " trr." & gsCAT_TRR_REWORKPURPOSE & " " & _
             " from " & gsCAT_TBL_LOT_ATTRIBUTE & " tlatt, " & _
                        gsCAT_TBL_LOT_INFO & " tli, " & _
                        gsCAT_TBL_IPN_MASTER & " tim, " & _
                        gsCAT_TBL_REWORK_REQ & " trr " & _
            " where trr." & gsCAT_TRR_DELETEFLAG & " = 'N' " & _
              " and tli." & gsCAT_TLI_SAPRWNO & " = trr." & gsCAT_TRR_SAPRWNO & " " & _
              " and tlatt." & gsCAT_TLATT_IPN & " = tim." & gsCAT_TIM_IPN & " " & _
              " and tlatt." & gsCAT_TLATT_LOTID & " = tli." & gsCAT_TLI_LOT_ID & " " & _
              " and tlatt." & gsCAT_TLATT_LOTID & "= '" & sLotID & "' "
    Set colSQLResult = oProRawSQL.QueryDatabase(sSQL)
    If colSQLResult.Count = 1 Then
        sStepName = colSQLResult.Item(1).Item(gsCAT_TLATT_STEPNAME)
        sPaerentSapono = colSQLResult.Item(1).Item(gsCAT_TLI_SAPRWNO)       '和其他Lot比對用
        sCargradeFlag = colSQLResult.Item(1).Item(gsCAT_TIM_CARGRADEFLAG)
        sReworkPurpose = colSQLResult.Item(1).Item(gsCAT_TRR_REWORKPURPOSE)
    End If
    
    '1. Lot站點需在MKBANK
    '2. 不屬於車規產品 (IPN Master.Cargradeflage='Y', 代表車規)
    '3. 具備制式特殊工單for Remark, Remark + FT & de-logo 這3種
    If sStepName = "MKBANK" And _
       sCargradeFlag <> "Y" And _
       (UCase(sReworkPurpose) = UCase("Remark") Or _
        UCase(sReworkPurpose) = UCase("Remark + FT") Or _
        UCase(sReworkPurpose) = UCase("De-Logo")) Then
        
        sLotIdList = ""
    
    
        '條件3附加, 唯一一種Sapwno, 且與母批相同, 同時數量與勾選數相同
        With oSpread
            For iIndex = 1 To .MaxRows
                .GetText 1, iIndex, vCheck
                If vCheck = "1" Then
                    lCnt = lCnt + 1

                    .GetText iLotCol, iIndex, vLotID
                    sLotIdList = sLotIdList & "," & CStr(vLotID)
                End If
            Next

            If Left(sLotIdList, 1) = "," Then
                sLotIdList = Mid(sLotIdList, 2)
            End If

        End With

        sSQL = " select " & gsCAT_TLI_SAPRWNO & ", count(*) as count " & _
                 " from " & gsCAT_TBL_LOT_INFO & " " & _
                " where " & gsCAT_TLI_LOT_ID & " in ('" & Replace(sLotIdList, ",", "','") & "') " & _
             " group by " & gsCAT_TLI_SAPRWNO & " "
        Set colSQLResult = oProRawSQL.QueryDatabase(sSQL)
        If colSQLResult.Count = 1 Then '唯一一種Sapwno, 且與母批相同, 同時數量與勾選數相同
            If colSQLResult.Item(1).Item(gsCAT_TLI_SAPRWNO) = sPaerentSapono And _
               colSQLResult.Item(1).Item("count") = CStr(lCnt) Then
                CanSkipMaxMergeCheck = True '不用作MaxMerge檢查
            End If
        End If
    End If
    
    '-------
    'Done
    '-------
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , oLogCtrl, True)

    End If
End Function

'================================================================================
'Name       : UpdateAGV_OrderEndTime()
'Description:
'Author     : Weilun huang, MXIC 2021/11/17 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
'Parameters : N/A
'================================================================================
Public Sub UpdateAGV_OrderEndTime(ByRef oProRawSQL As Object, _
                                  ByRef oLogCtrl As Object, _
                                  ByVal sEqID As String, _
                                  ByVal sEndTime As String)
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection

    '-----
    ' Init
    '-----
    sProcID = "UpdateAGV_OrderEndTime"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl) '
    

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    '機台ID必要, EndTime給空值表示清空
    If sEqID <> "" Then
        '資料更新
        sSQL = " update " & gsCAT_TBL_AGV_DISPATCH_ORDER & " " & _
                  " set " & gsCAT_TADO_PREVIOUS_LOT_ENDTIME & " = '" & sEndTime & "', " & _
                            gsCAT_TADO_UPDATEUSERID & " = 'SYS', " & _
                            gsCAT_TADO_UPDATETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                " where " & gsCAT_TADO_EQID & " = '" & sEqID & "' " & _
                  " and " & gsCAT_TADO_AGVID & " is null " & _
                  " and " & gsCAT_TADO_DELETEFLAG & " = 'N' "
        Call oProRawSQL.QueryDatabase(sSQL)
    End If
    '-------
    'Done
    '-------

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , oLogCtrl, True)

    End If
End Sub

'================================================================================
'Name       : UpdateAGV_LotEnd()
'Description:
'Author     : Weilun huang, MXIC 2021/11/17 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
'Parameters : N/A
'================================================================================
Public Sub UpdateAGV_LotEnd(ByRef oProRawSQL As Object, _
                            ByRef oLogCtrl As Object, _
                            ByRef oFwWIP As Object, _
                            ByVal sEqID As String, _
                            ByVal sLotID As String, _
                            Optional ByVal bIsLotComplete As Boolean = False)
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection
    
    Dim bIsVirtualMerge             As Boolean
    Dim bIsCPSmallLot               As Boolean
    Dim sVirtualLotId               As String
    Dim colVirtualList              As Collection
    
    Dim iListIndex                  As Long
    Dim sLotEndTime                 As String

    '-----
    ' Init
    '-----
    sProcID = "UpdateAGV_LotEnd"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl) '
    

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    '還得轉虛擬併批
    If sEqID <> "" And sLotID <> "" Then
        Call CheckMergeType(oProRawSQL, _
                            oLogCtrl, _
                            CStr(sLotID), _
                            bIsVirtualMerge, _
                            bIsCPSmallLot, _
                            sVirtualLotId)
        
        Set colVirtualList = New Collection
        If bIsVirtualMerge = True And sVirtualLotId <> "" Then
            '為虛擬併批批號, 需轉成底下LotID
            Set colVirtualList = modCPMerge.CPScanLotID(oProRawSQL, oLogCtrl, sVirtualLotId)
        End If
        
        '檢查是否都已經測完
        If bIsLotComplete = False Then 'LotComplete要到齊才能做, 這裡就不多做檢查
            If colVirtualList.Count > 0 Then '虛擬併批
                For iListIndex = 1 To colVirtualList.Count
                    If CheckWaferEndByLot(oProRawSQL, oLogCtrl, oFwWIP, _
                                          CStr(colVirtualList.Item(iListIndex).Item(gsCAT_TLI_LOT_ID))) = False Then
                        '有資料沒到齊
                        GoTo ExitHandler
                    End If
                Next
            Else '一般批
                If CheckWaferEndByLot(oProRawSQL, oLogCtrl, oFwWIP, sLotID) = False Then
                    '有資料沒到齊
                    GoTo ExitHandler
                End If
            End If
        End If
        
        '資料都已到齊, 更新下貨資料
        Call UpdateAGV_UnloadInfo(oProRawSQL, oLogCtrl, sEqID, sLotID)

        '由於監控還無法處理虛擬併批, 因此虛擬併批在SortLotEnd/SortLotComplete將前批結束時間更新為Now
        '注意SortLotEnd應該要先檢查是否已測完(bIsLotComplete = False)
        If bIsVirtualMerge = True And sVirtualLotId <> "" Then
            sLotEndTime = GetCurrentTime(oLogCtrl, oProRawSQL)
            Call UpdateAGV_OrderEndTime(oProRawSQL, oLogCtrl, sEqID, sLotEndTime)
            
        End If
        
    End If
    


    '-------
    'Done
    '-------

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , oLogCtrl, True)

    End If
End Sub

'================================================================================
'Name       : UpdateAGV_UnloadInfo()
'Description:
'Author     : Weilun huang, MXIC 2021/11/17 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
'Parameters : N/A
'================================================================================
Public Sub UpdateAGV_UnloadInfo(ByRef oProRawSQL As Object, _
                               ByRef oLogCtrl As Object, _
                               ByVal sEqID As String, _
                               ByVal sLotID As String)
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection
    Dim sWaferSize                  As String
    Dim sWaferQty                   As String
    
    Dim sUnloadLotID                As String
    Dim sUnloadEndTime              As String
    Dim sUnloadPcs                  As String
    
    Dim bIsVirtualMerge             As Boolean
    Dim bIsCPSmallLot               As Boolean
    Dim sVirtualLotId               As String
    Dim colVirtualList              As Collection
    Dim sVirtualWQty                As String

    '-----
    ' Init
    '-----
    sProcID = "UpdateAGV_UnloadInfo"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl) '
    

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    '還得轉虛擬併批
    If sEqID <> "" And sLotID <> "" Then
        Call CheckMergeType(oProRawSQL, _
                            oLogCtrl, _
                            CStr(sLotID), _
                            bIsVirtualMerge, _
                            bIsCPSmallLot, _
                            sVirtualLotId)
                            
        If bIsVirtualMerge = True And sVirtualLotId = sLotID Then
            '為虛擬併批批號, 需轉成底下LotID
            Set colVirtualList = modCPMerge.CPScanLotID(oProRawSQL, oLogCtrl, sVirtualLotId)
            
            If colVirtualList.Count > 0 Then
                sLotID = CStr(colVirtualList.Item(1).Item(gsCAT_TLI_LOT_ID)) '取第一個當代表
            End If
        End If
            
        '取得虛擬併批的總WQty
        sVirtualWQty = "0"
        If bIsVirtualMerge = True And sVirtualLotId <> "" Then
            sSQL = "select tli." & gsCAT_TLI_VIRTUALLOTID & ", " & _
                         " sum(tlatt." & gsCAT_TLATT_WAFERQTY & ") as VirtualWQty " & _
                    " from " & gsCAT_TBL_LOT_INFO & " tli, " & _
                               gsCAT_TBL_LOT_ATTRIBUTE & " tlatt " & _
                   " where tlatt." & gsCAT_TLATT_LOTID & " = tli." & gsCAT_TLI_LOT_ID & " " & _
                     " and tli." & gsCAT_TLI_VIRTUALLOTID & " = '" & sVirtualLotId & "'" & _
                     " and tli." & gsCAT_TLI_MERGETYPE & " = '" & gsLOT_MERGETYPE_VIRTUALMERGE & "' " & _
                " group by " & gsCAT_TLI_VIRTUALLOTID & " "
            
            Set colRaws = oProRawSQL.QueryDatabase(sSQL)
            If colRaws.Count > 0 Then
                sVirtualWQty = colRaws.Item(1).Item("VirtualWQty")
            End If
        End If
        
        '取得WaferSize
        sWaferSize = ""
        sWaferQty = "0"
        sSQL = "select tlatt." & gsCAT_TLATT_LOTID & ", " & _
                     " tlatt." & gsCAT_TLATT_WAFERQTY & ", " & _
                     " tpb." & gsCAT_TPB_WAFER_SIZE & " " & _
                " from " & gsCAT_TBL_LOT_ATTRIBUTE & " tlatt, " & _
                     gsCAT_TBL_IPN_MASTER & " tim, " & _
                     gsCAT_TBL_PROD_BODY & " tpb " & _
               " where tim." & gsCAT_TIM_PRODBODY & " = tpb." & gsCAT_TPB_PROD_BODY & " " & _
                 " and tlatt." & gsCAT_TLATT_IPN & " = tim." & gsCAT_TIM_IPN & " " & _
                 " and tlatt." & gsCAT_TLATT_LOTID & " = '" & sLotID & "' "
        Set colRaws = oProRawSQL.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
            sWaferSize = colRaws.Item(1).Item(gsCAT_TPB_WAFER_SIZE)
            sWaferQty = colRaws.Item(1).Item(gsCAT_TLATT_WAFERQTY)
        End If
    
        '寫入下貨資料
        If sWaferSize = "12" Then
            If bIsVirtualMerge = True And sVirtualLotId <> "" Then
                sUnloadLotID = sVirtualLotId
                sUnloadPcs = sVirtualWQty
            Else
                sUnloadLotID = sLotID
                sUnloadPcs = sWaferQty
            End If
            sUnloadEndTime = "to_char(sysdate, 'yyyymmdd hh24miss') || '000'" '含單引號
        Else
            sUnloadLotID = "NA"
            sUnloadPcs = "0"
            sUnloadEndTime = "'NA'" '含單引號
        End If
        
        '資料更新
        sSQL = " update " & gsCAT_TBL_AGV_EQ_STATUS & " " & _
                  " set " & gsCAT_TAES_LOTID & " = '" & sUnloadLotID & "', " & _
                            gsCAT_TAES_LOTPCS & " = '" & sUnloadPcs & "', " & _
                            gsCAT_TAES_LOTENDTIME & " = " & sUnloadEndTime & ", " & _
                            gsCAT_TAES_UPDATEUSERID & " = 'SYS', " & _
                            gsCAT_TAES_UPDATETIME & "  = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                " where " & gsCAT_TAES_EQID & " = '" & sEqID & "' " & _
                  " and " & gsCAT_TAES_DELETEFLAG & " = 'N' "
        Call oProRawSQL.QueryDatabase(sSQL)

    End If
    


    '-------
    'Done
    '-------

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , oLogCtrl, True)

    End If
End Sub

'================================================================================
'Name       : CheckWaferEndByLot()
'Description: 取得WS_TDS_SUM顆數資料, 修改自CPAutoLotEnd.frmMain.CheckWaferEnd()
'Author     : Weilun huang, MXIC 2021/11/23 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
'Parameters : N/A
'================================================================================
Public Function CheckWaferEndByLot(ByRef oProRawSQL As Object, _
                                   ByRef oLogCtrl As Object, _
                                   ByRef oFwWIP As Object, _
                                   ByVal sLotID As String) As Boolean
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo
    
    Dim sParentLotID                As String
    Dim sStepName                   As String
    
    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection
    Dim sTimeHereSince              As String
    Dim iIndex                      As Integer
    Dim sWaferIDList                As String
    Dim sCheckList                  As String
    Dim sGetWaferId                 As String
    
    Dim oComp                       As FwComponent
    Dim oLot                        As FwLot
    
    Dim sCurLotEndStatus            As String
    Dim sCurLotCompleteStatus       As String
    
    Dim colAllWafer                 As Collection
    Dim oWafer                      As Object
    
    '-----
    ' Init
    '-----
    sProcID = "CheckWaferEndByLot"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl) '
    
    CheckWaferEndByLot = False
    
    '--------------------
    ' Condition Checking
    '--------------------
    
    
    '-------
    ' Action
    '-------
    Set oLot = oFwWIP.LotById(sLotID)
    If oLot Is Nothing Then
        GoTo ExitHandler
    End If
    
    sTimeHereSince = ConvertToString(oLot.CurrentStep.Steps.Item(1).TimeHereSince, oLogCtrl)
    sStepName = oLot.CurrentStep.Steps.Item(1).Description
    
    sParentLotID = Left(sLotID, 8)

    '列出該Lot下的WaferId清單
    sWaferIDList = ""
    
    If oLot.CurrentStep.Steps.Item(1).Description Like "SORT*" And _
        Mid(oLot.CurrentStep.Steps.Item(1).Id, 3, 1) <> "2" Then
        '抽測以Pgm_List的為主
        sSQL = " select " & gsCAT_TPL_LOT_ID & ", " & _
                            gsCAT_TPL_USERWAFERID & " " & _
                 " from " & gsCAT_TBL_PGM_LIST & " " & _
                " where (" & gsCAT_TPL_LOT_ID & ", " & _
                             gsCAT_TPL_CREATE_TIME & ") in " & _
                       "( select " & gsCAT_TPL_LOT_ID & ", " & _
                           " max(" & gsCAT_TPL_CREATE_TIME & ") " & _
                          " from " & gsCAT_TBL_PGM_LIST & " " & _
                         " where " & gsCAT_TPL_LOT_ID & " ='" & sLotID & "' " & _
                           " and " & gsCAT_TPL_PROD_TYPE & " = 'WAFER' " & _
                           " and " & gsCAT_TPL_SOURCE & " <> 'Setup' " & _
                           " and " & gsCAT_TPL_CREATE_TIME & " >= '" & sTimeHereSince & "' " & _
                           " and " & gsCAT_TPL_DELETE_FLAG & " ='N' " & _
                      " group by " & gsCAT_TPL_LOT_ID & ") "
                           
        
        Set colRaws = oProRawSQL.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
            sWaferIDList = colRaws.Item(1).Item(gsCAT_TPL_USERWAFERID)
        End If
        
        '找不到資料或資料有誤
        If sWaferIDList = "" Then
            GoTo ExitHandler
        End If
        
    Else '原本的
        For iIndex = 1 To 25
            Set oComp = oLot.Components.Item(Format(iIndex, "00"))
            If Not oComp Is Nothing Then
                If oComp.Status <> modConstGlobal.gsCOMP_STATUS_SCRAPPED Then
                    sWaferIDList = sWaferIDList & "," & oComp.Id
                End If
            End If
        Next iIndex
    End If
    
    If Left(sWaferIDList, 1) = "," Then
        sWaferIDList = Mid(sWaferIDList, 2)
    End If

    '準備檢查用清單
    sWaferIDList = SortWaferID(sWaferIDList, ",")
    sCheckList = Replace(sWaferIDList, ",", " ")    '最後搭配Trim

    '搜尋tbl_cp_wafer_end, 尋找該Lot在自動化開始後相關的資料
    sSQL = " select " & gsCAT_TCWE_WAFERID & " " & _
             " from " & gsCAT_TBL_CP_WAFER_END & " " & _
            " where " & gsCAT_TCWE_LOTID & " like '" & sParentLotID & "%' " & _
              " and " & gsCAT_TCWE_WAFERID & " is not null " & _
              " and " & gsCAT_TCWE_STEPNAME & " = '" & sStepName & "' " & _
              " and " & gsCAT_TCWE_WAFERENDTIME & " > '" & sTimeHereSince & "' " & _
              " and " & gsCAT_TCWE_DELETEFLAG & " = 'N'"

    Set colRaws = oProRawSQL.QueryDatabase(sSQL)
    If colRaws.Count > 0 Then
        For Each oRaws In colRaws
            sGetWaferId = Format(oRaws.Item(gsCAT_TCWE_WAFERID), "00")
            If IsNumeric(sGetWaferId) = True And Len(sGetWaferId) = 2 Then
                sCheckList = Replace(sCheckList, sGetWaferId, "")  '有找到, 從檢查清單中扣除
            End If
        Next oRaws
    End If

    'WaferEnd找不到的再去找Tds Summary找
    If Trim(sCheckList) <> "" Then
    
        Set colAllWafer = GetWsTdsSumAllWafer(oProRawSQL, oLogCtrl, oFwWIP, oLot.Id, sTimeHereSince, oLot.CurrentStep.Steps.Item(1).Id)

        '檢查Summary ReviseQty與ReasonCode
        For Each oWafer In colAllWafer
            'TDS Good Dien有值
            '或(Revise Good Dien有值且Reasoncode有值)
            If Len(Trim(oWafer.Item(gsCAT_TWTS_TDS_GOOD_DIEN))) <> 0 Or _
               (Len(Trim(oWafer.Item(gsCAT_TWTS_REVISE_GOOD_DIEN))) <> 0 And _
                Len(Trim(oWafer.Item(gsCAT_TWTS_REASON_CODE))) <> 0) Then

                sGetWaferId = Format(oWafer.Item(gsCAT_TWTS_WAFER_IDN), "00")
                If IsNumeric(sGetWaferId) = True And Len(sGetWaferId) = 2 Then
                    sCheckList = Replace(sCheckList, sGetWaferId, "")  '有找到, 從檢查清單中扣除
                End If


            End If
        Next oWafer
    End If

    '尚未到齊
    If Trim(sCheckList) <> "" Then
        GoTo ExitHandler
    End If
    
    CheckWaferEndByLot = True    '通過檢查
    '-------
    'Done
    '-------

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , oLogCtrl, True)

    End If
End Function

'================================================================================
'Name       : GetWsTdsSumAllWafer()
'Description: 取得WS_TDS_SUM顆數資料, 修改自CPAutoLotEnd.frmMain.GetTdsTestSum()
'Author     : Weilun huang, MXIC 2021/11/23 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
'Parameters :
'================================================================================
Public Function GetWsTdsSumAllWafer(ByRef oProRawSQL As Object, _
                                    ByRef oLogCtrl As Object, _
                                    ByRef oFwWIP As Object, _
                                    ByVal sLotID As String, _
                                    ByVal sTimeHereSince As String, _
                                    ByVal sStepNo As String) As Collection
On Error GoTo ExitHandler:
Dim sProcID                 As String
Dim typErrInfo              As tErrInfo
Dim colRaws                 As Collection

Dim sSQL                    As String
Dim oRaws                   As Collection
Dim oLot                    As FwLot
Dim iIndex                  As Integer
Dim oComp                   As FwComponent

Dim sParentLot              As String

Dim colAllWafer             As Collection
Dim oWafer                  As Collection
Dim iCount                  As Integer
Dim bFoundData              As Boolean
'----
' Init
'----
    sProcID = "GetWsTdsSumAllWafer"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    If InStr(1, sLotID, ".") > 0 Then
        sParentLot = Left(Left(sLotID, InStr(1, sLotID, ".") - 1), 8)
    Else
        sParentLot = Left(sLotID, 8)
    End If
    
    
    sSQL = " select " & gsCAT_TWTS_WAFER_IDN & ", " & _
                        gsCAT_TWTS_TDS_GOOD_DIEN & ", " & _
                        gsCAT_TWTS_REVISE_GOOD_DIEN & ", " & _
                        gsCAT_TWTS_REASON_CODE & " " & _
             " from " & gsCAT_TBL_WS_TDS_SUM & " " & _
            " where (" & gsCAT_TWTS_TEST_MODE & ", " & _
                         gsCAT_TWTS_PARENT_LOT_ID & ", " & _
                         gsCAT_TWTS_WAFER_IDN & ", " & _
                         gsCAT_TWTS_STEPNO & ", " & _
                         gsCAT_TWTS_INSTEPTIME & ", " & _
                         "nvl(" & gsCAT_TWTS_TIME_STAMP & ",'NULL'))" & _
                  " IN"
                  
    sSQL = sSQL & _
                  " (SELECT " & gsCAT_TWTS_TEST_MODE & ", " & _
                                gsCAT_TWTS_PARENT_LOT_ID & ", " & _
                                gsCAT_TWTS_WAFER_IDN & ", " & _
                                gsCAT_TWTS_STEPNO & ", " & _
                                gsCAT_TWTS_INSTEPTIME & ", " & _
                                "NVL(MAX(" & gsCAT_TWTS_TIME_STAMP & "),'NULL')" & _
                     " FROM " & gsCAT_TBL_WS_TDS_SUM & " " & _
                    " WHERE " & gsCAT_TWTS_INSTEPTIME & " ='" & sTimeHereSince & "' " & _
                      " and " & gsCAT_TWTS_STEPNO & " ='" & sStepNo & "' " & _
                      " and " & gsCAT_TWTS_PARENT_LOT_ID & " ='" & sParentLot & "' " & _
                      " and " & gsCAT_TWTS_DELETE_FLAG & " ='N' " & _
                 " GROUP BY " & gsCAT_TWTS_TEST_MODE & ", " & _
                                gsCAT_TWTS_PARENT_LOT_ID & ", " & _
                                gsCAT_TWTS_WAFER_IDN & ", " & _
                                gsCAT_TWTS_STEPNO & ", " & _
                                gsCAT_TWTS_INSTEPTIME & " )" & _
          "order by " & " nvl(" & gsCAT_TWTS_TIME_STAMP & ",' ') desc"
             
    Set colRaws = oProRawSQL.QueryDatabase(sSQL)
    
    Set oLot = FwuRetrieveLot(oFwWIP, sLotID)
    Set colAllWafer = New Collection
    iCount = 0
    
    For iIndex = 1 To 25
        Set oComp = oLot.Components.Item(Format(iIndex, "00"))
        If Not oComp Is Nothing Then
            If oComp.Status <> modConstGlobal.gsCOMP_STATUS_SCRAPPED Then
                Set oWafer = New Collection
                bFoundData = False
                oWafer.Add oComp.Id, gsCAT_TWTS_WAFER_IDN

                '在搜尋結果裡尋找
                For Each oRaws In colRaws
                    If Trim(CStr(oRaws.Item(gsCAT_TWTS_WAFER_IDN))) = oComp.Id Then
                        bFoundData = True
                        oWafer.Add oRaws.Item(gsCAT_TWTS_TDS_GOOD_DIEN), gsCAT_TWTS_TDS_GOOD_DIEN
                        oWafer.Add oRaws.Item(gsCAT_TWTS_REVISE_GOOD_DIEN), gsCAT_TWTS_REVISE_GOOD_DIEN
                        oWafer.Add oRaws.Item(gsCAT_TWTS_REASON_CODE), gsCAT_TWTS_REASON_CODE
                        Exit For
                    End If
                Next oRaws
                
                If bFoundData = False Then
                    Exit For
                    oWafer.Add "", gsCAT_TWTS_TDS_GOOD_DIEN
                    oWafer.Add "", gsCAT_TWTS_REVISE_GOOD_DIEN
                    oWafer.Add "", gsCAT_TWTS_REASON_CODE
                End If
                
                iCount = iCount + 1
                colAllWafer.Add oWafer, CStr(iCount)
                Set oWafer = Nothing
            End If
        End If
        Set oComp = Nothing
    Next iIndex
    
    Set GetWsTdsSumAllWafer = colAllWafer
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , oLogCtrl, True)

    End If
End Function

'================================================================================
' Function: CheckStepInAndOut()
'--------------------------------------------------------------------------------
' Description:  Copy from OutStepCheck, 避免自動化卡關的提前檢查
'--------------------------------------------------------------------------------
' Author:       Weilun Huang,MXIC 2022/03/07 for Req.預先檢查自動化轉帳autolotcomplete背景作業卡關
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
Public Function CheckStepInAndOut(ByRef oProRawSQL As Object, _
                                  ByRef oLogCtrl As Object, _
                                  ByRef oFwWIP As Object, _
                                  ByRef oFwWF As Object, _
                                  ByRef oCwMbx As Object, _
                                  ByVal sLotID As String, _
                                  ByRef sCheckMsg As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID             As String
Dim typErrInfo          As tErrInfo

Dim sSQL                As String
Dim colRS               As Collection
Dim sIPN                As String

Dim oLot                As FwLot

Dim sErunTicNO          As String
Dim sFollowProd         As String
'----
' Init
'----
    sProcID = "CheckStepInAndOut"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    
    CheckStepInAndOut = False
    sCheckMsg = ""
    
    Set oLot = oFwWIP.LotById(sLotID)
    If oLot Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                        FormatErrorText(gsETX_INVALIDOBJECT, "FwLot"))
    End If
    
    
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
'----
' Action
'----
    ' <Put your Action codes here>...
    
    'OutStepCheck
        
    sErunTicNO = ""
    sFollowProd = ""
    Call modCustom.GetSapRwNoAndErunTicNo(oLogCtrl, oFwWIP, oFwWF, oCwMbx, _
                                          oLot.Id, sErunTicNO, "")
    
    If Len(Trim(sErunTicNO)) <> 0 Then
        Call modCustom.GetERunReqInfo(oLogCtrl, oFwWIP, oFwWF, oCwMbx, _
                                      oLot.Id, sErunTicNO, _
                                      oLot.CurrentStep.Steps.Item(1).Attributes(modConstFwAttr.gsSTEP_CUSTOMATTR_STAGE).Value, _
                                      sFollowProd)
                
        If oLot.PlanId <> gsERUN_ROUTE_ENGWS Then
            If oLot.CurrentStep.Steps.Item(1).Attributes(modConstFwAttr.gsSTEP_CUSTOMATTR_STAGE) = gsSTAGE_WS Then
                If sFollowProd = "N" And Not (Trim(sErunTicNO) Like "V*") Then
                    '架構保留, 於20220307時是空的
                    '在OutStepCheck包含Reassign, 以及在判斷前的取值
                    
                Else
                    If Not (Trim(sErunTicNO) Like "V*") Then
                        sIPN = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value
                        sSQL = "select c.ipn, c.prodgroup, b.keydata " & _
                                "from fwproductversion a, fwproductversion_n2m b, tbl_ipn_master c " & _
                                "Where C.IPN='" & sIPN & "' " & _
                                "and a.SysId = b.fromid " & _
                                "and b.linkname='processes' " & _
                                "and a.revstate='Active' " & _
                                "and a.productname=c.prodgroup " & _
                                "and c.deleteflag='N' "
                        Set colRS = oProRawSQL.QueryDatabase(sSQL)
                                                                    
                        If colRS.Count > 0 Then
                            If colRS.Item(1).Item(3) <> oLot.PlanId And CheckMCP_BINNO_Custom(oProRawSQL, oLogCtrl, oLot.Id) = False Then
                                '主要問題點, 請注意條件裡的CheckMCP_BINNO_Custom與原來OutStepCheck需要一致
                                sCheckMsg = "LotID:" & oLot.Id & " 此批工程品之Route與委測需求內容不符, " & vbCrLf & _
                                            "請call B/E PE 相關人員, 檢核CAT/Tbl_Erun_Req !!"
                                GoTo ExitHandler
                                    
                            Else
                                '架構保留, 於20220307時是空的
                                '在OutStepCheck已被註解
                                
                            End If
                        Else
                            '這裡順便做
                            sCheckMsg = "此批工程品要FollowProd,但IPN尚未Release或不存在!!" & vbCrLf & _
                                        "若IPN尚未Release請Call B/E PE !!" & vbNewLine & _
                                        "若IPN不存在請Call B/E PC !!" & vbNewLine & vbNewLine & _
                                        "完成上述動作後，請 B/E PC檢查ProdGroup、Route是否正確!!"
                            GoTo ExitHandler
                        End If
                    End If
                    '其他作業
                    
                End If
            End If
        End If
        '其他作業
        
    End If
    CheckStepInAndOut = True
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , oLogCtrl, True)

    End If
End Function


'================================================================================
' Function: CheckMCP_BINNO_Custom()
'--------------------------------------------------------------------------------
' Description:  Copy from OutStepCheck, 模組化版本, 加上"_Custom"避免GEN打包異常
'--------------------------------------------------------------------------------
' Author:       Weilun Huang,MXIC 2022/03/07 for Req.預先檢查自動化轉帳autolotcomplete背景作業卡關
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
Public Function CheckMCP_BINNO_Custom(ByRef oProRawSQL As Object, _
                                       ByRef oLogCtrl As Object, _
                                       ByVal sLotID As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

Dim sSQL As String
Dim colRS As Collection

'----
' Init
'----
    sProcID = "CheckMCP_BINNO_Custom"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl) '"Entering Function...", moAppLog, glLOG_PROC, msMODULE_ID, sProcID)
    CheckMCP_BINNO_Custom = False
    
'----
' Condition Checking
'----
' <Put your condition checking codes here>...

'----
' Action
'----
' <Put your Action codes here>...

    sSQL = " SELECT " & gsCAT_TLI_LOT_ID & _
           " From " & gsCAT_TBL_LOT_INFO & _
           " WHERE " & gsCAT_TLI_LOT_ID & " = '" & sLotID & "'" & _
           " AND " & gsCAT_TLI_MCP_FLAG & " = 'F'" & _
           " AND " & gsCAT_TLI_BINNO & " = 'BIN3'"
    Set colRS = oProRawSQL.QueryDatabase(sSQL)
            
    If Not colRS Is Nothing And colRS.Count > 0 Then
        CheckMCP_BINNO_Custom = True
    End If

'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Function

Public Function InsertWaferSplitReq_Single(ByVal sStage As String, ByVal sAction As String, ByVal sOrigWQty As String, _
                                    ByVal sLotID As String, ByVal sChildLotID As String, ByVal sWaferID As String, _
                                    ByVal sWaferQty As String, ByVal sChipQty As String, ByVal sComments As String, _
                                    ByVal sUserID As String, ByVal moProRawSql As Object, ByVal moAppLog As Object, _
                                    ByVal moFwWIP As Object, ByVal moFwWF As Object, ByVal moCwMbx As Object, _
                                    ByVal oLot As Object, ByVal sGroupHistkey As String, Optional ByVal sHoldComment As String = "", _
                                    Optional ByVal bNeedHoldLot As Boolean = False) As Boolean
'Added by Jack on 2022/07/08 for BE MES Phase 83 - CP併批放寬併批條件及功能改善
'bNeedHoldLot 說明 :
' <1> frmReworkPrepare 不需 Hold Lot.
' <2> frmAdjustLotAttribute 需要 Hold Lot.
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo

Dim sTicketNo       As String
Dim sSeqNo          As String
'Dim oLot            As FwLot

Dim sSQL            As String
Dim colRS           As Collection
Dim colRS2          As Collection
Dim lIdx            As Long

Dim sHoldCode       As String
Dim sHoldReason     As String

'----
' Init
'----
    InsertWaferSplitReq_Single = False
    
    sProcID = "InsertWaferSplitReq"
    
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    
'----
' Condition Chking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    sSQL = "select 'VS' || TO_CHAR(SYSDATE, 'YYMM') || nvl(LPAD(TO_NUMBER(SUBSTR(max(W." & gsCAT_TWSQ_TICKETNO & "), 7)) + 1, 6, '0'),'000001') "
    sSQL = sSQL & " FROM " & gsCAT_TBL_WAFERSPLIT_REQ & " W WHERE W." & gsCAT_TWSQ_TICKETNO & " LIKE 'VS' || TO_CHAR(SYSDATE, 'YYMM') || '%' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sTicketNo = colRS.Item(1).Item(1)
    End If
    
   sSQL = "Insert into " & gsCAT_TBL_WAFERSPLIT_REQ & " ( " & _
          " " & gsCAT_TWSQ_STAGE & " " & _
          " ," & gsCAT_TWSQ_LOTID & " " & _
          " ," & gsCAT_TWSQ_CHILDLOTID & " " & _
          " ," & gsCAT_TWSQ_ACTION & " " & _
          " ," & gsCAT_TWSQ_ORIWAFERQTY & " " & _
          " ," & gsCAT_TWSQ_TICKETNO & " " & _
          " ," & gsCAT_TWSQ_WAFERID & " " & _
          " ," & gsCAT_TWSQ_WAFERQTY & " " & _
          " ," & gsCAT_TWSQ_CHIPQTY & " " & _
          " ," & gsCAT_TWSQ_COMMENTS & " " & _
          " ," & gsCAT_TWSQ_SPLITFLAG & " " & _
          " ," & gsCAT_TWSQ_CREATETIME & " " & _
          " ," & gsCAT_TWSQ_CREATEUSERID & " " & _
          " ) "
    sSQL = sSQL & " values ( " & _
            " '" & sStage & "' " & _
            " ,'" & sLotID & "' " & _
            " ,'" & sChildLotID & "' " & _
            " ,'" & sAction & "' " & _
            " ,'" & sOrigWQty & "' " & _
            " ,'" & sTicketNo & "' " & _
            " ,'" & sWaferID & "' " & _
            " ,'" & sWaferQty & "' " & _
            " ,'" & sChipQty & "' " & _
            " ,'" & SqlString(sComments) & "' " & _
            " ,'N' " & _
            " ,to_char(sysdate, 'yyyymmdd hh24miss')||'000' " & _
            " ,'" & sUserID & "' " & _
            ") "
    Call moProRawSql.QueryDatabase(sSQL)
    
    
    If bNeedHoldLot Then
        '比照 AutoDropShipment
        'Modify by Sam on 20240426 for #184635,M0370改M0360
'        sSQL = "select " & gsCAT_TRCO_REASON_CODE & " " & _
'                " from " & gsCAT_TBL_REASON_CODE & " c " & _
'               " where c." & gsCAT_TRCO_CATEGORY & " = 'Department' " & _
'                 " and c." & gsCAT_TRCO_GROUP1 & " = 'WSPC' "
        sSQL = "select " & gsCAT_TRCO_REASON_CODE & " " & _
                " from " & gsCAT_TBL_REASON_CODE & " c " & _
               " where c." & gsCAT_TRCO_CATEGORY & " = 'Department' " & _
                 " and c." & gsCAT_TRCO_GROUP1 & " = 'FABPC' "
        Set colRS2 = moProRawSql.QueryDatabase(sSQL)
        If colRS2.Count > 0 Then
            sHoldCode = colRS2.Item(1).Item(gsCAT_TRCO_REASON_CODE)
        Else
             'Modify by Sam on 20240426 for #184635,M0370改M0360
            'sHoldCode = "M0370"
            'Modify by Sam on 20250117 for #200612
            'sHoldCode = "M0360"
            sHoldCode = gsDEPARTMENT_FABPC
        End If
                            
        sHoldReason = "Wait Split"
        
        oLot.Refresh
        Call modCustom.HoldLot(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot, _
            sHoldCode, sHoldReason, "LotHold", _
            sUserID, sGroupHistkey, "")
        oLot.Refresh
    End If
    
    InsertWaferSplitReq_Single = True
    
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
            Case glERR_INVALIDOBJECT, glERR_ERRMSG
                typErrInfo.sUserText = typErrInfo.sErrDescription
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
            End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Function


'================================================================================
' Sub: InsertMergeSplitLot()
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   moapplog            (I) [Optional] Valid reference to a Object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
'Added by Jack on 2022/07/08 for BE MES Phase 83 - CP併批放寬併批條件及功能改善
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'================================================================================
Public Sub InsertMergeSplitLot_Single(ByVal sMergeLotID As String, ByVal sOriginalLotID As String, _
                                ByVal sComments As String, ByVal sUserID As String, _
                                ByVal moProRawSql As Object, ByVal moAppLog As Object)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

Dim sSQL As String
Dim colRS As Collection

'----
' Init
'----
    sProcID = "InsertMergeSplitLot_Single"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    sSQL = "select ms." & gsCAT_TMGS_MERGELOTID & " from " & gsCAT_TBL_MERGE_SPLITLOT & " ms " & _
           " where ms." & gsCAT_TMGS_DELETEFLAG & "= 'N' " & _
           " and ms." & gsCAT_TMGS_MERGELOTID & " = '" & sMergeLotID & "' " & _
           " and ms." & gsCAT_TMGS_ORIGINALLOTID & " = '" & sOriginalLotID & "' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count = 0 Then
        sSQL = "Insert into " & gsCAT_TBL_MERGE_SPLITLOT & " ( " & _
              " " & gsCAT_TMGS_MERGELOTID & " " & _
              " ," & gsCAT_TMGS_ORIGINALLOTID & " " & _
              " ," & gsCAT_TMGS_SPLITFLAG & " " & _
              " ," & gsCAT_TMGS_COMMENTS & " " & _
              " ," & gsCAT_TMGS_CREATETIME & " " & _
              " ," & gsCAT_TMGS_CREATEUSERID & " " & _
              " ) "
        sSQL = sSQL & " values ( " & _
                " '" & sMergeLotID & "' " & _
                " ,'" & sOriginalLotID & "' " & _
                " ,'Y' " & _
                " ,'" & SqlString(sComments) & "' " & _
                " ,to_char(sysdate, 'yyyymmdd hh24miss')||'000' " & _
                " ,'" & sUserID & "' " & _
                ") "
        Call moProRawSql.QueryDatabase(sSQL)
    End If
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
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Sub


'================================================================================
' Sub: InsertReworkRecord()
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   moapplog            (I) [Optional] Valid reference to a Object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
'Added by Jack on 2022/07/08 for BE MES Phase 83 - CP併批放寬併批條件及功能改善
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'================================================================================
Public Sub InsertReworkRecord(ByVal moProRawSql As Object, ByVal moAppLog As Object, _
                              ByVal sStage As String, ByVal sReworkType As String, _
                              ByVal sSource As String, ByVal sLotID As String, _
                              ByVal sChildLotID As String, ByVal sAction As String, _
                              ByVal sWaferID As String, ByVal sIBBIN As String, _
                              ByVal sStepID As String, ByVal sStepName As String, _
                              ByVal sRelUserID As String, _
                              ByVal sReworkReason1 As String, ByVal sReworkReason2 As String, _
                              ByVal sUserID As String)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

Dim sRelDept As String

Dim sSQL As String
Dim colRS As Collection

'----
' Init
'----
    sProcID = "InsertReworkRecord"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    If sSource <> "MES" Then
        GoTo ExitHandler
    End If
    
    
    If sRelUserID <> "" Then
        sSQL = "select a." & gsCAT_TME_DEPTNO & " from " & gsCAT_TBL_MXIC_EMP & " a " & _
              " where a." & gsCAT_TME_EMPNO & " = '" & sRelUserID & "' "
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        If colRS.Count > 0 Then
            sRelDept = colRS.Item(1).Item(gsCAT_TME_DEPTNO)
        End If
    End If
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    
    sSQL = "insert into " & gsCAT_TBL_REWORK_RECORD & " " & _
          "( " & gsCAT_TRRD_STAGE & "," & _
           "" & gsCAT_TRRD_REWORKTYPE & "," & _
           "" & gsCAT_TRRD_SOURCE & "," & _
           "" & gsCAT_TRRD_LOTID & "," & _
           "" & gsCAT_TRRD_CHILDLOTID & "," & _
           "" & gsCAT_TRRD_ACTION & "," & _
           "" & gsCAT_TRRD_WAFERID & "," & _
           "" & gsCAT_TRRD_IB_BIN & "," & _
           "" & gsCAT_TRRD_STEPID & "," & _
           "" & gsCAT_TRRD_STEPNAME & "," & _
           "" & gsCAT_TRRD_RELEASEUSERID & "," & _
           "" & gsCAT_TRRD_RELEASE_DEPT & "," & _
           "" & gsCAT_TRRD_REASONLEVEL1 & "," & _
           "" & gsCAT_TRRD_REASONLEVEL2 & "," & _
           "" & gsCAT_TRRD_CREATEUSERID & "," & _
           "" & gsCAT_TRRD_CREATETIME & " ) " & _
        "values ( "
    sSQL = sSQL & " " & _
                 " '" & sStage & "' " & _
                 ",'" & sReworkType & "' " & _
                 ",'" & sSource & "' " & _
                 ",'" & sLotID & "' " & _
                 ",'" & sChildLotID & "' " & _
                 ",'" & sAction & "' " & _
                 ",'" & sWaferID & "' " & _
                 ",'" & sIBBIN & "' " & _
                 ",'" & sStepID & "' " & _
                 ",'" & sStepName & "' " & _
                 ",'" & sRelUserID & "' " & _
                 ",'" & sRelDept & "' " & _
                 ",'" & sReworkReason1 & "' " & _
                 ",'" & sReworkReason2 & "' " & _
                 ",'" & sUserID & "' " & _
                 ", TO_CHAR(SYSDATE,'YYYYMMDD HH24MISS') || '000' "
    
    sSQL = sSQL & " ) "
        
    Call moProRawSql.QueryDatabase(sSQL)
    
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
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Sub

