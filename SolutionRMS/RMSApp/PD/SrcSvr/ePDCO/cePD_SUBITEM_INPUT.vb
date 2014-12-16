
'****************************************************************************************
'시스템구분 : RMS/PD/Server Entity Class
'실행  환경 : 
'프로그램명 : cePD_SUBITEM_DTL.vb (PD_SUBITEM_DTL Entity 처리 Class)
'기      능 : PD_SUBITEM_DTL Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009-10-19 
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class cePD_SUBITEM_INPUT
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "cePD_SUBITEM_INPUT"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    'PREESTNO,ITEMCODESEQ,ITEMCODE,SEQ,SORTSEQ,PRICE,QTY,TERM,AMT,MEMO,CONFIRMGBN
    Public Function InsertDo(ByVal strPREESTNO As String, _
             Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
             Optional ByVal dblSUBITEMCODESEQ As Double = OPTIONAL_NUM, _
             Optional ByVal dblITEMCODESEQ As Double = OPTIONAL_NUM, _
             Optional ByVal strITEMCODE As String = OPTIONAL_STR, _
             Optional ByVal dblIMESEQ As Double = OPTIONAL_NUM, _
             Optional ByVal dblPRINT_SEQ As Double = OPTIONAL_NUM, _
             Optional ByVal dblSORTSEQ As Double = OPTIONAL_NUM, _
             Optional ByVal dblPRICE As Double = OPTIONAL_NUM, _
             Optional ByVal dblQTY As Double = OPTIONAL_NUM, _
             Optional ByVal dblTERM As Double = OPTIONAL_NUM, _
             Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
             Optional ByVal strMEMO As String = OPTIONAL_STR, _
             Optional ByVal dblEXEPRICE As Double = OPTIONAL_NUM, _
             Optional ByVal dblEXEQTY As Double = OPTIONAL_NUM, _
             Optional ByVal dblEXETERM As Double = OPTIONAL_NUM, _
             Optional ByVal dblEXEAMT As Double = OPTIONAL_NUM, _
             Optional ByVal strEXEMEMO As String = OPTIONAL_STR, _
             Optional ByVal strCONFIRMGBN As String = OPTIONAL_STR, _
             Optional ByVal strNEWFLAG As String = OPTIONAL_STR, _
             Optional ByVal strSUBITEMNAME As String = OPTIONAL_STR, _
             Optional ByVal strATTR01 As String = OPTIONAL_STR, _
             Optional ByVal strATTR02 As String = OPTIONAL_STR, _
             Optional ByVal strATTR03 As String = OPTIONAL_STR, _
             Optional ByVal strATTR04 As String = OPTIONAL_STR, _
             Optional ByVal strATTR05 As String = OPTIONAL_STR, _
             Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
             Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
             Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
             Optional ByVal dblATTR09 As Double = OPTIONAL_NUM, _
             Optional ByVal dblATTR10 As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            BuildNameValues(",", "PREESTNO", strPREESTNO, strFields, strValues)
            BuildNameValues(",", "SEQ", dblSEQ, strFields, strValues)
            BuildNameValues(",", "SUBITEMCODESEQ", dblSUBITEMCODESEQ, strFields, strValues)
            BuildNameValues(",", "ITEMCODESEQ", dblITEMCODESEQ, strFields, strValues)
            BuildNameValues(",", "ITEMCODE", strITEMCODE, strFields, strValues)
            BuildNameValues(",", "IMESEQ", dblIMESEQ, strFields, strValues)
            BuildNameValues(",", "PRINT_SEQ", dblPRINT_SEQ, strFields, strValues)
            BuildNameValues(",", "SORTSEQ", dblSORTSEQ, strFields, strValues)
            BuildNameValues(",", "PRICE", dblPRICE, strFields, strValues)
            BuildNameValues(",", "QTY", dblQTY, strFields, strValues)
            BuildNameValues(",", "TERM", dblTERM, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "EXEPRICE", dblEXEPRICE, strFields, strValues)
            BuildNameValues(",", "EXEQTY", dblEXEQTY, strFields, strValues)
            BuildNameValues(",", "EXETERM", dblEXETERM, strFields, strValues)
            BuildNameValues(",", "EXEAMT", dblEXEAMT, strFields, strValues)
            BuildNameValues(",", "EXEMEMO", strEXEMEMO, strFields, strValues)
            BuildNameValues(",", "CONFIRMGBN", strCONFIRMGBN, strFields, strValues)
            BuildNameValues(",", "NEWFLAG", strNEWFLAG, strFields, strValues)
            BuildNameValues(",", "SUBITEMNAME", strSUBITEMNAME, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)
            BuildNameValues(",", "ATTR02", strATTR02, strFields, strValues)
            BuildNameValues(",", "ATTR03", strATTR03, strFields, strValues)
            BuildNameValues(",", "ATTR04", strATTR04, strFields, strValues)
            BuildNameValues(",", "ATTR05", strATTR05, strFields, strValues)
            BuildNameValues(",", "ATTR06", dblATTR06, strFields, strValues)
            BuildNameValues(",", "ATTR07", dblATTR07, strFields, strValues)
            BuildNameValues(",", "ATTR08", dblATTR08, strFields, strValues)
            BuildNameValues(",", "ATTR09", dblATTR09, strFields, strValues)
            BuildNameValues(",", "ATTR10", dblATTR10, strFields, strValues)
            BuildNameValues(",", "CUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "CDATE", strNOW, strFields, strValues)
            BuildNameValues(",", "UUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "UDATE", strNOW, strFields, strValues)

            strSQL = String.Format("INSERT INTO {0} ({1}) VALUES({2})", EntityName, strFields, strValues)

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Update 처리
    '참고 : Key 조건과 Value Field가선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    '
    Public Function UpdateDo(ByVal strPREESTNO As String, _
           Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
           Optional ByVal dblSUBITEMCODESEQ As Double = OPTIONAL_NUM, _
           Optional ByVal dblITEMCODESEQ As Double = OPTIONAL_NUM, _
           Optional ByVal strITEMCODE As String = OPTIONAL_STR, _
           Optional ByVal dblIMESEQ As Double = OPTIONAL_NUM, _
           Optional ByVal dblPRINT_SEQ As Double = OPTIONAL_NUM, _
           Optional ByVal dblSORTSEQ As Double = OPTIONAL_NUM, _
           Optional ByVal dblPRICE As Double = OPTIONAL_NUM, _
           Optional ByVal dblQTY As Double = OPTIONAL_NUM, _
           Optional ByVal dblTERM As Double = OPTIONAL_NUM, _
           Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
           Optional ByVal strMEMO As String = OPTIONAL_STR, _
           Optional ByVal dblEXEPRICE As Double = OPTIONAL_NUM, _
           Optional ByVal dblEXEQTY As Double = OPTIONAL_NUM, _
           Optional ByVal dblEXETERM As Double = OPTIONAL_NUM, _
           Optional ByVal dblEXEAMT As Double = OPTIONAL_NUM, _
           Optional ByVal strEXEMEMO As String = OPTIONAL_STR, _
           Optional ByVal strCONFIRMGBN As String = OPTIONAL_STR, _
           Optional ByVal strNEWFLAG As String = OPTIONAL_STR, _
           Optional ByVal strSUBITEMNAME As String = OPTIONAL_STR, _
           Optional ByVal strATTR01 As String = OPTIONAL_STR, _
           Optional ByVal strATTR02 As String = OPTIONAL_STR, _
           Optional ByVal strATTR03 As String = OPTIONAL_STR, _
           Optional ByVal strATTR04 As String = OPTIONAL_STR, _
           Optional ByVal strATTR05 As String = OPTIONAL_STR, _
           Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
           Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
           Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
           Optional ByVal dblATTR09 As Double = OPTIONAL_NUM, _
           Optional ByVal dblATTR10 As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("SUBITEMCODESEQ", dblSUBITEMCODESEQ), _
                        GetFieldNameValue("ITEMCODESEQ", dblITEMCODESEQ), _
                        GetFieldNameValue("ITEMCODE", strITEMCODE), _
                        GetFieldNameValue("IMESEQ", dblIMESEQ), _
                        GetFieldNameValue("PRINT_SEQ", dblPRINT_SEQ), _
                        GetFieldNameValue("SORTSEQ", dblSORTSEQ), _
                        GetFieldNameValue("PRICE", dblPRICE), _
                        GetFieldNameValue("QTY", dblQTY), _
                        GetFieldNameValue("TERM", dblTERM), _
                        GetFieldNameValue("AMT", dblAMT), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("EXEPRICE", dblEXEPRICE), _
                        GetFieldNameValue("EXEQTY", dblEXEQTY), _
                        GetFieldNameValue("EXETERM", dblEXETERM), _
                        GetFieldNameValue("EXEAMT", dblEXEAMT), _
                        GetFieldNameValue("EXEMEMO", strEXEMEMO), _
                        GetFieldNameValue("CONFIRMGBN", strCONFIRMGBN), _
                        GetFieldNameValue("NEWFLAG", strNEWFLAG), _
                        GetFieldNameValue("SUBITEMNAME", strSUBITEMNAME), _
                        GetFieldNameValue("ATTR01", strATTR01), _
                        GetFieldNameValue("ATTR02", strATTR02), _
                        GetFieldNameValue("ATTR03", strATTR03), _
                        GetFieldNameValue("ATTR04", strATTR04), _
                        GetFieldNameValue("ATTR05", strATTR05), _
                        GetFieldNameValue("ATTR06", dblATTR06), _
                        GetFieldNameValue("ATTR07", dblATTR07), _
                        GetFieldNameValue("ATTR08", dblATTR08), _
                        GetFieldNameValue("ATTR09", dblATTR09), _
                        GetFieldNameValue("ATTR10", dblATTR10), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("PREESTNO", strPREESTNO), GetFieldNameValue("SEQ", dblSEQ), GetFieldNameValue("SUBITEMCODESEQ", dblSUBITEMCODESEQ), GetFieldNameValue("ITEMCODESEQ", dblITEMCODESEQ), GetFieldNameValue("ITEMCODE", strITEMCODE)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo")
        End Try
    End Function


    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function DeleteDo(ByVal strPREESTNO As String, _
                             ByVal dblSEQ As Double, _
                             ByVal dblSUBITEMCODESEQ As Double, _
                             ByVal dblITEMCODESEQ As Double, _
                             ByVal strITEMCODE As String) As Integer
        'dblSUBITEMCODESEQ,dblITEMCODESEQ,strITEMCODE
        Dim strSQL As String

        Try
            strSQL = "DELETE FROM PD_SUBITEM_INPUT WHERE PREESTNO = '" & strPREESTNO & "' AND SEQ =" & dblSEQ & " AND SUBITEMCODESEQ =" & dblSUBITEMCODESEQ & " AND ITEMCODESEQ=" & dblITEMCODESEQ & " AND ITEMCODE = '" & strITEMCODE & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)strPREESTNO, strITEMCODE, strITEMCODESEQ, strUSER)
    '*****************************************************************
    Public Function DeleteRnt_Input(ByVal strJOBNO As String, _
                                    ByVal strITEMCODE As String, _
                                    ByVal strITEMCODESEQ As Double, _
                                    ByVal strUSER As String) As Integer
        'dblSUBITEMCODESEQ,dblITEMCODESEQ,strITEMCODE
        Dim strSQL As String

        Try
            strSQL = "DELETE FROM PD_SUBITEM_INPUT"
            strSQL = strSQL & " WHERE ATTR01 = '" & strJOBNO & "'"
            strSQL = strSQL & " AND ITEMCODE = '" & strITEMCODE & "' "
            strSQL = strSQL & " AND ITEMCODESEQ = '" & strITEMCODESEQ & "' "
            strSQL = strSQL & " AND CUSER = '" & strUSER & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteRnt_Input")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Select 처리
    '*****************************************************************
    Public Function SelectDo(ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                            Optional ByVal strSEQNO As String = OPTIONAL_STR, _
                            Optional ByVal strSelFields As String = "*", _
                            Optional ByVal intLimitRow As Integer = 0, _
                            Optional ByVal intSelMode As Integer = SELMODE.ARR, _
                            Optional ByVal blnBindingHeader As Boolean = False) As Object
        Dim strSQL As String
        Dim strKeyFields As String

        Try
            strKeyFields = BuildFields("AND", _
                                    GetFieldNameValue("SEQNO", strSEQNO))

            Return SelectDoExt(intRowCnt, intColCnt, strSelFields, strKeyFields, intLimitRow, intSelMode, blnBindingHeader)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".SeleteDo")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : strSQL 해당문을 그대로 처리 청구견적 변경시 외주비 정산 헤더 무조건 업데이트
    '*****************************************************************
    Public Function SqlExe(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".SqlExe")
        End Try
    End Function
#End Region

#Region "객체 생성/해제"
    '*****************************************************************
    '입력 : strInfoXML = 공통기본정보에 대한 XML
    'objSCGLSql = DB 처리 객체 인스턴싱 변수    '반환 : 없음
    '기능 : DB 처리를 위한 공통기본정보 설정
    '*****************************************************************
    Public Sub New(Optional ByVal objSCGLConfig As SCGLUtil.cbSCGLConfig = Nothing, Optional ByVal strInfoXML As String = "")
        MyBase.SetConfig(objSCGLConfig, strInfoXML)
        MyBase.EntityName = "PD_SUBITEM_INPUT"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class