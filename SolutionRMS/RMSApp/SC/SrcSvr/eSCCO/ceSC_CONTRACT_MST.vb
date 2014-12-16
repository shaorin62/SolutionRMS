
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-01-14 오전 11:09:29 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체
Public Class ceSC_CONTRACT_MST
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceSC_CONTRACT_MST"    '자신의 클래스명
#End Region

#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Update 처리
    '참고 : Key 조건과 Value Field가선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************

    Public Function InsertDo(Optional ByVal strSEQ As Double = OPTIONAL_NUM, _
                             Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strCUSTNAME As String = OPTIONAL_STR, _
                             Optional ByVal strCONTRACTNO As String = OPTIONAL_STR, _
                             Optional ByVal strCONTRACTNAME As String = OPTIONAL_STR, _
                             Optional ByVal strGBN As String = OPTIONAL_STR, _
                             Optional ByVal strAMT As Double = OPTIONAL_NUM, _
                             Optional ByVal strSTDATE As String = OPTIONAL_STR, _
                             Optional ByVal strEDDATE As String = OPTIONAL_STR, _
                             Optional ByVal strCONTRACTDAY As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRMDATE As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRM_USER As String = OPTIONAL_STR, _
                             Optional ByVal strMEMO As String = OPTIONAL_STR, _
                             Optional ByVal strCONDITION As String = OPTIONAL_STR, _
                             Optional ByVal strATTR01 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR02 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR03 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR04 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR05 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR06 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR07 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR08 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR09 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR10 As Double = OPTIONAL_NUM)


        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            BuildNameValues(",", "SEQ", strSEQ, strFields, strValues)
            BuildNameValues(",", "CUSTCODE", strCUSTCODE, strFields, strValues)
            BuildNameValues(",", "CUSTNAME", strCUSTNAME, strFields, strValues)
            BuildNameValues(",", "CONTRACTNO", strCONTRACTNO, strFields, strValues)
            BuildNameValues(",", "CONTRACTNAME", strCONTRACTNAME, strFields, strValues)
            BuildNameValues(",", "GBN", strGBN, strFields, strValues)
            BuildNameValues(",", "AMT", strAMT, strFields, strValues)
            BuildNameValues(",", "STDATE", strSTDATE, strFields, strValues)
            BuildNameValues(",", "EDDATE", strEDDATE, strFields, strValues)
            BuildNameValues(",", "CONTRACTDAY", strCONTRACTDAY, strFields, strValues)
            BuildNameValues(",", "CONFIRMFLAG", strCONFIRMFLAG, strFields, strValues)
            BuildNameValues(",", "CONFIRMDATE", strCONFIRMDATE, strFields, strValues)
            BuildNameValues(",", "CONFIRM_USER", strCONFIRM_USER, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "CONDITION", strCONDITION, strFields, strValues)
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

    Public Function UpdateDo(Optional ByVal strSEQ As Double = OPTIONAL_NUM, _
                             Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strCUSTNAME As String = OPTIONAL_STR, _
                             Optional ByVal strCONTRACTNO As String = OPTIONAL_STR, _
                             Optional ByVal strCONTRACTNAME As String = OPTIONAL_STR, _
                             Optional ByVal strGBN As String = OPTIONAL_STR, _
                             Optional ByVal strAMT As Double = OPTIONAL_NUM, _
                             Optional ByVal strSTDATE As String = OPTIONAL_STR, _
                             Optional ByVal strEDDATE As String = OPTIONAL_STR, _
                             Optional ByVal strCONTRACTDAY As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRMDATE As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRM_USER As String = OPTIONAL_STR, _
                             Optional ByVal strMEMO As String = OPTIONAL_STR, _
                             Optional ByVal strCONDITION As String = OPTIONAL_STR, _
                             Optional ByVal strATTR01 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR02 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR03 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR04 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR05 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR06 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR07 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR08 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR09 As Double = OPTIONAL_NUM, _
                             Optional ByVal strATTR10 As Double = OPTIONAL_NUM)

        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("CUSTCODE", strCUSTCODE), _
                        GetFieldNameValue("CUSTNAME", strCUSTNAME), _
                        GetFieldNameValue("CONTRACTNO", strCONTRACTNO), _
                        GetFieldNameValue("CONTRACTNAME", strCONTRACTNAME), _
                        GetFieldNameValue("GBN", strGBN), _
                        GetFieldNameValue("AMT", strAMT), _
                        GetFieldNameValue("STDATE", strSTDATE), _
                        GetFieldNameValue("EDDATE", strEDDATE), _
                        GetFieldNameValue("CONTRACTDAY", strCONTRACTDAY), _
                        GetFieldNameValue("CONFIRMFLAG", strCONFIRMFLAG), _
                        GetFieldNameValue("CONFIRMDATE", strCONFIRMDATE), _
                        GetFieldNameValue("CONFIRM_USER", strCONFIRM_USER), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("CONDITION", strCONDITION), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("SEQ", strSEQ)))

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
    Public Function DeleteDo(Optional ByVal strSEQ As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("SEQ", strSEQ)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function Update_ConfOK(Optional ByVal strSEQNO As String = OPTIONAL_STR, _
                                  Optional ByVal strCONTRACTNO As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
            strNOW = Now

            strSQL = "  UPDATE SC_CONTRACT_MST"
            strSQL = strSQL & "  SET CONFIRMFLAG = 'Y',"
            strSQL = strSQL & "  CONFIRM_USER = '" & mobjSCGLConfig.WRKUSR & "', "
            strSQL = strSQL & "  CONFIRMDATE =  '" & strNOW & "', "
            strSQL = strSQL & "  CONTRACTNO = '" & strCONTRACTNO & "'"
            strSQL = strSQL & "  WHERE SEQ = '" & strSEQNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function Update_ConfCAN(Optional ByVal strSEQNO As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
            strNOW = Now

            strSQL = "  UPDATE SC_CONTRACT_MST"
            strSQL = strSQL & "  SET CONFIRMFLAG = 'N',"
            strSQL = strSQL & "  CONFIRM_USER = '', "
            strSQL = strSQL & "  CONFIRMDATE =  '', "
            strSQL = strSQL & "  CONTRACTNO = '' "
            strSQL = strSQL & "  WHERE SEQ = '" & strSEQNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
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
        MyBase.EntityName = "SC_CONTRACT_MST"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region

End Class