'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceMD_INTERNET_MEDIUM.vb ( MD_INTERNET_MEDIUM Entity 처리 Class)
'기      능 : MD_INTERNET_MEDIUM Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-01-19 오후 7:19:00 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceMD_IFCMALL_SUSU
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_IFCMALL_SUSU"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(ByVal strYEARMON As String, _
                             ByVal dblSEQ As Double, _
                             Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
                             Optional ByVal strCONT_CODE As String = OPTIONAL_STR, _
                             Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
                             Optional ByVal dblRATE As Double = OPTIONAL_NUM, _
                             Optional ByVal dblSUSU_AMT As Double = OPTIONAL_NUM, _
                             Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strVOCHNO As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRM_USER As String = OPTIONAL_STR, _
                             Optional ByVal strCONFIRM_DATE As String = OPTIONAL_STR, _
                             Optional ByVal strMEMO As String = OPTIONAL_STR)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now

        Try
            BuildNameValues(",", "YEARMON", strYEARMON, strFields, strValues)
            BuildNameValues(",", "SEQ", dblSEQ, strFields, strValues)
            BuildNameValues(",", "DEMANDDAY", strDEMANDDAY, strFields, strValues)
            BuildNameValues(",", "CONT_CODE", strCONT_CODE, strFields, strValues)
            BuildNameValues(",", "DEPT_CD", strDEPT_CD, strFields, strValues)
            BuildNameValues(",", "RATE", dblRATE, strFields, strValues)
            BuildNameValues(",", "SUSU_AMT", dblSUSU_AMT, strFields, strValues)
            BuildNameValues(",", "EXCLIENTCODE", strEXCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "VOCHNO", strVOCHNO, strFields, strValues)
            BuildNameValues(",", "CONFIRM_USER", strCONFIRM_USER, strFields, strValues)
            BuildNameValues(",", "CONFIRM_DATE", strCONFIRM_DATE, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
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
    Public Function UpdateDo(ByVal strYEARMON As String, _
                             ByVal dblSEQ As Double, _
                             Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
                             Optional ByVal strCONT_CODE As String = OPTIONAL_STR, _
                             Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
                             Optional ByVal dblRATE As Double = OPTIONAL_NUM, _
                             Optional ByVal dblSUSU_AMT As Double = OPTIONAL_NUM, _
                             Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strMEMO As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now

        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("DEMANDDAY", strDEMANDDAY), _
                        GetFieldNameValue("DEPT_CD", strDEPT_CD), _
                        GetFieldNameValue("RATE", dblRATE), _
                        GetFieldNameValue("SUSU_AMT", dblSUSU_AMT), _
                        GetFieldNameValue("EXCLIENTCODE", strEXCLIENTCODE), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("YEARMON", strYEARMON), GetFieldNameValue("SEQ", dblSEQ), GetFieldNameValue("CONT_CODE", strCONT_CODE)))

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
    Public Function DeleteDo(Optional ByVal strYEARMON As String = OPTIONAL_STR, _
                             Optional ByVal dblSEQ As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String
        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("YEARMON", strYEARMON), _
                                   GetFieldNameValue("SEQ", dblSEQ)))

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
    Public Function DeleteDo_ALL(Optional ByVal strYEARMON As String = OPTIONAL_STR, _
                                 Optional ByVal strCONT_CODE As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("YEARMON", strYEARMON), _
                                   GetFieldNameValue("CONT_CODE", strCONT_CODE)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문 By OSH
    '반환 : 처리건수
    '기능 : 청약 데이터를 확정한다.
    '*****************************************************************
    Public Function Update_confirm(ByVal strYEARMON As String, _
                                   ByVal dblSEQ As Double, _
                                   ByVal strCONT_CODE As String, _
                                   ByVal strUSER As String) As Integer
        Dim strSQL As String
        Dim strNOW As String
        strNOW = Now

        Try
            strSQL = "  UPDATE MD_IFCMALL_SUSU"
            strSQL = strSQL & "  SET CONFIRM_USER = '" & strUSER & "', CONFIRM_DATE = '" & strNOW & "'"
            strSQL = strSQL & "  WHERE YEARMON = '" & strYEARMON & "'"
            strSQL = strSQL & "  AND SEQ = " & dblSEQ
            strSQL = strSQL & "  AND CONT_CODE = '" & strCONT_CODE & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_confirm")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문 By OSH
    '반환 : 처리건수
    '기능 : 청약 데이터를 확정취소한다.
    '*****************************************************************
    Public Function Update_confirmcancel(ByVal strYEARMON As String, _
                                         ByVal dblSEQ As Double, _
                                         ByVal strCONT_CODE As String) As Integer
        Dim strSQL As String
        Try
            strSQL = "  UPDATE MD_IFCMALL_SUSU"
            strSQL = strSQL & "  SET CONFIRM_USER = '', CONFIRM_DATE = ''"
            strSQL = strSQL & "  WHERE YEARMON = '" & strYEARMON & "'"
            strSQL = strSQL & "  AND SEQ = " & dblSEQ
            strSQL = strSQL & "  AND CONT_CODE = '" & strCONT_CODE & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_confirmcancel")
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
        MyBase.EntityName = "MD_IFCMALL_SUSU"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class