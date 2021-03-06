'****************************************************************************************
'Generated By: JNF KTY
'시스템구분 : Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceMD_MATTER_MST.vb ( ceMD_MATTER_MST Entity 처리 Class)
'기      능 : MD_MATTER_MST Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009-07-03 By KTY
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceMD_MATTER_MST
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_MATTER_MST"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(ByVal strMATTERCODE As String, _
                             Optional ByVal strMATTER As String = OPTIONAL_STR, _
                             Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strSUBSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strMEDFLAG As String = OPTIONAL_STR, _
                             Optional ByVal strMEMO As String = OPTIONAL_STR, _
                             Optional ByVal strATTR01 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR02 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR03 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR04 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR05 As String = OPTIONAL_STR, _
                             Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
                             Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
                             Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
                             Optional ByVal dblATTR09 As Double = OPTIONAL_NUM, _
                             Optional ByVal dblATTR10 As Double = OPTIONAL_NUM)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            BuildNameValues(",", "MATTERCODE", strMATTERCODE, strFields, strValues)
            BuildNameValues(",", "MATTER", strMATTER, strFields, strValues)
            BuildNameValues(",", "CUSTCODE", strCUSTCODE, strFields, strValues)
            BuildNameValues(",", "SUBSEQ", strSUBSEQ, strFields, strValues)
            BuildNameValues(",", "EXCLIENTCODE", strEXCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "MEDFLAG", strMEDFLAG, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
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
    Public Function UpdateDo(Optional ByVal strMATTERCODE As String = OPTIONAL_STR, _
                             Optional ByVal strMATTER As String = OPTIONAL_STR, _
                             Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strSUBSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strMEMO As String = OPTIONAL_STR, _
                             Optional ByVal strATTR02 As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("MATTER", strMATTER), _
                        GetFieldNameValue("CUSTCODE", strCUSTCODE), _
                        GetFieldNameValue("SUBSEQ", strSUBSEQ), _
                        GetFieldNameValue("EXCLIENTCODE", strEXCLIENTCODE), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("ATTR02", strATTR02), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("MATTERCODE", strMATTERCODE)))

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
    Public Function Update_Conf(Optional ByVal strMATTERCODE As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("UPDATE {0} SET ATTR02 = 'S' WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("MATTERCODE", strMATTERCODE)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_Conf")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function Update_ConfOK(Optional ByVal strMATTERCODE As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("UPDATE {0} SET ATTR02 = 'Y' WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("MATTERCODE", strMATTERCODE)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_ConfOK")
        End Try
    End Function

    Public Function DeleteDo(Optional ByVal strMATTERCODE As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("MATTERCODE", strMATTERCODE)))

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
        MyBase.EntityName = "MD_MATTER_MST"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class

