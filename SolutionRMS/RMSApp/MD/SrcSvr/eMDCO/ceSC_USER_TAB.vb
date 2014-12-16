'****************************************************************************************
'Generated By :::BELLWORLD Making Entity Bean Ver1.0:::
'시스템구분 : SFAR/??/Server단 Entity Base Component
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceSC_USER_TAB.vb ( SC_USER_TAB Entity 처리 Class)
'기      능 : SC_USER_TAB Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003-07-01 오후 5:42:33 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceSC_USER_TAB
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceSC_USER_TAB"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(ByRef dblTAB_ID As Double, _
            Optional ByVal strSC_BU_CODE As String = OPTIONAL_STR, _
            Optional ByVal strTAB_NAME As String = OPTIONAL_STR, _
            Optional ByVal strTAB_USER_NAME As String = OPTIONAL_STR, _
            Optional ByVal strTAB_TYPE As String = OPTIONAL_STR, _
            Optional ByVal strTAB_DESC As String = OPTIONAL_STR)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder()
        Dim strValues As New System.Text.StringBuilder()

        Try
            'Sequence ID 자동생성
            'dblTAB_ID = MyBase.GetNextSeqID()

            BuildNameValues(",", "SC_BU_CODE", strSC_BU_CODE, strFields, strValues)
            BuildNameValues(",", "TAB_ID", dblTAB_ID, strFields, strValues)
            BuildNameValues(",", "TAB_NAME", strTAB_NAME, strFields, strValues)
            BuildNameValues(",", "TAB_USER_NAME", strTAB_USER_NAME, strFields, strValues)
            BuildNameValues(",", "TAB_TYPE", strTAB_TYPE, strFields, strValues)
            BuildNameValues(",", "TAB_DESC", strTAB_DESC, strFields, strValues)

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
    Public Function UpdateDo(Optional ByVal strSC_BU_CODE As String = OPTIONAL_STR, _
            Optional ByVal dblTAB_ID As Double = OPTIONAL_NUM, _
            Optional ByVal strTAB_NAME As String = OPTIONAL_STR, _
            Optional ByVal strTAB_USER_NAME As String = OPTIONAL_STR, _
            Optional ByVal strTAB_TYPE As String = OPTIONAL_STR, _
            Optional ByVal strTAB_DESC As String = OPTIONAL_STR) As Integer

        Dim strSQL As String

        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("TAB_NAME", strTAB_NAME), _
                        GetFieldNameValue("TAB_USER_NAME", strTAB_USER_NAME), _
                        GetFieldNameValue("TAB_TYPE", strTAB_TYPE), _
                        GetFieldNameValue("TAB_DESC", strTAB_DESC)), _
                     BuildFields("AND", _
                        GetFieldNameValue("SC_BU_CODE", strSC_BU_CODE), GetFieldNameValue("TAB_ID", dblTAB_ID)))

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
    Public Function DeleteDo(Optional ByVal strSC_BU_CODE As String = OPTIONAL_STR, Optional ByVal dblTAB_ID As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("SC_BU_CODE", strSC_BU_CODE), GetFieldNameValue("TAB_ID", dblTAB_ID)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Select 처리
    '*****************************************************************
    Public Function SelectDo(ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                            Optional ByVal strSC_BU_CODE As String = OPTIONAL_STR, _
                            Optional ByVal dblTAB_ID As Double = OPTIONAL_NUM, _
                            Optional ByVal strSelFields As String = "*", _
                            Optional ByVal intLimitRow As Integer = 0, _
                            Optional ByVal intSelMode As Integer = SELMODE.ARR, _
                            Optional ByVal blnBindingHeader As Boolean = False) As Object
        Dim strSQL As String
        Dim strKeyFields As String

        Try
            strKeyFields = BuildFields("AND", _
                                    GetFieldNameValue("SC_BU_CODE", strSC_BU_CODE), GetFieldNameValue("TAB_ID", dblTAB_ID))

            Return SelectDoExt(intRowCnt, intColCnt, strSelFields, strKeyFields, intLimitRow, intSelMode, blnBindingHeader)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".SeleteDo")
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
        MyBase.EntityName = "SC_USER_TAB"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class

