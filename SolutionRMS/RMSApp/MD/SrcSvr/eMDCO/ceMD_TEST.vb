'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : cePD_GROUP_DIVAMT.vb ( PD_GROUP_DIVAMT Entity 처리 Class)
'기      능 : PD_GROUP_DIVAMT Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2007-11-12 오후 8:21:21 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceMD_TEST
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_TEST"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(ByVal dblSEQ As Double, _
            ByVal strNAME As String, _
            ByVal strATTR01 As String)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder

        Try
            BuildNameValues(",", "SEQ", dblSEQ, strFields, strValues)
            BuildNameValues(",", "NAME", strNAME, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)


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
    'Public Function UpdateDo(Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
    '        Optional ByVal strWOCODE As String = OPTIONAL_STR, _
    '        Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
    '        Optional ByVal strYEARMON As String = OPTIONAL_STR, _
    '        Optional ByVal strDEMANDYEARMON As String = OPTIONAL_STR, _
    '        Optional ByVal dblDIVAMT As Double = OPTIONAL_NUM, _
    '        Optional ByVal dblADJAMT As Double = OPTIONAL_NUM, _
    '        Optional ByVal strATTR01 As String = OPTIONAL_STR, _
    '        Optional ByVal strATTR02 As String = OPTIONAL_STR, _
    '        Optional ByVal strATTR03 As String = OPTIONAL_STR, _
    '        Optional ByVal strATTR04 As String = OPTIONAL_STR, _
    '        Optional ByVal strATTR05 As String = OPTIONAL_STR, _
    '        Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
    '        Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
    '        Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
    '        Optional ByVal dblATTR09 As Double = OPTIONAL_NUM, _
    '        Optional ByVal dblATTR10 As Double = OPTIONAL_NUM) As Integer

    '    Dim strSQL As String

    '    Try
    '        strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
    '                 BuildFields(",", _
    '                    GetFieldNameValue("YEARMON", strDEMANDYEARMON), _
    '                    GetFieldNameValue("CUSTCODE", strCUSTCODE), _
    '                    GetFieldNameValue("DIVAMT", dblDIVAMT), _
    '                    GetFieldNameValue("ADJAMT", dblADJAMT), _
    '                    GetFieldNameValue("ATTR01", strATTR01), _
    '                    GetFieldNameValue("ATTR02", strATTR02), _
    '                    GetFieldNameValue("ATTR03", strATTR03), _
    '                    GetFieldNameValue("ATTR04", strATTR04), _
    '                    GetFieldNameValue("ATTR05", strATTR05), _
    '                    GetFieldNameValue("ATTR06", dblATTR06), _
    '                    GetFieldNameValue("ATTR07", dblATTR07), _
    '                    GetFieldNameValue("ATTR08", dblATTR08), _
    '                    GetFieldNameValue("ATTR09", dblATTR09), _
    '                    GetFieldNameValue("ATTR10", dblATTR10), _
    '                    GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
    '                    GetFieldNameValue("UDATE", Now, False)), _
    '                 BuildFields("AND", _
    '                    GetFieldNameValue("SEQ", dblSEQ), GetFieldNameValue("WOCODE", strWOCODE), GetFieldNameValue("TRIM(YEARMON)", strYEARMON)))

    '        Return ProcEntity(strSQL)
    '    Catch err As Exception
    '        Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo")
    '    End Try
    'End Function
    ''DeleteUpdateDo
    ''intJOBNOSEQ, _
    ''strJOBNO, _
    ''strCUSTCODE, _
    ''lngADJAMT)
    'Public Function DeleteUpdateDo(Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
    '        Optional ByVal strWOCODE As String = OPTIONAL_STR, _
    '        Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
    '        Optional ByVal dblADJAMT As Double = OPTIONAL_NUM) As Integer

    '    Dim strSQL As String

    '    Try
    '        strSQL = "UPDATE PD_GROUP_DIVAMT SET  ADJAMT =  ADJAMT - " & dblADJAMT & " WHERE SEQ = " & dblSEQ & " AND WOCODE = '" & strWOCODE & "'"
    '        strSQL = strSQL & " AND CUSTCODE = '" & strCUSTCODE & "'"

    '        Return ProcEntity(strSQL)
    '    Catch err As Exception
    '        Throw RaiseSysErr(err, CLASS_NAME & ".DeleteUpdateDo")
    '    End Try
    'End Function
    ''*****************************************************************
    ''입력 : strSQL = SQL 문
    ''반환 : 처리건수
    ''기능 : 해당 Entity에 Delete 처리
    ''참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    ''*****************************************************************
    'Public Function DeleteDo(Optional ByVal dblSEQ As Double = OPTIONAL_NUM, Optional ByVal strWOCODE As String = OPTIONAL_STR, Optional ByVal strCUSTCODE As String = OPTIONAL_STR) As Integer
    '    Dim strSQL As String

    '    Try
    '        strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
    '                 BuildFields("AND", _
    '                               GetFieldNameValue("SEQ", dblSEQ), GetFieldNameValue("WOCODE", strWOCODE), GetFieldNameValue("CUSTCODE", strCUSTCODE)))

    '        Return ProcEntity(strSQL)
    '    Catch err As Exception
    '        Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
    '    End Try
    'End Function

    ''*****************************************************************
    ''입력 : strSQL = SQL 문
    ''반환 : 처리건수
    ''기능 : 해당 Entity에 Select 처리
    ''*****************************************************************
    'Public Function SelectDo(ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
    '                        Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
    '                        Optional ByVal strWOCODE As String = OPTIONAL_STR, _
    '                        Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
    '                        Optional ByVal strSelFields As String = "*", _
    '                        Optional ByVal intLimitRow As Integer = 0, _
    '                        Optional ByVal intSelMode As Integer = SELMODE.ARR, _
    '                        Optional ByVal blnBindingHeader As Boolean = False) As Object
    '    Dim strSQL As String
    '    Dim strKeyFields As String

    '    Try
    '        strKeyFields = BuildFields("AND", _
    '                                GetFieldNameValue("SEQ", dblSEQ), GetFieldNameValue("WOCODE", strWOCODE), GetFieldNameValue("CUSTCODE", strCUSTCODE))

    '        Return SelectDoExt(intRowCnt, intColCnt, strSelFields, strKeyFields, intLimitRow, intSelMode, blnBindingHeader)
    '    Catch err As Exception
    '        Throw RaiseSysErr(err, CLASS_NAME & ".SeleteDo")
    '    End Try
    'End Function
    ''*****************************************************************
    ''입력 : strSQL = SQL 문
    ''반환 : 처리건수
    ''기능 : 해당 Entity에 정산금액 업데이트 처리
    ''*****************************************************************
    'Public Function ADJAMTUpdateDo(Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
    '                               Optional ByVal strWOCODE As String = OPTIONAL_STR, _
    '                               Optional ByVal dblADJAMT As Double = OPTIONAL_NUM, _
    '                               Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
    '                               Optional ByVal strATTR02 As String = OPTIONAL_STR, _
    '                               Optional ByVal dblATTR07 As Double = OPTIONAL_NUM)

    '    Dim strSQL As String

    '    Try
    '        strSQL = " UPDATE PD_GROUP_DIVAMT SET ADJAMT = NVL(ADJAMT,0)+'" & dblADJAMT & "', "
    '        strSQL = strSQL & " ATTR06 = '" & dblATTR06 & "', "
    '        strSQL = strSQL & " ATTR02 = '" & strATTR02 & "',"
    '        strSQL = strSQL & " ATTR07 = '" & dblATTR07 & "' "
    '        strSQL = strSQL & " WHERE SEQ = '" & dblSEQ & "' AND WOCODE = '" & strWOCODE & "' "
    '        Return ProcEntity(strSQL)
    '    Catch err As Exception
    '        Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo")
    '    End Try
    'End Function
#End Region

#Region "객체 생성/해제"
    '*****************************************************************
    '입력 : strInfoXML = 공통기본정보에 대한 XML
    'objSCGLSql = DB 처리 객체 인스턴싱 변수    '반환 : 없음
    '기능 : DB 처리를 위한 공통기본정보 설정
    '*****************************************************************
    Public Sub New(Optional ByVal objSCGLConfig As SCGLUtil.cbSCGLConfig = Nothing, Optional ByVal strInfoXML As String = "")
        MyBase.SetConfig(objSCGLConfig, strInfoXML)
        MyBase.EntityName = "MD_TEST"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class





