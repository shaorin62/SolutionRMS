'****************************************************************************************
'시스템구분 : RMS/PD/Server Entity Class
'실행  환경 : 
'프로그램명 : cePD_OUTLIST_MST.vb (PD_PD_OUTLIST_MST Entity 처리 Class)
'기      능 : PD_PD_OUTLIST_MST Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009-10-19 
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class cePD_OUTLIST_MST
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "cePD_OUTLIST_MST"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    'PREESTNO,JOBNO,PRODUCTIONNAME,DIRECTORNAME,EDIT,CG,TELECINE,RECORDING,CMSONG,STUDIO,MODELAGENCY,DATE,MEETINGDATE,SHOOTDATE,
    'DAYS,HOURS,TITLE,LENGTHS,LENGTHS

    Public Function InsertDo(ByVal strPREESTNO As String, _
            Optional ByVal strJOBNO As String = OPTIONAL_STR, _
            Optional ByVal strPRODUCTIONNAME As String = OPTIONAL_STR, _
            Optional ByVal strDIRECTORNAME As String = OPTIONAL_STR, _
            Optional ByVal strEDIT As String = OPTIONAL_STR, _
            Optional ByVal strCG As String = OPTIONAL_STR, _
            Optional ByVal strTELECINE As String = OPTIONAL_STR, _
            Optional ByVal strRECORDING As String = OPTIONAL_STR, _
            Optional ByVal strCMSONG As String = OPTIONAL_STR, _
            Optional ByVal strSTUDIO As String = OPTIONAL_STR, _
            Optional ByVal strMODELAGENCY As String = OPTIONAL_STR, _
            Optional ByVal strDATE As String = OPTIONAL_STR, _
            Optional ByVal strMEETINGDATE As String = OPTIONAL_STR, _
            Optional ByVal strSHOOTDATE As String = OPTIONAL_STR, _
            Optional ByVal strDAYS As String = OPTIONAL_STR, _
            Optional ByVal strHOURS As String = OPTIONAL_STR, _
            Optional ByVal strTITLE As String = OPTIONAL_STR, _
            Optional ByVal strLENGTHS As String = OPTIONAL_STR, _
            Optional ByVal strCOMMENTS As String = OPTIONAL_STR, _
            Optional ByVal strPRODUCT As String = OPTIONAL_STR, _
            Optional ByVal strPROJECT As String = OPTIONAL_STR, _
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
            BuildNameValues(",", "JOBNO", strJOBNO, strFields, strValues)
            BuildNameValues(",", "PRODUCTIONNAME", strPRODUCTIONNAME, strFields, strValues)
            BuildNameValues(",", "DIRECTORNAME", strDIRECTORNAME, strFields, strValues)
            BuildNameValues(",", "EDIT", strEDIT, strFields, strValues)
            BuildNameValues(",", "CG", strCG, strFields, strValues)
            BuildNameValues(",", "TELECINE", strTELECINE, strFields, strValues)
            BuildNameValues(",", "RECORDING", strRECORDING, strFields, strValues)
            BuildNameValues(",", "CMSONG", strCMSONG, strFields, strValues)
            BuildNameValues(",", "STUDIO", strSTUDIO, strFields, strValues)
            BuildNameValues(",", "MODELAGENCY", strMODELAGENCY, strFields, strValues)
            BuildNameValues(",", "DATE", strDATE, strFields, strValues)
            BuildNameValues(",", "MEETINGDATE", strMEETINGDATE, strFields, strValues)
            BuildNameValues(",", "SHOOTDATE", strSHOOTDATE, strFields, strValues)
            BuildNameValues(",", "DAYS", strDAYS, strFields, strValues)
            BuildNameValues(",", "HOURS", strHOURS, strFields, strValues)
            BuildNameValues(",", "TITLE", strTITLE, strFields, strValues)
            BuildNameValues(",", "LENGTHS", strLENGTHS, strFields, strValues)
            BuildNameValues(",", "COMMENTS", strCOMMENTS, strFields, strValues)
            BuildNameValues(",", "PRODUCT", strPRODUCT, strFields, strValues)
            BuildNameValues(",", "PROJECT", strPROJECT, strFields, strValues)
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
            Optional ByVal strJOBNO As String = OPTIONAL_STR, _
            Optional ByVal strPRODUCTIONNAME As String = OPTIONAL_STR, _
            Optional ByVal strDIRECTORNAME As String = OPTIONAL_STR, _
            Optional ByVal strEDIT As String = OPTIONAL_STR, _
            Optional ByVal strCG As String = OPTIONAL_STR, _
            Optional ByVal strTELECINE As String = OPTIONAL_STR, _
            Optional ByVal strRECORDING As String = OPTIONAL_STR, _
            Optional ByVal strCMSONG As String = OPTIONAL_STR, _
            Optional ByVal strSTUDIO As String = OPTIONAL_STR, _
            Optional ByVal strMODELAGENCY As String = OPTIONAL_STR, _
            Optional ByVal strDATE As String = OPTIONAL_STR, _
            Optional ByVal strMEETINGDATE As String = OPTIONAL_STR, _
            Optional ByVal strSHOOTDATE As String = OPTIONAL_STR, _
            Optional ByVal strDAYS As String = OPTIONAL_STR, _
            Optional ByVal strHOURS As String = OPTIONAL_STR, _
            Optional ByVal strTITLE As String = OPTIONAL_STR, _
            Optional ByVal strLENGTHS As String = OPTIONAL_STR, _
            Optional ByVal strCOMMENTS As String = OPTIONAL_STR, _
            Optional ByVal strPRODUCT As String = OPTIONAL_STR, _
            Optional ByVal strPROJECT As String = OPTIONAL_STR, _
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
                        GetFieldNameValue("JOBNO", strJOBNO), _
                        GetFieldNameValue("PRODUCTIONNAME", strPRODUCTIONNAME), _
                        GetFieldNameValue("DIRECTORNAME", strDIRECTORNAME), _
                        GetFieldNameValue("EDIT", strEDIT), _
                        GetFieldNameValue("CG", strCG), _
                        GetFieldNameValue("TELECINE", strTELECINE), _
                        GetFieldNameValue("RECORDING", strRECORDING), _
                        GetFieldNameValue("CMSONG", strCMSONG), _
                        GetFieldNameValue("STUDIO", strSTUDIO), _
                        GetFieldNameValue("MODELAGENCY", strMODELAGENCY), _
                        GetFieldNameValue("DATE", strDATE), _
                        GetFieldNameValue("MEETINGDATE", strMEETINGDATE), _
                        GetFieldNameValue("SHOOTDATE", strSHOOTDATE), _
                        GetFieldNameValue("DAYS", strDAYS), _
                        GetFieldNameValue("HOURS", strHOURS), _
                        GetFieldNameValue("TITLE", strTITLE), _
                        GetFieldNameValue("LENGTHS", strLENGTHS), _
                        GetFieldNameValue("COMMENTS", strCOMMENTS), _
                        GetFieldNameValue("PRODUCT", strPRODUCT), _
                        GetFieldNameValue("PROJECT", strPROJECT), _
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
                        GetFieldNameValue("PREESTNO", strPREESTNO)))

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
    'strPREESTNO,dblSEQ,dblSUBITEMCODESEQ,dblITEMCODESEQ,strITEMCODE
    Public Function DeleteDo(ByVal strPREESTNO As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "DELETE FROM PD_OUTLIST_MST WHERE PREESTNO = '" & strPREESTNO & "'"

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
        MyBase.EntityName = "PD_OUTLIST_MST"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class

