'****************************************************************************************
'Generated By: JNF BY Solution
'시스템구분 : SolutionOLD/MD/ceMD_BOOKING_MEDIUM Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceMD_BOOKING_MEDIUM.vb ( MD_BOOKING_MEDIUM Entity 처리 Class)
'기      능 : MD_BOOKING_MEDIUM Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-08-18 오후 2:40:17 By Kim Tae Ho
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceMD_WONGO_MST
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_WONGO_MST"    '자신의 클래스명
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
            ByVal dblITEM_SEQ As Double, _
            ByVal strPUB_DATE As String, _
            ByVal strCLIENTCODE As String, _
            ByVal strMEDCODE As String, _
            ByVal strREAL_MED_CODE As String, _
            ByVal strMED_FLAG As String, _
            Optional ByVal strPROGRAM_NAME As String = OPTIONAL_STR, _
            Optional ByVal strCOL_DEG As String = OPTIONAL_STR, _
            Optional ByVal dblSTD_CM As Double = OPTIONAL_NUM, _
            Optional ByVal dblSTD_STEP As Double = OPTIONAL_NUM, _
            Optional ByVal strEND_TIME As String = OPTIONAL_STR, _
            Optional ByVal strOLD_DATE As String = OPTIONAL_STR, _
            Optional ByVal strCRE_NAME As String = OPTIONAL_STR, _
            Optional ByVal strDELIVER_NAME As String = OPTIONAL_STR, _
            Optional ByVal strCONTACT_FLAG As String = OPTIONAL_STR, _
            Optional ByVal strGUBUN As String = OPTIONAL_STR, _
            Optional ByVal strOUTFLAG As String = OPTIONAL_STR, _
            Optional ByVal strNOTE As String = OPTIONAL_STR, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR, _
            Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR09 As String = OPTIONAL_STR, _
            Optional ByVal dblATTR10 As String = OPTIONAL_STR)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now

        Try
            BuildNameValues(",", "YEARMON", strYEARMON, strFields, strValues)
            BuildNameValues(",", "SEQ", dblSEQ, strFields, strValues)
            BuildNameValues(",", "ITEM_SEQ", dblITEM_SEQ, strFields, strValues)
            BuildNameValues(",", "PUB_DATE", strPUB_DATE, strFields, strValues)
            BuildNameValues(",", "CLIENTCODE", strCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "MEDCODE", strMEDCODE, strFields, strValues)
            BuildNameValues(",", "REAL_MED_CODE", strREAL_MED_CODE, strFields, strValues)
            BuildNameValues(",", "MED_FLAG", strMED_FLAG, strFields, strValues)
            BuildNameValues(",", "PROGRAM_NAME", strPROGRAM_NAME, strFields, strValues)
            BuildNameValues(",", "COL_DEG", strCOL_DEG, strFields, strValues)
            BuildNameValues(",", "STD_CM", dblSTD_CM, strFields, strValues)
            BuildNameValues(",", "STD_STEP", dblSTD_STEP, strFields, strValues)
            BuildNameValues(",", "END_TIME", strEND_TIME, strFields, strValues)
            BuildNameValues(",", "OLD_DATE", strOLD_DATE, strFields, strValues)
            BuildNameValues(",", "CRE_NAME", strCRE_NAME, strFields, strValues)
            BuildNameValues(",", "DELIVER_NAME", strDELIVER_NAME, strFields, strValues)
            BuildNameValues(",", "CONTACT_FLAG", strCONTACT_FLAG, strFields, strValues)
            BuildNameValues(",", "GUBUN", strGUBUN, strFields, strValues)
            BuildNameValues(",", "OUTFLAG", strOUTFLAG, strFields, strValues)
            BuildNameValues(",", "NOTE", strNOTE, strFields, strValues)
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
    Public Function UpdateDo(Optional ByVal strYEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal dblITEM_SEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strPUB_DATE As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strMEDCODE As String = OPTIONAL_STR, _
            Optional ByVal strREAL_MED_CODE As String = OPTIONAL_STR, _
            Optional ByVal strMED_FLAG As String = OPTIONAL_STR, _
            Optional ByVal strPROGRAM_NAME As String = OPTIONAL_STR, _
            Optional ByVal strCOL_DEG As String = OPTIONAL_STR, _
            Optional ByVal dblSTD_CM As Double = OPTIONAL_NUM, _
            Optional ByVal dblSTD_STEP As Double = OPTIONAL_NUM, _
            Optional ByVal strEND_TIME As String = OPTIONAL_STR, _
            Optional ByVal strOLD_DATE As String = OPTIONAL_STR, _
            Optional ByVal strCRE_NAME As String = OPTIONAL_STR, _
            Optional ByVal strDELIVER_NAME As String = OPTIONAL_STR, _
            Optional ByVal strCONTACT_FLAG As String = OPTIONAL_STR, _
            Optional ByVal strGUBUN As String = OPTIONAL_STR, _
            Optional ByVal strOUTFLAG As String = OPTIONAL_STR, _
            Optional ByVal strNOTE As String = OPTIONAL_STR, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR, _
            Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR09 As String = OPTIONAL_STR, _
            Optional ByVal dblATTR10 As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("PUB_DATE", strPUB_DATE), _
                        GetFieldNameValue("CLIENTCODE", strCLIENTCODE), _
                        GetFieldNameValue("MEDCODE", strMEDCODE), _
                        GetFieldNameValue("REAL_MED_CODE", strREAL_MED_CODE), _
                        GetFieldNameValue("MED_FLAG", strMED_FLAG), _
                        GetFieldNameValue("PROGRAM_NAME", strPROGRAM_NAME), _
                        GetFieldNameValue("COL_DEG", strCOL_DEG), _
                        GetFieldNameValue("STD_CM", dblSTD_CM), _
                        GetFieldNameValue("STD_STEP", dblSTD_STEP), _
                        GetFieldNameValue("END_TIME", strEND_TIME), _
                        GetFieldNameValue("OLD_DATE", strOLD_DATE), _
                        GetFieldNameValue("CRE_NAME", strCRE_NAME), _
                        GetFieldNameValue("DELIVER_NAME", strDELIVER_NAME), _
                        GetFieldNameValue("CONTACT_FLAG", strCONTACT_FLAG), _
                        GetFieldNameValue("GUBUN", strGUBUN), _
                        GetFieldNameValue("OUTFLAG", strOUTFLAG), _
                        GetFieldNameValue("NOTE", strNOTE), _
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
                        GetFieldNameValue("YEARMON", strYEARMON), GetFieldNameValue("SEQ", dblSEQ), GetFieldNameValue("ITEM_SEQ", dblITEM_SEQ)))

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
    Public Function DeleteDo(Optional ByVal strYEARMON As String = OPTIONAL_STR, Optional ByVal dblSEQ As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("YEARMON", strYEARMON), GetFieldNameValue("SEQ", dblSEQ)))

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
                            Optional ByVal strYEARMON As String = OPTIONAL_STR, _
Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
                            Optional ByVal strSelFields As String = "*", _
                            Optional ByVal intLimitRow As Integer = 0, _
                            Optional ByVal intSelMode As Integer = SELMODE.ARR, _
                            Optional ByVal blnBindingHeader As Boolean = False) As Object
        Dim strSQL As String
        Dim strKeyFields As String

        Try
            strKeyFields = BuildFields("AND", _
                                    GetFieldNameValue("YEARMON", strYEARMON), GetFieldNameValue("SEQ", dblSEQ))

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
        MyBase.EntityName = "MD_WONGO_MST"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class

'------->>엔티티 INSERT/UPDATE 샘플입니다. 반드시 자신의 환경에 맞추어서 변경하시기 바랍니다.
'=========================================================
'       'vntData Array를 사용할 때 Insert/Update 입니다.
'=========================================================
'        Dim intRtn As Integer
'        intRtn = mobjceMD_BOOKING_MEDIUM.InsertDo( _
'                                       GetElement(vntData,"YEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"PUB_DATE", intColCnt, intRow), _
'                                       GetElement(vntData,"CLIENTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MEDCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"REAL_MED_CODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MED_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"SUBSEQ", intColCnt, intRow), _
'                                       GetElement(vntData,"DEPT_CD", intColCnt, intRow), _
'                                       GetElement(vntData,"PROGRAM_NAME", intColCnt, intRow), _
'                                       GetElement(vntData,"COL_DEG", intColCnt, intRow), _
'                                       GetElement(vntData,"PUB_FACE", intColCnt, intRow), _
'                                       GetElement(vntData,"STD_CM", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"STD_STEP", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"PRICE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"AMOUNT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"COMMI_RATE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"COMMISSION", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"GFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_TAX_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_TAX_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_TRANS_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_TAX_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_VOCH_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_TRANS_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_TAX_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_VOCH_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"NOTE", intColCnt, intRow), _
'                                       GetElement(vntData,"PROJECTION", intColCnt, intRow), _
'                                       GetElement(vntData,"SPONSOR", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR01", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR02", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR03", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR04", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR05", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR06", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR07", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR08", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR09", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR10", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CUSER", intColCnt, intRow), _
'                                       GetElement(vntData,"CDATE", intColCnt, intRow, NULL_DTM, true ), _
'                                       GetElement(vntData,"UUSER", intColCnt, intRow), _
'                                       GetElement(vntData,"UDATE", intColCnt, intRow, NULL_DTM, true ) _
'                                       )
'        Return intRtn

'        Dim intRtn As Integer
'        intRtn = mobjceMD_BOOKING_MEDIUM.UpdateDo( _
'                                       GetElement(vntData,"YEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"PUB_DATE", intColCnt, intRow), _
'                                       GetElement(vntData,"CLIENTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MEDCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"REAL_MED_CODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MED_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"SUBSEQ", intColCnt, intRow), _
'                                       GetElement(vntData,"DEPT_CD", intColCnt, intRow), _
'                                       GetElement(vntData,"PROGRAM_NAME", intColCnt, intRow), _
'                                       GetElement(vntData,"COL_DEG", intColCnt, intRow), _
'                                       GetElement(vntData,"PUB_FACE", intColCnt, intRow), _
'                                       GetElement(vntData,"STD_CM", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"STD_STEP", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"PRICE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"AMOUNT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"COMMI_RATE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"COMMISSION", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"GFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_TAX_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_TAX_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_TRANS_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_TAX_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_VOCH_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_TRANS_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_TAX_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_VOCH_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"NOTE", intColCnt, intRow), _
'                                       GetElement(vntData,"PROJECTION", intColCnt, intRow), _
'                                       GetElement(vntData,"SPONSOR", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR01", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR02", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR03", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR04", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR05", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR06", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR07", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR08", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR09", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR10", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CUSER", intColCnt, intRow), _
'                                       GetElement(vntData,"CDATE", intColCnt, intRow, NULL_DTM, true ), _
'                                       GetElement(vntData,"UUSER", intColCnt, intRow), _
'                                       GetElement(vntData,"UDATE", intColCnt, intRow, NULL_DTM, true ) _
'                                       )
'        Return intRtn


'=========================================================
'       'XmlData 를 사용할 때 Insert/Update 입니다.
'=========================================================
'        Dim intRtn As Integer
'        intRtn = mobjceMD_BOOKING_MEDIUM.InsertDo( _
'                                       XMLGetElement(xmlRoot,"YEARMON"), _
'                                       XMLGetElement(xmlRoot,"SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"PUB_DATE"), _
'                                       XMLGetElement(xmlRoot,"CLIENTCODE"), _
'                                       XMLGetElement(xmlRoot,"MEDCODE"), _
'                                       XMLGetElement(xmlRoot,"REAL_MED_CODE"), _
'                                       XMLGetElement(xmlRoot,"MED_FLAG"), _
'                                       XMLGetElement(xmlRoot,"SUBSEQ"), _
'                                       XMLGetElement(xmlRoot,"DEPT_CD"), _
'                                       XMLGetElement(xmlRoot,"PROGRAM_NAME"), _
'                                       XMLGetElement(xmlRoot,"COL_DEG"), _
'                                       XMLGetElement(xmlRoot,"PUB_FACE"), _
'                                       XMLGetElement(xmlRoot,"STD_CM", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"STD_STEP", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"PRICE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"AMOUNT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"COMMI_RATE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"COMMISSION", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"GFLAG"), _
'                                       XMLGetElement(xmlRoot,"TRU_TAX_FLAG"), _
'                                       XMLGetElement(xmlRoot,"COMMI_TAX_FLAG"), _
'                                       XMLGetElement(xmlRoot,"TRU_TRANS_NO"), _
'                                       XMLGetElement(xmlRoot,"TRU_TAX_NO"), _
'                                       XMLGetElement(xmlRoot,"TRU_VOCH_NO"), _
'                                       XMLGetElement(xmlRoot,"COMMI_TRANS_NO"), _
'                                       XMLGetElement(xmlRoot,"COMMI_TAX_NO"), _
'                                       XMLGetElement(xmlRoot,"COMMI_VOCH_NO"), _
'                                       XMLGetElement(xmlRoot,"NOTE"), _
'                                       XMLGetElement(xmlRoot,"PROJECTION"), _
'                                       XMLGetElement(xmlRoot,"SPONSOR"), _
'                                       XMLGetElement(xmlRoot,"ATTR01"), _
'                                       XMLGetElement(xmlRoot,"ATTR02"), _
'                                       XMLGetElement(xmlRoot,"ATTR03"), _
'                                       XMLGetElement(xmlRoot,"ATTR04"), _
'                                       XMLGetElement(xmlRoot,"ATTR05"), _
'                                       XMLGetElement(xmlRoot,"ATTR06", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR07", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR08", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR09", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR10", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CUSER"), _
'                                       XMLGetElement(xmlRoot,"CDATE", NULL_DTM, true ), _
'                                       XMLGetElement(xmlRoot,"UUSER"), _
'                                       XMLGetElement(xmlRoot,"UDATE", NULL_DTM, true ) _
'                                       )
'        Return intRtn

'        Dim intRtn As Integer
'        intRtn = mobjceMD_BOOKING_MEDIUM.UpdateDo( _
'                                       XMLGetElement(xmlRoot,"YEARMON"), _
'                                       XMLGetElement(xmlRoot,"SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"PUB_DATE"), _
'                                       XMLGetElement(xmlRoot,"CLIENTCODE"), _
'                                       XMLGetElement(xmlRoot,"MEDCODE"), _
'                                       XMLGetElement(xmlRoot,"REAL_MED_CODE"), _
'                                       XMLGetElement(xmlRoot,"MED_FLAG"), _
'                                       XMLGetElement(xmlRoot,"SUBSEQ"), _
'                                       XMLGetElement(xmlRoot,"DEPT_CD"), _
'                                       XMLGetElement(xmlRoot,"PROGRAM_NAME"), _
'                                       XMLGetElement(xmlRoot,"COL_DEG"), _
'                                       XMLGetElement(xmlRoot,"PUB_FACE"), _
'                                       XMLGetElement(xmlRoot,"STD_CM", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"STD_STEP", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"PRICE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"AMOUNT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"COMMI_RATE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"COMMISSION", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"GFLAG"), _
'                                       XMLGetElement(xmlRoot,"TRU_TAX_FLAG"), _
'                                       XMLGetElement(xmlRoot,"COMMI_TAX_FLAG"), _
'                                       XMLGetElement(xmlRoot,"TRU_TRANS_NO"), _
'                                       XMLGetElement(xmlRoot,"TRU_TAX_NO"), _
'                                       XMLGetElement(xmlRoot,"TRU_VOCH_NO"), _
'                                       XMLGetElement(xmlRoot,"COMMI_TRANS_NO"), _
'                                       XMLGetElement(xmlRoot,"COMMI_TAX_NO"), _
'                                       XMLGetElement(xmlRoot,"COMMI_VOCH_NO"), _
'                                       XMLGetElement(xmlRoot,"NOTE"), _
'                                       XMLGetElement(xmlRoot,"PROJECTION"), _
'                                       XMLGetElement(xmlRoot,"SPONSOR"), _
'                                       XMLGetElement(xmlRoot,"ATTR01"), _
'                                       XMLGetElement(xmlRoot,"ATTR02"), _
'                                       XMLGetElement(xmlRoot,"ATTR03"), _
'                                       XMLGetElement(xmlRoot,"ATTR04"), _
'                                       XMLGetElement(xmlRoot,"ATTR05"), _
'                                       XMLGetElement(xmlRoot,"ATTR06", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR07", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR08", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR09", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR10", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CUSER"), _
'                                       XMLGetElement(xmlRoot,"CDATE", NULL_DTM, true ), _
'                                       XMLGetElement(xmlRoot,"UUSER"), _
'                                       XMLGetElement(xmlRoot,"UDATE", NULL_DTM, true ) _
'                                       )
'        Return intRtn


