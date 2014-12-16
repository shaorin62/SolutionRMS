'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : cePD_TAX_MST_TEMP.vb ( PD_TAX_MST_TEMP Entity 처리 Class)
'기      능 : PD_TAX_MST_TEMP Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-01-18 오후 5:57:23 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class cePD_TAX_MST_TEMP
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "cePD_TAX_MST_TEMP"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(ByVal strTAXYEARMON As String, _
            ByVal dblTAXNO As Double, _
            ByVal dblTAXSEQ As Double, _
            ByVal strTRANSYEARMON As String, _
            ByVal dblTRANSNO As Double, _
            ByVal dblSEQ As Double, _
            ByVal dblJOBNOSEQ As Double, _
            ByVal strJOBNO As String, _
            ByVal strJOBNAME As String, _
            ByVal strDEPTCD As String, _
            ByVal strCLIENTCODE As String, _
            ByVal strREAL_MED_CODE As String, _
            ByVal strDEMANDDAY As String, _
            ByVal strPRINTDAY As String, _
            Optional ByVal strVATFLAG As String = OPTIONAL_STR, _
            Optional ByVal dblADJAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strJOBGUBUN As String = OPTIONAL_STR, _
            Optional ByVal strSUMM As String = OPTIONAL_STR, _
            Optional ByVal dblVAT As Double = OPTIONAL_NUM, _
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

        Try
            BuildNameValues(",", "TAXYEARMON", strTAXYEARMON, strFields, strValues)
            BuildNameValues(",", "TAXNO", dblTAXNO, strFields, strValues)
            BuildNameValues(",", "TAXSEQ", dblTAXSEQ, strFields, strValues)
            BuildNameValues(",", "TRANSYEARMON", strTRANSYEARMON, strFields, strValues)
            BuildNameValues(",", "TRANSNO", dblTRANSNO, strFields, strValues)
            BuildNameValues(",", "SEQ", dblSEQ, strFields, strValues)
            BuildNameValues(",", "JOBNOSEQ", dblJOBNOSEQ, strFields, strValues)
            BuildNameValues(",", "JOBNO", strJOBNO, strFields, strValues)
            BuildNameValues(",", "JOBNAME", strJOBNAME, strFields, strValues)
            BuildNameValues(",", "DEPTCD", strDEPTCD, strFields, strValues)
            BuildNameValues(",", "CLIENTCODE", strCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "REAL_MED_CODE", strREAL_MED_CODE, strFields, strValues)
            BuildNameValues(",", "DEMANDDAY", strDEMANDDAY, strFields, strValues)
            BuildNameValues(",", "PRINTDAY", strPRINTDAY, strFields, strValues)
            BuildNameValues(",", "VATFLAG", strVATFLAG, strFields, strValues)
            BuildNameValues(",", "ADJAMT", dblADJAMT, strFields, strValues)
            BuildNameValues(",", "JOBGUBUN", strJOBGUBUN, strFields, strValues)
            BuildNameValues(",", "SUMM", strSUMM, strFields, strValues)
            BuildNameValues(",", "VAT", dblVAT, strFields, strValues)
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
            BuildNameValues(",", "CDATE", NOW, strFields, strValues, False)
            BuildNameValues(",", "UUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "UDATE", NOW, strFields, strValues, False)

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
    Public Function UpdateDo(Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTAXNO As Double = OPTIONAL_NUM, _
            Optional ByVal dblTAXSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strTRANSYEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTRANSNO As Double = OPTIONAL_NUM, _
            Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal dblJOBNOSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strJOBNO As String = OPTIONAL_STR, _
            Optional ByVal strJOBNAME As String = OPTIONAL_STR, _
            Optional ByVal strDEPTCD As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strREAL_MED_CODE As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
            Optional ByVal strPRINTDAY As String = OPTIONAL_STR, _
            Optional ByVal strVATFLAG As String = OPTIONAL_STR, _
            Optional ByVal dblADJAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strJOBGUBUN As String = OPTIONAL_STR, _
            Optional ByVal strSUMM As String = OPTIONAL_STR, _
            Optional ByVal dblVAT As Double = OPTIONAL_NUM, _
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

        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("TRANSYEARMON", strTRANSYEARMON), _
                        GetFieldNameValue("TRANSNO", dblTRANSNO), _
                        GetFieldNameValue("SEQ", dblSEQ), _
                        GetFieldNameValue("JOBNOSEQ", dblJOBNOSEQ), _
                        GetFieldNameValue("JOBNO", strJOBNO), _
                        GetFieldNameValue("JOBNAME", strJOBNAME), _
                        GetFieldNameValue("DEPTCD", strDEPTCD), _
                        GetFieldNameValue("CLIENTCODE", strCLIENTCODE), _
                        GetFieldNameValue("REAL_MED_CODE", strREAL_MED_CODE), _
                        GetFieldNameValue("DEMANDDAY", strDEMANDDAY), _
                        GetFieldNameValue("PRINTDAY", strPRINTDAY), _
                        GetFieldNameValue("VATFLAG", strVATFLAG), _
                        GetFieldNameValue("ADJAMT", dblADJAMT), _
                        GetFieldNameValue("JOBGUBUN", strJOBGUBUN), _
                        GetFieldNameValue("SUMM", strSUMM), _
                        GetFieldNameValue("VAT", dblVAT), _
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
                        GetFieldNameValue("UDATE", NOW, False)), _
                     BuildFields("AND", _
                        GetFieldNameValue("TAXYEARMON", strTAXYEARMON), GetFieldNameValue("TAXNO", dblTAXNO), GetFieldNameValue("TAXSEQ", dblTAXSEQ)))

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
    Public Function DeleteDo(Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, Optional ByVal dblTAXNO As Double = OPTIONAL_NUM, Optional ByVal dblTAXSEQ As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("TAXYEARMON", strTAXYEARMON), GetFieldNameValue("TAXNO", dblTAXNO), GetFieldNameValue("TAXSEQ", dblTAXSEQ)))

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
                            Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
                            Optional ByVal dblTAXNO As Double = OPTIONAL_NUM, _
                            Optional ByVal dblTAXSEQ As Double = OPTIONAL_NUM, _
                            Optional ByVal strSelFields As String = "*", _
                            Optional ByVal intLimitRow As Integer = 0, _
                            Optional ByVal intSelMode As Integer = SELMODE.ARR, _
                            Optional ByVal blnBindingHeader As Boolean = False) As Object
        Dim strSQL As String
        Dim strKeyFields As String

        Try
            strKeyFields = BuildFields("AND", _
                                    GetFieldNameValue("TAXYEARMON", strTAXYEARMON), GetFieldNameValue("TAXNO", dblTAXNO), GetFieldNameValue("TAXSEQ", dblTAXSEQ))

            Return SelectDoExt(intRowCnt, intColCnt, strSelFields, strKeyFields, intLimitRow, intSelMode, blnBindingHeader)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".SeleteDo")
        End Try
    End Function
    Public Function TempDeleteDo()
        Dim strSQL As String
        Try
            strSQL = String.Format("DELETE FROM {0} ", EntityName)
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".TempDeleteDo")
        End Try
    End Function

    Public Function NotUseTempDeleteDo(ByVal strTRANSYEARMON As String, _
                                       ByVal intTRANSNO As Integer) As Integer
        Dim strSQL As String
        Try
            strSQL = "DELETE FROM PD_TAX_MST_TEMP WHERE TRANSYEARMON||TO_CHAR(TRANSNO) != '" & strTRANSYEARMON & intTRANSNO & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".TempDeleteDo")
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
        MyBase.EntityName = "PD_TAX_MST_TEMP"     'Entity Name 설정
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
'        intRtn = mobjcePD_TAX_MST_TEMP.InsertDo( _
'                                       GetElement(vntData,"TAXYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TAXNO", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TAXSEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TRANSYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TRANSNO", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"JOBNOSEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"JOBNO", intColCnt, intRow), _
'                                       GetElement(vntData,"JOBNAME", intColCnt, intRow), _
'                                       GetElement(vntData,"DEPTCD", intColCnt, intRow), _
'                                       GetElement(vntData,"CLIENTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"REAL_MED_CODE", intColCnt, intRow), _
'                                       GetElement(vntData,"DEMANDDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PRINTDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"VATFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"ADJAMT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"JOBGUBUN", intColCnt, intRow), _
'                                       GetElement(vntData,"SUMM", intColCnt, intRow), _
'                                       GetElement(vntData,"VAT", intColCnt, intRow, NULL_NUM, true ), _
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
'        intRtn = mobjcePD_TAX_MST_TEMP.UpdateDo( _
'                                       GetElement(vntData,"TAXYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TAXNO", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TAXSEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TRANSYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TRANSNO", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"JOBNOSEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"JOBNO", intColCnt, intRow), _
'                                       GetElement(vntData,"JOBNAME", intColCnt, intRow), _
'                                       GetElement(vntData,"DEPTCD", intColCnt, intRow), _
'                                       GetElement(vntData,"CLIENTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"REAL_MED_CODE", intColCnt, intRow), _
'                                       GetElement(vntData,"DEMANDDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PRINTDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"VATFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"ADJAMT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"JOBGUBUN", intColCnt, intRow), _
'                                       GetElement(vntData,"SUMM", intColCnt, intRow), _
'                                       GetElement(vntData,"VAT", intColCnt, intRow, NULL_NUM, true ), _
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
'        intRtn = mobjcePD_TAX_MST_TEMP.InsertDo( _
'                                       XMLGetElement(xmlRoot,"TAXYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TAXNO", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TAXSEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TRANSYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TRANSNO", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"JOBNOSEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"JOBNO"), _
'                                       XMLGetElement(xmlRoot,"JOBNAME"), _
'                                       XMLGetElement(xmlRoot,"DEPTCD"), _
'                                       XMLGetElement(xmlRoot,"CLIENTCODE"), _
'                                       XMLGetElement(xmlRoot,"REAL_MED_CODE"), _
'                                       XMLGetElement(xmlRoot,"DEMANDDAY"), _
'                                       XMLGetElement(xmlRoot,"PRINTDAY"), _
'                                       XMLGetElement(xmlRoot,"VATFLAG"), _
'                                       XMLGetElement(xmlRoot,"ADJAMT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"JOBGUBUN"), _
'                                       XMLGetElement(xmlRoot,"SUMM"), _
'                                       XMLGetElement(xmlRoot,"VAT", NULL_NUM, true ), _
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
'        intRtn = mobjcePD_TAX_MST_TEMP.UpdateDo( _
'                                       XMLGetElement(xmlRoot,"TAXYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TAXNO", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TAXSEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TRANSYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TRANSNO", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"JOBNOSEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"JOBNO"), _
'                                       XMLGetElement(xmlRoot,"JOBNAME"), _
'                                       XMLGetElement(xmlRoot,"DEPTCD"), _
'                                       XMLGetElement(xmlRoot,"CLIENTCODE"), _
'                                       XMLGetElement(xmlRoot,"REAL_MED_CODE"), _
'                                       XMLGetElement(xmlRoot,"DEMANDDAY"), _
'                                       XMLGetElement(xmlRoot,"PRINTDAY"), _
'                                       XMLGetElement(xmlRoot,"VATFLAG"), _
'                                       XMLGetElement(xmlRoot,"ADJAMT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"JOBGUBUN"), _
'                                       XMLGetElement(xmlRoot,"SUMM"), _
'                                       XMLGetElement(xmlRoot,"VAT", NULL_NUM, true ), _
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


