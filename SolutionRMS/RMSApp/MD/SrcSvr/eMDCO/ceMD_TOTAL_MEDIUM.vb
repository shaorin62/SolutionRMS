'****************************************************************************************
'Generated By: JNF BY Solution
'시스템구분 : SolutionOLD/MD/ceMD_TOTAL_MEDIUM Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceMD_TOTAL_MEDIUM.vb ( MD_TOTAL_MEDIUM Entity 처리 Class)
'기      능 : MD_TOTAL_MEDIUM Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-08-18 오후 2:40:17 By Kim Tae Ho
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceMD_TOTAL_MEDIUM
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_TOTAL_MEDIUM"    '자신의 클래스명
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
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTSUBCODE As String = OPTIONAL_STR, _
            Optional ByVal strTIMCODE As String = OPTIONAL_STR, _
            Optional ByVal strREAL_MED_CODE As String = OPTIONAL_STR, _
            Optional ByVal strMEDCODE As String = OPTIONAL_STR, _
            Optional ByVal strSUBSEQ As String = OPTIONAL_STR, _
            Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
            Optional ByVal strMATTERCODE As String = OPTIONAL_STR, _
            Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strGREATCODE As String = OPTIONAL_STR, _
            Optional ByVal strMPP As String = OPTIONAL_STR, _
            Optional ByVal strPROGRAM As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblCOMMI_RATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblCOMMISSION As Double = OPTIONAL_NUM, _
            Optional ByVal strTBRDSTDATE As String = OPTIONAL_STR, _
            Optional ByVal strTBRDEDDATE As String = OPTIONAL_STR, _
            Optional ByVal strCNT As String = OPTIONAL_STR, _
            Optional ByVal strGFLAG As String = OPTIONAL_STR, _
            Optional ByVal strVOCH_TYPE As String = OPTIONAL_STR, _
            Optional ByVal strTRU_TAX_FLAG As String = OPTIONAL_STR, _
            Optional ByVal strCOMMI_TAX_FLAG As String = OPTIONAL_STR, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
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
            BuildNameValues(",", "CLIENTCODE", strCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "CLIENTSUBCODE", strCLIENTSUBCODE, strFields, strValues)
            BuildNameValues(",", "TIMCODE", strTIMCODE, strFields, strValues)
            BuildNameValues(",", "REAL_MED_CODE", strREAL_MED_CODE, strFields, strValues)
            BuildNameValues(",", "MEDCODE", strMEDCODE, strFields, strValues)
            BuildNameValues(",", "SUBSEQ", strSUBSEQ, strFields, strValues)
            BuildNameValues(",", "DEPT_CD", strDEPT_CD, strFields, strValues)
            BuildNameValues(",", "MATTERCODE", strMATTERCODE, strFields, strValues)
            BuildNameValues(",", "EXCLIENTCODE", strEXCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "GREATCODE", strGREATCODE, strFields, strValues)
            BuildNameValues(",", "MPP", strMPP, strFields, strValues)
            BuildNameValues(",", "PROGRAM", strPROGRAM, strFields, strValues)
            BuildNameValues(",", "DEMANDDAY", strDEMANDDAY, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "COMMI_RATE", dblCOMMI_RATE, strFields, strValues)
            BuildNameValues(",", "COMMISSION", dblCOMMISSION, strFields, strValues)
            BuildNameValues(",", "TBRDSTDATE", strTBRDSTDATE, strFields, strValues)
            BuildNameValues(",", "TBRDEDDATE", strTBRDEDDATE, strFields, strValues)
            BuildNameValues(",", "CNT", strCNT, strFields, strValues)
            BuildNameValues(",", "GFLAG", strGFLAG, strFields, strValues)
            BuildNameValues(",", "VOCH_TYPE", strVOCH_TYPE, strFields, strValues)
            BuildNameValues(",", "TRU_TAX_FLAG", strTRU_TAX_FLAG, strFields, strValues)
            BuildNameValues(",", "COMMI_TAX_FLAG", strCOMMI_TAX_FLAG, strFields, strValues)
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
    Public Function UpdateDo(Optional ByVal strYEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTSUBCODE As String = OPTIONAL_STR, _
            Optional ByVal strTIMCODE As String = OPTIONAL_STR, _
            Optional ByVal strREAL_MED_CODE As String = OPTIONAL_STR, _
            Optional ByVal strMEDCODE As String = OPTIONAL_STR, _
            Optional ByVal strSUBSEQ As String = OPTIONAL_STR, _
            Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
            Optional ByVal strMATTERCODE As String = OPTIONAL_STR, _
            Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strGREATCODE As String = OPTIONAL_STR, _
            Optional ByVal strMPP As String = OPTIONAL_STR, _
            Optional ByVal strPROGRAM As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblCOMMI_RATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblCOMMISSION As Double = OPTIONAL_NUM, _
            Optional ByVal strTBRDSTDATE As String = OPTIONAL_STR, _
            Optional ByVal strTBRDEDDATE As String = OPTIONAL_STR, _
            Optional ByVal strCNT As String = OPTIONAL_STR, _
            Optional ByVal strGFLAG As String = OPTIONAL_STR, _
            Optional ByVal strVOCH_TYPE As String = OPTIONAL_STR, _
            Optional ByVal strTRU_TAX_FLAG As String = OPTIONAL_STR, _
            Optional ByVal strCOMMI_TAX_FLAG As String = OPTIONAL_STR, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
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
                        GetFieldNameValue("CLIENTCODE", strCLIENTCODE), _
                        GetFieldNameValue("CLIENTSUBCODE", strCLIENTSUBCODE), _
                        GetFieldNameValue("TIMCODE", strTIMCODE), _
                        GetFieldNameValue("REAL_MED_CODE", strREAL_MED_CODE), _
                        GetFieldNameValue("MEDCODE", strMEDCODE), _
                        GetFieldNameValue("SUBSEQ", strSUBSEQ), _
                        GetFieldNameValue("DEPT_CD", strDEPT_CD), _
                        GetFieldNameValue("MATTERCODE", strMATTERCODE), _
                        GetFieldNameValue("EXCLIENTCODE", strEXCLIENTCODE), _
                        GetFieldNameValue("GREATCODE", strGREATCODE), _
                        GetFieldNameValue("MPP", strMPP), _
                        GetFieldNameValue("PROGRAM", strPROGRAM), _
                        GetFieldNameValue("DEMANDDAY ", strDEMANDDAY), _
                        GetFieldNameValue("AMT", dblAMT), _
                        GetFieldNameValue("COMMI_RATE", dblCOMMI_RATE), _
                        GetFieldNameValue("COMMISSION", dblCOMMISSION), _
                        GetFieldNameValue("TBRDSTDATE", strTBRDSTDATE), _
                        GetFieldNameValue("TBRDEDDATE", strTBRDEDDATE), _
                        GetFieldNameValue("CNT", strCNT), _
                        GetFieldNameValue("GFLAG", strGFLAG), _
                        GetFieldNameValue("VOCH_TYPE", strVOCH_TYPE), _
                        GetFieldNameValue("TRU_TAX_FLAG", strTRU_TAX_FLAG), _
                        GetFieldNameValue("COMMI_TAX_FLAG", strCOMMI_TAX_FLAG), _
                        GetFieldNameValue("MEMO", strMEMO), _
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
                        GetFieldNameValue("YEARMON", strYEARMON), GetFieldNameValue("SEQ", dblSEQ)))

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
    '기능 : 해당 Entity에 Delete 처리
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function Update_TruTrans(ByVal strTRUST_YEARMON As String, _
                                    ByVal lngTrust_seq As Double, _
                                    ByVal strTRUTRANS As String, _
                                    ByVal strDEMANDDAY As String) As Integer

        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_TOTAL_MEDIUM SET TRU_TRANS_NO = '" & strTRUTRANS & "' , DEMANDDAY = '" & strDEMANDDAY & "'"
            strSQL = strSQL & " WHERE YEARMON = '" & strTRUST_YEARMON & "' AND SEQ = " & lngTrust_seq
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_TruTrans")
        End Try
    End Function


    '*****************************************************************
    '입력 : strSQL = SQL 문 By KTH
    '반환 : 처리건수
    '기능 : 신탁테이블의 거래명세삭제시에 trans_trans_no 를 없애준다.
    '*****************************************************************
    Public Function TRANSDeleteUpdateDo(ByVal strTRANSYEARMON As String, _
                                         ByVal strTRANSNO As String) As Integer
        Dim strSQL As String
        Try

            strSQL = strSQL & "  UPDATE MD_TOTAL_MEDIUM "
            strSQL = strSQL & "  SET TRU_TRANS_NO = ''  "
            strSQL = strSQL & "  FROM MD_TOTAL_MEDIUM"
            strSQL = strSQL & "  WHERE YEARMON + '-' + CAST(SEQ AS VARCHAR(10)) IN("
            strSQL = strSQL & "   SELECT TRUST_YEARMON + '-' + CAST(TRUST_SEQ AS VARCHAR(10))"
            strSQL = strSQL & "   FROM MD_TOTALTRANS_DTL "
            strSQL = strSQL & "   WHERE TRANSYEARMON = '" & strTRANSYEARMON & "'"
            strSQL = strSQL & "   AND TRANSNO = '" & strTRANSNO & "'"
            strSQL = strSQL & "  )"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".TRANSDeleteUpdateDo")
        End Try
    End Function

    '위수탁 거래명세서 승인 처리
    Public Function GFLAGUpdate(ByVal strGFLAG As String, _
                                ByVal strYEARMON As String, _
                                ByVal strSEQ As Double) As Integer
        Dim strSQL
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Dim strUSER As String
        strUSER = mobjSCGLConfig.WRKUSR
        Try
            strSQL = "UPDATE MD_TOTAL_MEDIUM SET GFLAG = '" & strGFLAG & "'"
            strSQL = strSQL & " ,ATTR04 = '" & strUSER & "',ATTR05 =  '" & strNOW & "'"
            strSQL = strSQL & " WHERE YEARMON = '" & strYEARMON & "' AND SEQ = " & strSEQ

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".ConfirmUpdate")
        End Try
    End Function

    Public Function Update_CommiTrans(ByVal strTRANSYEARMON As String, _
                                ByVal lngTrust_seq As Double, _
                                ByVal strTRUTRANS As String) As Integer

        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_TOTAL_MEDIUM SET COMMI_TRANS_NO = '" & strTRUTRANS & "' "
            strSQL = strSQL & " WHERE YEARMON = '" & strTRANSYEARMON & "' AND SEQ = " & lngTrust_seq
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Delete_Batch")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문 By KTY
    '반환 : 처리건수
    '기능 : 신탁테이블의 거래명세삭제시에 commi_trans_no 를 없애준다.
    '*****************************************************************
    Public Function COMMIDeleteUpdateDo(ByVal strTRANSYEARMON As String, _
                                        ByVal strTRANSNO As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strSQL & "  UPDATE MD_TOTAL_MEDIUM "
            strSQL = strSQL & "  SET COMMI_TRANS_NO = '' "
            strSQL = strSQL & "  FROM MD_TOTAL_MEDIUM"
            strSQL = strSQL & "  WHERE YEARMON + '-' + CAST(SEQ AS VARCHAR(10)) IN("
            strSQL = strSQL & "  	SELECT TRUST_YEARMON + '-' + CAST(TRUST_SEQ AS VARCHAR(10))"
            strSQL = strSQL & "  	FROM MD_TOTALCOMMI_DTL "
            strSQL = strSQL & "  	WHERE TRANSYEARMON = '" & strTRANSYEARMON & "'"
            strSQL = strSQL & "  	AND TRANSNO = '" & strTRANSNO & "'"
            strSQL = strSQL & "  )"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".COMMIDeleteUpdateDo")
        End Try
    End Function

    'Update_TruTax(strTRANSYEARMON, lngTRANSNO, lngSEQ, strTAXYEARMON, intTAXNO)
    '*****************************************************************
    '입력 : strSQL = SQL 문 By KTH
    '반환 : 처리건수
    '기능 : 신탁테이블의 세금계산서 생성시 번호 업데이트
    '*****************************************************************
    Public Function Update_TruTax(ByVal strTRANSNO As String, _
                                  ByVal strTAXNO As String) As Integer

        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_TOTAL_MEDIUM SET TRU_TAX_NO = '" & strTAXNO & "' "
            strSQL = strSQL & "  WHERE TRU_TRANS_NO = '" & strTRANSNO & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_TruTax")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문 By KTH
    '반환 : 처리건수
    '기능 : 신탁테이블의 세금계산서 삭제시 번호 업데이트
    '*****************************************************************
    Public Function TruTaxDeleteUpdateDo(ByVal strTAXNO As String) As Integer

        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_TOTAL_MEDIUM SET TRU_TAX_NO = '' "
            strSQL = strSQL & "  WHERE TRU_TAX_NO = '" & strTAXNO & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".TruTaxDeleteUpdateDo")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문 By KTH
    '반환 : 처리건수
    '기능 : 신탁테이블의 세금계산서 삭제시 번호 업데이트
    '*****************************************************************
    Public Function CommiTaxDeleteUpdateDo(ByVal strTAXNO As String) As Integer

        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_TOTAL_MEDIUM SET COMMI_TAX_NO = '' "
            strSQL = strSQL & "  WHERE COMMI_TAX_NO = '" & strTAXNO & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".CommiTaxDeleteUpdateDo")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문 By KTH
    '반환 : 처리건수
    '기능 : 신탁에 수수료세금계산서 번호 업데이트
    '*****************************************************************
    Public Function Update_CommiTax(ByVal strTRANSNO As String, _
                                    ByVal strTAXNO As String) As Integer

        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_TOTAL_MEDIUM SET COMMI_TAX_NO = '" & strTAXNO & "' "
            strSQL = strSQL & "  WHERE COMMI_TRANS_NO = '" & strTRANSNO & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_CommiTax")
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
        MyBase.EntityName = "MD_TOTAL_MEDIUM"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class

