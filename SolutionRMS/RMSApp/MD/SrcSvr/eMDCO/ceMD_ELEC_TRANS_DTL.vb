'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceMD_ELEC_TRANS_DTL.vb ( MD_ELEC_TRANS_DTL Entity 처리 Class)
'기      능 : MD_ELEC_TRANS_DTL Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-04-21 오후 9:35:10 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceMD_ELEC_TRANS_DTL
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_ELEC_TRANS_DTL"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(ByVal strTRANSYEARMON As String, _
            ByVal dblTRANSNO As Double, _
            ByVal dblSEQ As Double, _
            ByVal strCLIENTCODE As String, _
            ByVal strCLIENTSUBCODE As String, _
            ByVal strTIMCODE As String, _
            ByVal strREAL_MED_CODE As String, _
            ByVal strMEDCODE As String, _
            ByVal strSUBSEQ As String, _
            ByVal strMATTERCODE As String, _
            Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
            Optional ByVal strPRINTDAY As String = OPTIONAL_STR, _
            Optional ByVal strPROGRAM As String = OPTIONAL_STR, _
            Optional ByVal strADLOCALFLAG As String = OPTIONAL_STR, _
            Optional ByVal strWEEKDAY As String = OPTIONAL_STR, _
            Optional ByVal dblCNT As Double = OPTIONAL_NUM, _
            Optional ByVal dblPRICE As Double = OPTIONAL_NUM, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblVAT As Double = OPTIONAL_NUM, _
            Optional ByVal strMED_FLAG As String = OPTIONAL_STR, _
            Optional ByVal strSPONSOR As String = OPTIONAL_STR, _
            Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strVOCH_TYPE As String = OPTIONAL_STR, _
            Optional ByVal strTRU_TAX_FLAG As String = OPTIONAL_STR, _
            Optional ByVal strTRUST_YEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTRUST_SEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTAXNO As Double = OPTIONAL_NUM, _
            Optional ByVal dblCONFIRMFLAG As Double = OPTIONAL_NUM, _
            Optional ByVal dblPRINT_SEQ As Double = OPTIONAL_NUM, _
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
            BuildNameValues(",", "TRANSYEARMON", strTRANSYEARMON, strFields, strValues)
            BuildNameValues(",", "TRANSNO", dblTRANSNO, strFields, strValues)
            BuildNameValues(",", "SEQ", dblSEQ, strFields, strValues)
            BuildNameValues(",", "CLIENTCODE", strCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "CLIENTSUBCODE", strCLIENTSUBCODE, strFields, strValues)
            BuildNameValues(",", "TIMCODE", strTIMCODE, strFields, strValues)
            BuildNameValues(",", "REAL_MED_CODE", strREAL_MED_CODE, strFields, strValues)
            BuildNameValues(",", "MEDCODE", strMEDCODE, strFields, strValues)
            BuildNameValues(",", "SUBSEQ", strSUBSEQ, strFields, strValues)
            BuildNameValues(",", "MATTERCODE", strMATTERCODE, strFields, strValues)
            BuildNameValues(",", "DEPT_CD", strDEPT_CD, strFields, strValues)
            BuildNameValues(",", "DEMANDDAY", strDEMANDDAY, strFields, strValues)
            BuildNameValues(",", "PRINTDAY", strPRINTDAY, strFields, strValues)
            BuildNameValues(",", "PROGRAM", strPROGRAM, strFields, strValues)
            BuildNameValues(",", "ADLOCALFLAG", strADLOCALFLAG, strFields, strValues)
            BuildNameValues(",", "WEEKDAY", strWEEKDAY, strFields, strValues)
            BuildNameValues(",", "CNT", dblCNT, strFields, strValues)
            BuildNameValues(",", "PRICE", dblPRICE, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "VAT", dblVAT, strFields, strValues)
            BuildNameValues(",", "MED_FLAG", strMED_FLAG, strFields, strValues)
            BuildNameValues(",", "SPONSOR", strSPONSOR, strFields, strValues)
            BuildNameValues(",", "EXCLIENTCODE", strEXCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "VOCH_TYPE", strVOCH_TYPE, strFields, strValues)
            BuildNameValues(",", "TRU_TAX_FLAG", strTRU_TAX_FLAG, strFields, strValues)
            BuildNameValues(",", "TRUST_YEARMON", strTRUST_YEARMON, strFields, strValues)
            BuildNameValues(",", "TRUST_SEQ", dblTRUST_SEQ, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "TAXYEARMON", strTAXYEARMON, strFields, strValues)
            BuildNameValues(",", "TAXNO", dblTAXNO, strFields, strValues)
            BuildNameValues(",", "CONFIRMFLAG", dblCONFIRMFLAG, strFields, strValues)
            BuildNameValues(",", "PRINT_SEQ", dblPRINT_SEQ, strFields, strValues)
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
    Public Function UpdateDo(ByVal strTRANSYEARMON As String, _
            ByVal dblTRANSNO As Double, _
            ByVal dblSEQ As Double, _
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTSUBCODE As String = OPTIONAL_STR, _
            Optional ByVal strTIMCODE As String = OPTIONAL_STR, _
            Optional ByVal strREAL_MED_CODE As String = OPTIONAL_STR, _
            Optional ByVal strMEDCODE As String = OPTIONAL_STR, _
            Optional ByVal strSUBSEQ As String = OPTIONAL_STR, _
            Optional ByVal strMATTERCODE As String = OPTIONAL_STR, _
            Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
            Optional ByVal strPROGRAM As String = OPTIONAL_STR, _
            Optional ByVal strADLOCALFLAG As String = OPTIONAL_STR, _
            Optional ByVal strWEEKDAY As String = OPTIONAL_STR, _
            Optional ByVal dblCNT As Double = OPTIONAL_NUM, _
            Optional ByVal dblPRICE As Double = OPTIONAL_NUM, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblVAT As Double = OPTIONAL_NUM, _
            Optional ByVal strMED_FLAG As String = OPTIONAL_STR, _
            Optional ByVal strSPONSOR As String = OPTIONAL_STR, _
            Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strVOCH_TYPE As String = OPTIONAL_STR, _
            Optional ByVal strTRU_TAX_FLAG As String = OPTIONAL_STR, _
            Optional ByVal strTRUST_YEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTRUST_SEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTAXNO As Double = OPTIONAL_NUM, _
            Optional ByVal dblCONFIRMFLAG As Double = OPTIONAL_NUM, _
            Optional ByVal dblPRINT_SEQ As Double = OPTIONAL_NUM, _
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
                        GetFieldNameValue("TRANSYEARMON", strTRANSYEARMON), _
                        GetFieldNameValue("TRANSNO", dblTRANSNO), _
                        GetFieldNameValue("SEQ", dblSEQ), _
                        GetFieldNameValue("CLIENTCODE", strCLIENTCODE), _
                        GetFieldNameValue("CLIENTSUBCODE", strCLIENTSUBCODE), _
                        GetFieldNameValue("TIMCODE", strTIMCODE), _
                        GetFieldNameValue("REAL_MED_CODE", strREAL_MED_CODE), _
                        GetFieldNameValue("MEDCODE", strMEDCODE), _
                        GetFieldNameValue("SUBSEQ", strSUBSEQ), _
                        GetFieldNameValue("MATTERCODE", strMATTERCODE), _
                        GetFieldNameValue("DEPT_CD", strDEPT_CD), _
                        GetFieldNameValue("DEMANDDAY", strDEMANDDAY), _
                        GetFieldNameValue("PROGRAM", strPROGRAM), _
                        GetFieldNameValue("ADLOCALFLAG", strADLOCALFLAG), _
                        GetFieldNameValue("WEEKDAY", strWEEKDAY), _
                        GetFieldNameValue("CNT", dblCNT), _
                        GetFieldNameValue("PRICE", dblPRICE), _
                        GetFieldNameValue("AMT", dblAMT), _
                        GetFieldNameValue("VAT", dblVAT), _
                        GetFieldNameValue("MED_FLAG", strMED_FLAG), _
                        GetFieldNameValue("SPONSOR", strSPONSOR), _
                        GetFieldNameValue("EXCLIENTCODE", strEXCLIENTCODE), _
                        GetFieldNameValue("VOCH_TYPE", strVOCH_TYPE), _
                        GetFieldNameValue("TRU_TAX_FLAG", strTRU_TAX_FLAG), _
                        GetFieldNameValue("TRUST_YEARMON", strTRUST_YEARMON), _
                        GetFieldNameValue("TRUST_SEQ", dblTRUST_SEQ), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("TAXYEARMON", strTAXYEARMON), _
                        GetFieldNameValue("TAXNO", dblTAXNO), _
                        GetFieldNameValue("CONFIRMFLAG", dblCONFIRMFLAG), _
                        GetFieldNameValue("PRINT_SEQ", dblPRINT_SEQ), _
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
                        GetFieldNameValue("CUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("CDATE", strNOW), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("TRANSYEARMON", strTRANSYEARMON), GetFieldNameValue("TRANSNO", dblTRANSNO), GetFieldNameValue("SEQ", dblSEQ)))

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
    Public Function DeleteDo(Optional ByVal strTRANSYEARMON As String = OPTIONAL_STR, Optional ByVal dblTRANSNO As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("TRANSYEARMON", strTRANSYEARMON), GetFieldNameValue("TRANSNO", dblTRANSNO)))

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
                            Optional ByVal strTRANSYEARMON As String = OPTIONAL_STR, _
Optional ByVal dblTRANSNO As Double = OPTIONAL_NUM, _
Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
                            Optional ByVal strSelFields As String = "*", _
                            Optional ByVal intLimitRow As Integer = 0, _
                            Optional ByVal intSelMode As Integer = SELMODE.ARR, _
                            Optional ByVal blnBindingHeader As Boolean = False) As Object
        Dim strSQL As String
        Dim strKeyFields As String

        Try
            strKeyFields = BuildFields("AND", _
                                    GetFieldNameValue("TRANSYEARMON", strTRANSYEARMON), GetFieldNameValue("TRANSNO", dblTRANSNO), GetFieldNameValue("SEQ", dblSEQ))

            Return SelectDoExt(intRowCnt, intColCnt, strSelFields, strKeyFields, intLimitRow, intSelMode, blnBindingHeader)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".SeleteDo")
        End Try
    End Function
#End Region


    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function UPDATE_TRUTAXNO_DO(ByVal strTRANSYEARMON As String, _
                                       ByVal strTRUTAXNO As String, _
                                       ByVal strSEQ As String) As Integer
        'strTRANSYEARMON, strTRUTAXNO,strSEQ
        Dim strSQL As String

        Try
            strSQL = " UPDATE MD_ELECTRIC_MEDIUM  SET TRU_TAX_NO = '" & strTRUTAXNO & "' "
            strSQL = strSQL & " WHERE YEARMON = '" & strTRANSYEARMON & "'"
            strSQL = strSQL & " AND SEQ = '" & strSEQ & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UPDATE_TRUTAXNO_DO")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문 By KTH
    '반환 : 처리건수
    '기능 : 거래명세 디테일 의 세금계산서 번호 업데이트
    '*****************************************************************
    Public Function Update_TruTax(ByVal strTRANSYEARMON As String, _
                                  ByVal lngTRANSNO As Double, _
                                  ByVal strTAXYEARMON As String, _
                                  ByVal intTAXNO As Double) As Integer
        'strTRANSNO, strTAXNO
        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_ELEC_TRANS_DTL SET TAXYEARMON = '" & strTAXYEARMON & "',TAXNO = " & intTAXNO
            strSQL = strSQL & "  WHERE TRANSYEARMON = '" & strTRANSYEARMON & "' AND TRANSNO =" & lngTRANSNO
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_TruTax")
        End Try
    End Function





    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 거래명세서 저장후 신탁에 거래명세서번호를 UPDATE
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function UPDATE_TRANS_DO(ByVal strTRANSNO As String, _
                                    ByVal strYEARMON As String, _
                                    ByVal strSEQ As String) As Integer
        'strTRANSNO, strYEARMON, strSEQ
        Dim strSQL As String

        Try

            strSQL = "UPDATE MD_ELECTRIC_MEDIUM SET TRU_TRANS_NO = '" & strTRANSNO & "'"
            strSQL = strSQL & " WHERE YEARMON = '" & strYEARMON & "' AND SEQ = '" & strSEQ & "'"


            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UPDATE_TRANS_DO")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 세금계산서 번호를 거래명세 디테일 ATTR02에 넣는다.
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function TRUST_InsertRtn(ByVal strSQL As String)
        Try
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".TRUST_InsertRtn")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문 By KTH
    '반환 : 처리건수
    '기능 : 삭제시 거래명세 디테일 의 세금계산서 번호 삭제
    '*****************************************************************
    Public Function TruTaxDeleteUpdateDo(ByVal strTAXYEARMON As String, _
                                         ByVal intTAXNO As Double) As Integer
        'strTRANSNO, strTAXNO
        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_ELEC_TRANS_DTL SET TAXYEARMON = '',TAXNO = NULL "
            strSQL = strSQL & "  WHERE TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO =" & intTAXNO
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".TruTaxDeleteUpdateDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문 By KTH
    '반환 : 처리건수
    '기능 : 삭제시 거래명세 디테일 의 세금계산서 번호 삭제
    '*****************************************************************
    'Public Function TruTaxDeleteUpdateDo2(ByVal strATTR05 As String) As Integer
    '    'strTRANSNO, strTAXNO
    '    Dim strSQL As String
    '    Try
    '        strSQL = "UPDATE MD_ELEC_TRANS_DTL SET TAXYEARMON = '',TAXNO = NULL "
    '        strSQL = strSQL & "  WHERE TRANSYEARMON + '-' + CAST(TRANSNO AS VARCHAR(20))  = '" & strATTR05 & "'"
    '        Return ProcEntity(strSQL)
    '    Catch err As Exception
    '        Throw RaiseSysErr(err, CLASS_NAME & ".TruTaxDeleteUpdateDo")
    '    End Try
    'End Function

    Public Function TruTaxDeleteUpdateDo2(ByVal strTAXYEARMON As String, _
                                          ByVal intTAXNO As Double) As Integer
        'strTRANSNO, strTAXNO
        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_ELEC_TRANS_DTL SET TAXYEARMON = '', TAXNO = NULL "
            strSQL = strSQL & " WHERE TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO =" & intTAXNO
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".TruTaxDeleteUpdateDo")
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
        MyBase.EntityName = "MD_ELEC_TRANS_DTL"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
End Class

'------->>엔티티 INSERT/UPDATE 샘플입니다. 반드시 자신의 환경에 맞추어서 변경하시기 바랍니다.
'=========================================================
'       'vntData Array를 사용할 때 Insert/Update 입니다.
'=========================================================
'        Dim intRtn As Integer
'        intRtn = mobjceMD_ELEC_TRANS_DTL.InsertDo( _
'                                       GetElement(vntData,"TRANSYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TRANSNO", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CLIENTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MEDCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"REAL_MED_CODE", intColCnt, intRow), _
'                                       GetElement(vntData,"DEPT_CD", intColCnt, intRow), _
'                                       GetElement(vntData,"DEMANDDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PRINTDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PROGRAM", intColCnt, intRow), _
'                                       GetElement(vntData,"ADLOCALFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"WEEKDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PRICE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CNT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"AMT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TRU_TAX_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"VAT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TRUST_SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"MEMO", intColCnt, intRow), _
'                                       GetElement(vntData,"MED_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"SPONSOR", intColCnt, intRow), _
'                                       GetElement(vntData,"TAXYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TAXNO", intColCnt, intRow, NULL_NUM, true ), _
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
'        intRtn = mobjceMD_ELEC_TRANS_DTL.UpdateDo( _
'                                       GetElement(vntData,"TRANSYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TRANSNO", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CLIENTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MEDCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"REAL_MED_CODE", intColCnt, intRow), _
'                                       GetElement(vntData,"DEPT_CD", intColCnt, intRow), _
'                                       GetElement(vntData,"DEMANDDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PRINTDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PROGRAM", intColCnt, intRow), _
'                                       GetElement(vntData,"ADLOCALFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"WEEKDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PRICE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CNT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"AMT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TRU_TAX_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"VAT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TRUST_SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"MEMO", intColCnt, intRow), _
'                                       GetElement(vntData,"MED_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"SPONSOR", intColCnt, intRow), _
'                                       GetElement(vntData,"TAXYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TAXNO", intColCnt, intRow, NULL_NUM, true ), _
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
'        intRtn = mobjceMD_ELEC_TRANS_DTL.InsertDo( _
'                                       XMLGetElement(xmlRoot,"TRANSYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TRANSNO", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CLIENTCODE"), _
'                                       XMLGetElement(xmlRoot,"MEDCODE"), _
'                                       XMLGetElement(xmlRoot,"REAL_MED_CODE"), _
'                                       XMLGetElement(xmlRoot,"DEPT_CD"), _
'                                       XMLGetElement(xmlRoot,"DEMANDDAY"), _
'                                       XMLGetElement(xmlRoot,"PRINTDAY"), _
'                                       XMLGetElement(xmlRoot,"PROGRAM"), _
'                                       XMLGetElement(xmlRoot,"ADLOCALFLAG"), _
'                                       XMLGetElement(xmlRoot,"WEEKDAY"), _
'                                       XMLGetElement(xmlRoot,"PRICE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CNT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"AMT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TRU_TAX_FLAG"), _
'                                       XMLGetElement(xmlRoot,"VAT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TRUST_SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"MEMO"), _
'                                       XMLGetElement(xmlRoot,"MED_FLAG"), _
'                                       XMLGetElement(xmlRoot,"SPONSOR"), _
'                                       XMLGetElement(xmlRoot,"TAXYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TAXNO", NULL_NUM, true ), _
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
'        intRtn = mobjceMD_ELEC_TRANS_DTL.UpdateDo( _
'                                       XMLGetElement(xmlRoot,"TRANSYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TRANSNO", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CLIENTCODE"), _
'                                       XMLGetElement(xmlRoot,"MEDCODE"), _
'                                       XMLGetElement(xmlRoot,"REAL_MED_CODE"), _
'                                       XMLGetElement(xmlRoot,"DEPT_CD"), _
'                                       XMLGetElement(xmlRoot,"DEMANDDAY"), _
'                                       XMLGetElement(xmlRoot,"PRINTDAY"), _
'                                       XMLGetElement(xmlRoot,"PROGRAM"), _
'                                       XMLGetElement(xmlRoot,"ADLOCALFLAG"), _
'                                       XMLGetElement(xmlRoot,"WEEKDAY"), _
'                                       XMLGetElement(xmlRoot,"PRICE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CNT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"AMT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TRU_TAX_FLAG"), _
'                                       XMLGetElement(xmlRoot,"VAT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TRUST_SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"MEMO"), _
'                                       XMLGetElement(xmlRoot,"MED_FLAG"), _
'                                       XMLGetElement(xmlRoot,"SPONSOR"), _
'                                       XMLGetElement(xmlRoot,"TAXYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TAXNO", NULL_NUM, true ), _
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

