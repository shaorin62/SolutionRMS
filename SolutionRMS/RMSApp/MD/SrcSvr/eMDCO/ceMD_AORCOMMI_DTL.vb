'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceMD_AORCOMMI_DTL.vb ( MD_AORCOMMI_DTL Entity 처리 Class)
'기      능 : MD_AORCOMMI_DTL Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2007-12-26 오후 5:45:57 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceMD_AORCOMMI_DTL
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_AORCOMMI_DTL"    '자신의 클래스명
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
                    ByVal strREAL_MED_CODE As String, _
                    ByVal strMEDCODE As String, _
                    ByVal strDEPT_CD As String, _
                    Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
                    Optional ByVal strPRINTDAY As String = OPTIONAL_STR, _
                    Optional ByVal strTITLE As String = OPTIONAL_STR, _
                    Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
                    Optional ByVal dblCOMMI_RATE As Double = OPTIONAL_NUM, _
                    Optional ByVal dblCOMMISSION As Double = OPTIONAL_NUM, _
                    Optional ByVal dblVAT As Double = OPTIONAL_NUM, _
                    Optional ByVal dblOUT_AMT As Double = OPTIONAL_NUM, _
                    Optional ByVal strMED_FLAG As String = OPTIONAL_STR, _
                    Optional ByVal strVOCH_TYPE As String = OPTIONAL_STR, _
                    Optional ByVal strTRUST_YEARMON As String = OPTIONAL_STR, _
                    Optional ByVal dblTRUST_SEQ As Double = OPTIONAL_NUM, _
                    Optional ByVal strMEMO As String = OPTIONAL_STR, _
                    Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
                    Optional ByVal dblTAXNO As Double = OPTIONAL_NUM, _
                    Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
                    Optional ByVal dblPRINT_SEQ As String = OPTIONAL_STR, _
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
            BuildNameValues(",", "REAL_MED_CODE", strREAL_MED_CODE, strFields, strValues)
            BuildNameValues(",", "MEDCODE", strMEDCODE, strFields, strValues)
            BuildNameValues(",", "DEPT_CD", strDEPT_CD, strFields, strValues)
            BuildNameValues(",", "DEMANDDAY", strDEMANDDAY, strFields, strValues)
            BuildNameValues(",", "PRINTDAY", strPRINTDAY, strFields, strValues)
            BuildNameValues(",", "TITLE", strTITLE, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "COMMI_RATE", dblCOMMI_RATE, strFields, strValues)
            BuildNameValues(",", "COMMISSION", dblCOMMISSION, strFields, strValues)
            BuildNameValues(",", "VAT", dblVAT, strFields, strValues)
            BuildNameValues(",", "OUT_AMT", dblOUT_AMT, strFields, strValues)
            BuildNameValues(",", "MED_FLAG", strMED_FLAG, strFields, strValues)
            BuildNameValues(",", "VOCH_TYPE", strVOCH_TYPE, strFields, strValues)
            BuildNameValues(",", "TRUST_YEARMON", strTRUST_YEARMON, strFields, strValues)
            BuildNameValues(",", "TRUST_SEQ", dblTRUST_SEQ, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "TAXYEARMON", strTAXYEARMON, strFields, strValues)
            BuildNameValues(",", "TAXNO", dblTAXNO, strFields, strValues)
            BuildNameValues(",", "CONFIRMFLAG", strCONFIRMFLAG, strFields, strValues)
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
    '입력 : strSQL = SQL 문 By KTH
    '반환 : 처리건수
    '기능 : 거래명세 디테일 의 세금계산서 번호 업데이트
    '*****************************************************************
    Public Function Update_CommiTax(ByVal strTRANSYEARMON As String, _
                                    ByVal lngTRANSNO As Double, _
                                    ByVal lngSEQ As Double, _
                                    ByVal strTAXYEARMON As String, _
                                    ByVal intTAXNO As Double) As Integer
        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_AORCOMMI_DTL SET TAXYEARMON = '" & strTAXYEARMON & "',TAXNO = " & intTAXNO
            strSQL = strSQL & "  WHERE TRANSYEARMON = '" & strTRANSYEARMON & "' AND TRANSNO =" & lngTRANSNO & " AND SEQ =" & lngSEQ
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_CommiTax")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문 By KTH
    '반환 : 처리건수
    '기능 : 삭제시 거래명세 디테일 의 세금계산서 번호 삭제
    '*****************************************************************
    Public Function CommiTaxDeleteUpdateDo(ByVal strTAXYEARMON As String, _
                                           ByVal intTAXNO As Double) As Integer
        'strTRANSNO, strTAXNO
        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_AORCOMMI_DTL SET TAXYEARMON = '',TAXNO = NULL "
            strSQL = strSQL & "  WHERE TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO =" & intTAXNO
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".CommiTaxDeleteUpdateDo")
        End Try
    End Function

    '거래명세서 승인 처리
    Public Function ConfirmUpdate(ByVal strTRANSYEARMON As String, _
                                  ByVal strTRANSNO As Double) As Integer
        Dim strSQL
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Dim strUSER As String
        strUSER = mobjSCGLConfig.WRKUSR
        Try
            strSQL = "UPDATE MD_AORCOMMI_DTL "
            strSQL = strSQL & " SET CONFIRMFLAG = 1, "
            strSQL = strSQL & " ATTR04 = '" & strUSER & "', "
            strSQL = strSQL & " ATTR05 =  '" & strNOW & "'"
            strSQL = strSQL & " WHERE TRANSYEARMON = '" & strTRANSYEARMON & "' AND TRANSNO = " & strTRANSNO
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".ConfirmUpdate")
        End Try
    End Function


    '거래명세서 승인 처리
    Public Function ConfirmCancelUpdate(ByVal strTRANSYEARMON As String, _
                                        ByVal strTRANSNO As Double) As Integer
        Dim strSQL
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Dim strUSER As String
        strUSER = mobjSCGLConfig.WRKUSR
        Try
            strSQL = "UPDATE MD_AORCOMMI_DTL "
            strSQL = strSQL & " SET CONFIRMFLAG = 0, "
            strSQL = strSQL & " ATTR04 = '', "
            strSQL = strSQL & " ATTR05 =  ''"
            strSQL = strSQL & " WHERE TRANSYEARMON = '" & strTRANSYEARMON & "' AND TRANSNO = " & strTRANSNO
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".ConfirmCancelUpdate")
        End Try
    End Function


    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 거래명세서 완료건에 대하여 부가세 수정
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************

    Public Function Update_Vat(ByVal strTRANSYEARMON, ByVal strTRANSNO, ByVal intTRANSSEQ, ByVal lngVAT) As Integer
        Dim strSQL As String

        Try
            strSQL = " UPDATE MD_AORCOMMI_DTL "
            strSQL = strSQL & " SET VAT = " & lngVAT
            strSQL = strSQL & " WHERE TRANSYEARMON = '" & strTRANSYEARMON & "' "
            strSQL = strSQL & " AND TRANSNO = " & strTRANSNO
            strSQL = strSQL & " AND SEQ = " & intTRANSSEQ

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_Vat")
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
        MyBase.EntityName = "MD_AORCOMMI_DTL"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class
