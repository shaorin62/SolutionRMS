'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceMD_TRUTAX_HDR.vb ( MD_TRUTAX_HDR Entity 처리 Class)
'기      능 : MD_TRUTAX_HDR Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-03-06 오후 8:08:54 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceMD_SENDTAX_MST
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_SENDTAX_MST"    '자신의 클래스명
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
            ByVal strMEDFLAG As String, _
            Optional ByVal strCOMPANYCD As String = OPTIONAL_STR, _
            Optional ByVal strBILLNO As String = OPTIONAL_STR, _
            Optional ByVal strFISCALLYY As String = OPTIONAL_STR, _
            Optional ByVal strBILLFLAG As String = OPTIONAL_STR, _
            Optional ByVal strSUPPBSN As String = OPTIONAL_STR, _
            Optional ByVal strSUPPLDSCR As String = OPTIONAL_STR, _
            Optional ByVal strSUPPCEO As String = OPTIONAL_STR, _
            Optional ByVal strSUPPADDR As String = OPTIONAL_STR, _
            Optional ByVal strSUPPBUSICOND As String = OPTIONAL_STR, _
            Optional ByVal strSUPPBUSIITEM As String = OPTIONAL_STR, _
            Optional ByVal strBUYBSN As String = OPTIONAL_STR, _
            Optional ByVal strBUYLDSCR As String = OPTIONAL_STR, _
            Optional ByVal strBUYCEO As String = OPTIONAL_STR, _
            Optional ByVal strBUYADDR As String = OPTIONAL_STR, _
            Optional ByVal strBUYBUSICOND As String = OPTIONAL_STR, _
            Optional ByVal strBUYBUSIITEM As String = OPTIONAL_STR, _
            Optional ByVal strREGDATE As String = OPTIONAL_STR, _
            Optional ByVal dblTOTAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUPPAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblVATAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strBILLRMRK As String = OPTIONAL_STR, _
            Optional ByVal strTITLE As String = OPTIONAL_STR, _
            Optional ByVal strREQFLAG As String = OPTIONAL_STR, _
            Optional ByVal strNORMFLAG As String = OPTIONAL_STR, _
            Optional ByVal strRECEIPTID As String = OPTIONAL_STR, _
            Optional ByVal strRECEIPTNM As String = OPTIONAL_STR, _
            Optional ByVal strPURTEAMCD As String = OPTIONAL_STR, _
            Optional ByVal strINSDATE As String = OPTIONAL_STR, _
            Optional ByVal strBILLSEQ As String = OPTIONAL_STR, _
            Optional ByVal strSUPPDATE As String = OPTIONAL_STR, _
            Optional ByVal strITEMNM As String = OPTIONAL_STR, _
            Optional ByVal dblSIZE As Double = OPTIONAL_NUM, _
            Optional ByVal dblQTY As Double = OPTIONAL_NUM, _
            Optional ByVal dblUNITPRC As Double = OPTIONAL_NUM, _
            Optional ByVal strITEMRMRK As String = OPTIONAL_STR, _
            Optional ByVal strEDITTYPECD As String = OPTIONAL_STR, _
            Optional ByVal strFIRSTBILLNO As String = OPTIONAL_STR, _
            Optional ByVal strBUYNM As String = OPTIONAL_STR, _
            Optional ByVal strBUYEMAIL As String = OPTIONAL_STR, _
            Optional ByVal strSUPPNM As String = OPTIONAL_STR, _
            Optional ByVal strSUPPEMAIL As String = OPTIONAL_STR, _
            Optional ByVal strCANCEL_YN As String = OPTIONAL_STR, _
            Optional ByVal strSENDNTS_YN As String = OPTIONAL_STR, _
            Optional ByVal strISTRUST_YN As String = OPTIONAL_STR, _
            Optional ByVal strTRUST_CUSCD As String = OPTIONAL_STR, _
            Optional ByVal strRMSNO As String = OPTIONAL_STR, _
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
            BuildNameValues(",", "TAXYEARMON", strTAXYEARMON, strFields, strValues)
            BuildNameValues(",", "TAXNO", dblTAXNO, strFields, strValues)
            BuildNameValues(",", "MEDFLAG", strMEDFLAG, strFields, strValues)
            BuildNameValues(",", "COMPANYCD", strCOMPANYCD, strFields, strValues)
            BuildNameValues(",", "BILLNO", strBILLNO, strFields, strValues)
            BuildNameValues(",", "FISCALLYY", strFISCALLYY, strFields, strValues)
            BuildNameValues(",", "BILLFLAG", strBILLFLAG, strFields, strValues)
            BuildNameValues(",", "SUPPBSN", strSUPPBSN, strFields, strValues)
            BuildNameValues(",", "SUPPLDSCR", strSUPPLDSCR, strFields, strValues)
            BuildNameValues(",", "SUPPCEO", strSUPPCEO, strFields, strValues)
            BuildNameValues(",", "SUPPADDR", strSUPPADDR, strFields, strValues)
            BuildNameValues(",", "SUPPBUSICOND", strSUPPBUSICOND, strFields, strValues)
            BuildNameValues(",", "SUPPBUSIITEM", strSUPPBUSIITEM, strFields, strValues)
            BuildNameValues(",", "BUYBSN", strBUYBSN, strFields, strValues)
            BuildNameValues(",", "BUYLDSCR", strBUYLDSCR, strFields, strValues)
            BuildNameValues(",", "BUYCEO", strBUYCEO, strFields, strValues)
            BuildNameValues(",", "BUYADDR", strBUYADDR, strFields, strValues)
            BuildNameValues(",", "BUYBUSICOND", strBUYBUSICOND, strFields, strValues)
            BuildNameValues(",", "BUYBUSIITEM", strBUYBUSIITEM, strFields, strValues)
            BuildNameValues(",", "REGDATE", strREGDATE, strFields, strValues)
            BuildNameValues(",", "TOTAMT", dblTOTAMT, strFields, strValues)
            BuildNameValues(",", "SUPPAMT", dblSUPPAMT, strFields, strValues)
            BuildNameValues(",", "VATAMT", dblVATAMT, strFields, strValues)
            BuildNameValues(",", "BILLRMRK", strBILLRMRK, strFields, strValues)
            BuildNameValues(",", "TITLE", strTITLE, strFields, strValues)
            BuildNameValues(",", "REQFLAG", strREQFLAG, strFields, strValues)
            BuildNameValues(",", "NORMFLAG", strNORMFLAG, strFields, strValues)
            BuildNameValues(",", "RECEIPTID", strRECEIPTID, strFields, strValues)
            BuildNameValues(",", "RECEIPTNM", strRECEIPTNM, strFields, strValues)
            BuildNameValues(",", "PURTEAMCD", strPURTEAMCD, strFields, strValues)
            BuildNameValues(",", "INSDATE", strINSDATE, strFields, strValues)
            BuildNameValues(",", "BILLSEQ", strBILLSEQ, strFields, strValues)
            BuildNameValues(",", "SUPPDATE", strSUPPDATE, strFields, strValues)
            BuildNameValues(",", "ITEMNM", strITEMNM, strFields, strValues)
            BuildNameValues(",", "SIZE", dblSIZE, strFields, strValues)
            BuildNameValues(",", "QTY", dblQTY, strFields, strValues)
            BuildNameValues(",", "UNITPRC", dblUNITPRC, strFields, strValues)
            BuildNameValues(",", "ITEMRMRK", strITEMRMRK, strFields, strValues)
            BuildNameValues(",", "EDITTYPECD", strEDITTYPECD, strFields, strValues)
            BuildNameValues(",", "FIRSTBILLNO", strFIRSTBILLNO, strFields, strValues)
            BuildNameValues(",", "BUYNM", strBUYNM, strFields, strValues)
            BuildNameValues(",", "BUYEMAIL", strBUYEMAIL, strFields, strValues)
            BuildNameValues(",", "SUPPNM", strSUPPNM, strFields, strValues)
            BuildNameValues(",", "SUPPEMAIL", strSUPPEMAIL, strFields, strValues)
            BuildNameValues(",", "CANCEL_YN", strCANCEL_YN, strFields, strValues)
            BuildNameValues(",", "SENDNTS_YN", strSENDNTS_YN, strFields, strValues)
            BuildNameValues(",", "ISTRUST_YN", strISTRUST_YN, strFields, strValues)
            BuildNameValues(",", "TRUST_CUSCD", strTRUST_CUSCD, strFields, strValues)
            BuildNameValues(",", "RMSNO", strRMSNO, strFields, strValues)
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
    Public Function DeleteDo(Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, Optional ByVal dblTAXNO As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("TAXYEARMON", strTAXYEARMON), GetFieldNameValue("TAXNO", dblTAXNO)))

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
    Public Function UPDATE_PRINT_MEDIUM(ByVal strTAXYEARMON, ByVal intTAXNO, ByVal strTRUST_SEQ) As Integer
        Dim strSQL As String 'strTAXYEARMON, intTAXNO, strTRUST_SEQ

        Try
            strSQL = " UPDATE MD_BOOKING_MEDIUM "
            strSQL = strSQL & "   SET TRU_TAX_NO = '" & strTAXYEARMON & "-" & intTAXNO & "'"
            strSQL = strSQL & "   WHERE YEARMON = '" & strTAXYEARMON & "' AND SEQ = '" & strTRUST_SEQ & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UPDATE_TRANS_DO")
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
        MyBase.EntityName = "MD_SENDTAX_MST"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class
