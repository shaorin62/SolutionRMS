'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceSC_REALMEDCODE_MST.vb ( SC_REALMEDCODE_MST Entity 처리 Class)
'기      능 : SC_REALMEDCODE_MST Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-01-14 오전 11:09:29 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class cePD_VOCH_MST
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "cePD_VOCH_MST"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Update 처리
    '참고 : Key 조건과 Value Field가선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    'POSTINGDATE,CUSTOMERCODE,SUMM,BA,SUMAMT,VAT,SEMU,BP,DEMANDDAY,VENDOR,TAXYEARMON,TAXNO,GBN,VOCHNO,RMSNO

    Public Function InsertDo(Optional ByVal strPOSTINGDATE As String = OPTIONAL_STR, _
                             Optional ByVal strCUSTOMERCODE As String = OPTIONAL_STR, _
                             Optional ByVal strSUMM As String = OPTIONAL_STR, _
                             Optional ByVal strBA As String = OPTIONAL_STR, _
                             Optional ByVal strCOSTCENTER As String = OPTIONAL_STR, _
                             Optional ByVal strAMT As Double = OPTIONAL_NUM, _
                             Optional ByVal strVAT As Double = OPTIONAL_NUM, _
                             Optional ByVal strSEMU As String = OPTIONAL_STR, _
                             Optional ByVal strBP As String = OPTIONAL_STR, _
                             Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
                             Optional ByVal strDUEDATE As String = OPTIONAL_STR, _
                             Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
                             Optional ByVal strTAXNO As Double = OPTIONAL_NUM, _
                             Optional ByVal strGBN As String = OPTIONAL_STR, _
                             Optional ByVal strVOCHNO As String = OPTIONAL_STR, _
                             Optional ByVal strRMSNO As String = OPTIONAL_STR, _
                             Optional ByVal strDOCUMENTDATE As String = OPTIONAL_STR, _
                             Optional ByVal strPAYCODE As String = OPTIONAL_STR, _
                             Optional ByVal strBANKTYPE As String = OPTIONAL_STR, _
                             Optional ByVal strBMORDER As String = OPTIONAL_STR, _
                             Optional ByVal strPREPAYMENT As String = OPTIONAL_STR, _
                             Optional ByVal strFROMDATE As String = OPTIONAL_STR, _
                             Optional ByVal strTODATE As String = OPTIONAL_STR, _
                             Optional ByVal strSUMMTEXT As String = OPTIONAL_STR, _
                             Optional ByVal strACCOUNT As String = OPTIONAL_STR, _
                             Optional ByVal strDEBTOR As String = OPTIONAL_STR, _
                             Optional ByVal strATTR01 As String = OPTIONAL_STR) As Integer


        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        'POSTINGDATE,CUSTOMERCODE,SUMM,BA,COSTCENTER,AMT,VAT,SEMU,BP,DEMANDDAY,TAXYEARMON,TAXNO,GBN,VOCHNO,RMSNO,DOCUMENTDATE,PREPAYMENT,FROMDATE,TODATE,SUMMTEXT,ACCOUNT,DEBTOR
        Try
            BuildNameValues(",", "POSTINGDATE", strPOSTINGDATE, strFields, strValues)
            BuildNameValues(",", "CUSTOMERCODE", strCUSTOMERCODE, strFields, strValues)
            BuildNameValues(",", "SUMM", strSUMM, strFields, strValues)
            BuildNameValues(",", "BA", strBA, strFields, strValues)
            BuildNameValues(",", "COSTCENTER", strCOSTCENTER, strFields, strValues)
            BuildNameValues(",", "AMT", strAMT, strFields, strValues)
            BuildNameValues(",", "VAT", strVAT, strFields, strValues)
            BuildNameValues(",", "SEMU", strSEMU, strFields, strValues)
            BuildNameValues(",", "BP", strBP, strFields, strValues)
            BuildNameValues(",", "DEMANDDAY", strDEMANDDAY, strFields, strValues)
            BuildNameValues(",", "DUEDATE", strDUEDATE, strFields, strValues)
            BuildNameValues(",", "TAXYEARMON", strTAXYEARMON, strFields, strValues)
            BuildNameValues(",", "TAXNO", strTAXNO, strFields, strValues)
            BuildNameValues(",", "GBN", strGBN, strFields, strValues)
            BuildNameValues(",", "VOCHNO", strVOCHNO, strFields, strValues)
            BuildNameValues(",", "RMSNO", strRMSNO, strFields, strValues)
            BuildNameValues(",", "DOCUMENTDATE", strDOCUMENTDATE, strFields, strValues)
            BuildNameValues(",", "PAYCODE", strPAYCODE, strFields, strValues)
            BuildNameValues(",", "BANKTYPE", strBANKTYPE, strFields, strValues)
            BuildNameValues(",", "BMORDER", strBMORDER, strFields, strValues)
            BuildNameValues(",", "PREPAYMENT", strPREPAYMENT, strFields, strValues)
            BuildNameValues(",", "FROMDATE", strFROMDATE, strFields, strValues)
            BuildNameValues(",", "TODATE", strTODATE, strFields, strValues)
            BuildNameValues(",", "SUMMTEXT", strSUMMTEXT, strFields, strValues)
            BuildNameValues(",", "ACCOUNT", strACCOUNT, strFields, strValues)
            BuildNameValues(",", "DEBTOR", strDEBTOR, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)
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

    Public Function UpdateRtn_VOCHNO(ByVal strGBN As String, _
                                     ByVal strTAXYEARMON As String, _
                                     ByVal intTAXNO As Integer, _
                                     ByVal strDOC_STATUS As String, _
                                     ByVal strDOC_MESSAGE As String, _
                                     ByVal strVOCHNO As String, _
                                     ByVal strIF_KEY As String) As Integer
        Dim strSQL As String
        Try
            '전표업데이트 시작
            strSQL = "UPDATE PD_VOCH_MST "
            strSQL = strSQL & " SET VOCHNO = '" & strVOCHNO & "', ERRCODE = '" & strDOC_STATUS & "', ERRMSG = '" & strDOC_MESSAGE & "', IF_GUBUN ='" & strIF_KEY & "'"
            strSQL = strSQL & " WHERE GBN = '" & strGBN & "' AND  TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO = " & intTAXNO


            '오류전표 아닌것만 세금계산서및 신탁 업데이트 시작
            If strDOC_STATUS = "" Then
                If strGBN = "P" Then
                    strSQL = strSQL & " ;update a set a.vochno = b.vochno "
                    strSQL = strSQL & " from pd_tax_hdr a,  (select taxyearmon,taxno,vochno,gbn from pd_voch_mst where ISNULL(errcode,'') = '' and gbn = 'P' and  TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO = " & intTAXNO & ") b "
                    strSQL = strSQL & " where a.taxyearmon = b.taxyearmon and a.taxno = CAST(b.taxno AS NUMERIC) and b.gbn = 'P'"

                ElseIf strGBN = "B" Then
                    strSQL = strSQL & " ;update a set a.vochno = b.vochno "
                    strSQL = strSQL & " from pd_exe_dtl a , (select taxyearmon,taxno,vochno,gbn from PD_voch_mst where ISNULL(errcode,'') = '' and gbn = 'B' and  TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO = " & intTAXNO & ") b "
                    strSQL = strSQL & " where substring(a.purchaseno,1,6) = b.taxyearmon and substring(a.purchaseno,7,4) = b.taxno and gbn = 'B' "

                End If
            End If

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_VOCHNO")
        End Try
    End Function

    Public Function UpdateDelete_P(ByVal strTAXYEARMON As String, _
                                   ByVal strTAXNO As Double) As Integer
        Dim strSQL As String
        Try
            strSQL = "UPDATE PD_TAX_HDR SET VOCHNO ='' WHERE TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO = '" & strTAXNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDelete_P")
        End Try
    End Function

    Public Function UpdateDelete_B(ByVal strTAXYEARMON As String, _
                                   ByVal strTAXNO As Double) As Integer
        Dim strSQL As String
        Try
            strSQL = "UPDATE PD_EXE_DTL SET VOCHNO ='' WHERE substring(purchaseno,1,6) = '" & strTAXYEARMON & "' AND cast( substring(purchaseno,7,len(purchaseno))  as numeric ) = " & strTAXNO

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDelete_B")
        End Try
    End Function


    Public Function Delete_vochno(ByVal strTAXYEARMON As String, _
                                  ByVal strVOCHNO As String, _
                                  ByVal strGBN As String) As Integer
        Dim strSQL As String
        Try

            strSQL = "DELETE FROM  PD_VOCH_MST WHERE TAXYEARMON ='" & strTAXYEARMON & "' "
            strSQL = strSQL & " AND ISNULL(GBN,'') = '" & strGBN & "'"
            strSQL = strSQL & " AND VOCHNO = '" & strVOCHNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Delete_vochno")
        End Try
    End Function


    Public Function Delete_ERR(ByVal strTAXYEARMON As String, _
                              ByVal strTAXNO As String, _
                              ByVal strGFLAG As String) As Integer
        Dim strSQL As String
        Try
            strSQL = "DELETE FROM PD_VOCH_MST WHERE TAXYEARMON ='" & strTAXYEARMON & "' AND TAXNO ='" & strTAXNO & "' AND GBN ='" & strGFLAG & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDelete")
        End Try
    End Function

    Public Function UpdateRtn_DELETEERR(ByVal strGBN As String, _
                                        ByVal strTAXYEARMON As String, _
                                        ByVal intTAXNO As Integer, _
                                        ByVal strDOC_STATUS As String, _
                                        ByVal strDOC_MESSAGE As String, _
                                        ByVal strVOCHNO As String) As Integer
        Dim strSQL As String
        Try

            '전표업데이트 시작
            strSQL = "UPDATE PD_VOCH_MST "
            strSQL = strSQL & " SET ERRCODE = '" & strDOC_STATUS & "', ERRMSG = '" & strDOC_MESSAGE & "'"
            strSQL = strSQL & " WHERE GBN = '" & strGBN & "' AND TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO = " & intTAXNO & " AND VOCHNO = '" & strVOCHNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_DELETEERR")
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
        MyBase.EntityName = "PD_VOCH_MST"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region

End Class








