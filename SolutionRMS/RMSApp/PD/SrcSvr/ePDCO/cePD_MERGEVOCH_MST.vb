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

Public Class cePD_MERGEVOCH_MST
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "cePD_MERGEVOCH_MST"    '자신의 클래스명
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
                             Optional ByVal strMTAXYEARMON As String = OPTIONAL_STR, _
                             Optional ByVal strMTAXNO As Double = OPTIONAL_NUM, _
                             Optional ByVal strMTAXNOSEQ As Double = OPTIONAL_NUM, _
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
                             Optional ByVal strMEDFLAG As String = OPTIONAL_STR) As Integer


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
            BuildNameValues(",", "MTAXYEARMON", strMTAXYEARMON, strFields, strValues)
            BuildNameValues(",", "MTAXNO", strMTAXNO, strFields, strValues)
            BuildNameValues(",", "MTAXNOSEQ", strMTAXNOSEQ, strFields, strValues)
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
            BuildNameValues(",", "MEDFLAG", strMEDFLAG, strFields, strValues)
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
                                     ByVal strMTAXYEARMON As String, _
                                     ByVal strMTAXNO As Integer, _
                                     ByVal strDOC_STATUS As String, _
                                     ByVal strDOC_MESSAGE As String, _
                                     ByVal strVOCHNO As String, _
                                     ByVal strIF_KEY As String, _
                                     ByVal strMEDFLAG As String) As Integer
        Dim strSQL As String
        Try

            If strMEDFLAG = "ALL" Then
                '전표업데이트 시작
                strSQL = "UPDATE PD_MERGEVOCH_MST "
                strSQL = strSQL & " SET VOCHNO = '" & strVOCHNO & "', "
                strSQL = strSQL & " ERRCODE = '" & strDOC_STATUS & "', "
                strSQL = strSQL & " ERRMSG = '" & strDOC_MESSAGE & "', "
                strSQL = strSQL & " ATTR01 = '" & strIF_KEY & "' "
                strSQL = strSQL & " WHERE GBN = '" & strGBN & "' "
                strSQL = strSQL & " AND  mtaxyearmon = '" & strMTAXYEARMON & "' "
                strSQL = strSQL & " AND mtaxno = " & strMTAXNO

            ElseIf strMEDFLAG = "AOR" Then

                strSQL = "UPDATE PD_MERGEVOCH_MST "
                strSQL = strSQL & " SET VOCHNO = '" & strVOCHNO & "', "
                strSQL = strSQL & " ERRCODE = '" & strDOC_STATUS & "', "
                strSQL = strSQL & " ERRMSG = '" & strDOC_MESSAGE & "', "
                strSQL = strSQL & " ATTR01 = '" & strIF_KEY & "' "
                strSQL = strSQL & " WHERE GBN = 'A' "
                strSQL = strSQL & " AND  mtaxyearmon = '" & strMTAXYEARMON & "' "
                strSQL = strSQL & " AND mtaxno = " & strMTAXNO
            End If

            '오류전표 아닌것만 세금계산서및 신탁 업데이트 시작
            If strDOC_STATUS = "" Then
                If strGBN = "M" Then
                    '통합청구
                    If strMEDFLAG = "ALL" Then
                        'D-옥외 'O-인터넷 'P-제작 'B,C 인쇄 'A2 케이블
                        strSQL = strSQL & " ;update a set a.vochno = b.vochno "
                        strSQL = strSQL & " from pd_mergetax_hdr a, (select mtaxyearmon,cast (cast (mtaxno as numeric)as varchar(5)) mtaxno,vochno from PD_MERGEVOCH_MST where ISNULL(errcode,'') = '' and gbn = 'M' and  mtaxyearmon = '" & strMTAXYEARMON & "' AND mtaxno = " & strMTAXNO & ") b "
                        strSQL = strSQL & " where a.mtaxyearmon = b.mtaxyearmon  "
                        strSQL = strSQL & "	and a.mtaxno = b.mtaxno and isnull(a.vochno,'') = '' "

                        strSQL = strSQL & " ;update a set a.vochno = b.vochno "
                        strSQL = strSQL & " from md_commitax_hdr a, (select mtaxyearmon,cast (cast (mtaxno as numeric)as varchar(5)) mtaxno,vochno from PD_MERGEVOCH_MST where ISNULL(errcode,'') = '' and gbn = 'M' and  mtaxyearmon = '" & strMTAXYEARMON & "' AND mtaxno = " & strMTAXNO & ") b "
                        strSQL = strSQL & " where substring(a.mtaxno,1,6) = b.mtaxyearmon "
                        strSQL = strSQL & " and substring(a.mtaxno,8,len(a.mtaxno)) = b.mtaxno "
                        strSQL = strSQL & "	and isnull(a.vochno,'') = '' "

                        strSQL = strSQL & " ;update a set a.vochno = b.vochno "
                        strSQL = strSQL & " from pd_tax_hdr a, (select mtaxyearmon,cast (cast (mtaxno as numeric)as varchar(5)) mtaxno,vochno from PD_MERGEVOCH_MST where ISNULL(errcode,'') = '' and gbn = 'M' and  mtaxyearmon = '" & strMTAXYEARMON & "' AND mtaxno = " & strMTAXNO & ") b "
                        strSQL = strSQL & " where substring(a.mtaxno,1,6) = b.mtaxyearmon "
                        strSQL = strSQL & " and substring(a.mtaxno,8,len(a.mtaxno)) = b.mtaxno"
                        strSQL = strSQL & "	and isnull(a.vochno,'') = '' "

                        strSQL = strSQL & "	;update a "
                        strSQL = strSQL & "	set a.commi_voch_no = c.vochno "
                        strSQL = strSQL & "	from md_outdoor_medium a "
                        strSQL = strSQL & "	inner join md_commitax_hdr b on a.commi_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag = 'D' "
                        strSQL = strSQL & "	and isnull(b.mtaxno,'') <> '' and isnull(a.commi_voch_no,'') = '' "
                        strSQL = strSQL & "	inner join PD_MERGEVOCH_MST c on b.mtaxno = c.mtaxyearmon +'-'+cast (cast (c.mtaxno as numeric)as varchar(10)) "
                        strSQL = strSQL & " where c.mtaxyearmon = '" & strMTAXYEARMON & "' AND c.mtaxno = " & strMTAXNO

                        strSQL = strSQL & "	;update a "
                        strSQL = strSQL & "	set a.commi_voch_no = c.vochno "
                        strSQL = strSQL & "	from md_internet_medium a "
                        strSQL = strSQL & "	inner join md_commitax_hdr b on a.commi_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag = 'O' "
                        strSQL = strSQL & "	and isnull(b.mtaxno,'') <> ''  and isnull(a.commi_voch_no,'') = '' "
                        strSQL = strSQL & "	inner join PD_MERGEVOCH_MST c on b.mtaxno = c.mtaxyearmon +'-'+cast (cast (c.mtaxno as numeric)as varchar(10)) "
                        strSQL = strSQL & " where c.mtaxyearmon = '" & strMTAXYEARMON & "' AND c.mtaxno = " & strMTAXNO

                        strSQL = strSQL & "	;update a "
                        strSQL = strSQL & "	set a.commi_voch_no = c.vochno "
                        strSQL = strSQL & "	from md_CLOUD_AMT a "
                        strSQL = strSQL & "	inner join md_commitax_hdr b on a.commi_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag = 'G' "
                        strSQL = strSQL & "	and isnull(b.mtaxno,'') <> ''  and isnull(a.commi_voch_no,'') = '' "
                        strSQL = strSQL & "	inner join PD_MERGEVOCH_MST c on b.mtaxno = c.mtaxyearmon +'-'+cast (cast (c.mtaxno as numeric)as varchar(10)) "
                        strSQL = strSQL & " where c.mtaxyearmon = '" & strMTAXYEARMON & "' AND c.mtaxno = " & strMTAXNO

                        strSQL = strSQL & "	;update a "
                        strSQL = strSQL & "	set a.VOCHNO = c.vochno "
                        strSQL = strSQL & "	from md_CLOUD_OUT a "
                        strSQL = strSQL & "	inner join md_commitax_hdr b on a.commi_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag = 'G' "
                        strSQL = strSQL & "	and isnull(b.mtaxno,'') <> ''  and isnull(a.vochno,'') = '' "
                        strSQL = strSQL & "	inner join PD_MERGEVOCH_MST c on b.mtaxno = c.mtaxyearmon +'-'+cast (cast (c.mtaxno as numeric)as varchar(10)) "
                        strSQL = strSQL & " where c.mtaxyearmon = '" & strMTAXYEARMON & "' AND c.mtaxno = " & strMTAXNO

                        strSQL = strSQL & "	;update a "
                        strSQL = strSQL & "	set a.commi_voch_no = c.vochno "
                        strSQL = strSQL & "	from md_ifcmall_AMT a "
                        strSQL = strSQL & "	inner join md_commitax_hdr b on a.commi_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag = 'Y' "
                        strSQL = strSQL & "	and isnull(b.mtaxno,'') <> ''  and isnull(a.commi_voch_no,'') = '' "
                        strSQL = strSQL & "	inner join PD_MERGEVOCH_MST c on b.mtaxno = c.mtaxyearmon +'-'+cast (cast (c.mtaxno as numeric)as varchar(10)) "
                        strSQL = strSQL & " where c.mtaxyearmon = '" & strMTAXYEARMON & "' AND c.mtaxno = " & strMTAXNO

                        strSQL = strSQL & "	;update a "
                        strSQL = strSQL & "	set a.commi_voch_no = c.vochno "
                        strSQL = strSQL & "	from MD_POINTAD_AMT a "
                        strSQL = strSQL & "	inner join md_commitax_hdr b on a.commi_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag = 'K' "
                        strSQL = strSQL & "	and isnull(b.mtaxno,'') <> ''  and isnull(a.commi_voch_no,'') = '' "
                        strSQL = strSQL & "	inner join PD_MERGEVOCH_MST c on b.mtaxno = c.mtaxyearmon +'-'+cast (cast (c.mtaxno as numeric)as varchar(10)) "
                        strSQL = strSQL & " where c.mtaxyearmon = '" & strMTAXYEARMON & "' AND c.mtaxno = " & strMTAXNO

                        strSQL = strSQL & " ;update a set a.vochno = b.vochno "
                        strSQL = strSQL & " from MD_TRUTAX_HDR a, (select mtaxyearmon,cast (cast (mtaxno as numeric)as varchar(5)) mtaxno,vochno from PD_MERGEVOCH_MST where ISNULL(errcode,'') = '' and gbn = 'M' and  mtaxyearmon = '" & strMTAXYEARMON & "' AND mtaxno = " & strMTAXNO & ") b "
                        strSQL = strSQL & " where substring(a.mtaxno,1,6) = b.mtaxyearmon "
                        strSQL = strSQL & " and substring(a.mtaxno,8,len(a.mtaxno)) = b.mtaxno "
                        strSQL = strSQL & "	and isnull(a.vochno,'') = '' "

                        strSQL = strSQL & "	;update a "
                        strSQL = strSQL & "	set a.VOCHNO = c.vochno "
                        strSQL = strSQL & "	from MD_TRUTAXGENERAL_HDR a "
                        strSQL = strSQL & "	inner join MD_TRUTAX_HDR b on a.TAXYEARMON = b.TAXYEARMON AND A.TAXNO = B.TAXNO AND b.medflag in ('A2') "
                        strSQL = strSQL & "	and isnull(b.mtaxno,'') <> ''  and isnull(a.vochno,'') = '' "
                        strSQL = strSQL & "	inner join PD_MERGEVOCH_MST c on b.mtaxno = c.mtaxyearmon +'-'+cast (cast (c.mtaxno as numeric)as varchar(10)) "
                        strSQL = strSQL & " where c.mtaxyearmon = '" & strMTAXYEARMON & "' AND c.mtaxno = " & strMTAXNO

                        strSQL = strSQL & "	;update a "
                        strSQL = strSQL & "	set a.VOCHNO = c.vochno "
                        strSQL = strSQL & "	from MD_TRUTAXOUTLIST_HDR a "
                        strSQL = strSQL & "	inner join MD_TRUTAX_HDR b on a.TAXYEARMON = b.TAXYEARMON AND A.TAXNO = B.TAXNO AND b.medflag in ('A2') "
                        strSQL = strSQL & "	and isnull(b.mtaxno,'') <> ''  and isnull(a.vochno,'') = '' "
                        strSQL = strSQL & "	inner join PD_MERGEVOCH_MST c on b.mtaxno = c.mtaxyearmon +'-'+cast (cast (c.mtaxno as numeric)as varchar(10)) "
                        strSQL = strSQL & " where c.mtaxyearmon = '" & strMTAXYEARMON & "' AND c.mtaxno = " & strMTAXNO

                        strSQL = strSQL & "	;update a "
                        strSQL = strSQL & "	set a.VOCHNO = c.vochno "
                        strSQL = strSQL & "	from MD_TRUTAXGENERAL_HDR a "
                        strSQL = strSQL & "	inner join MD_TRUTAX_HDR b on a.TAXYEARMON = b.TAXYEARMON AND A.TAXNO = B.TAXNO AND b.medflag in ('B','C') "
                        strSQL = strSQL & "	and isnull(b.mtaxno,'') <> ''  and isnull(a.vochno,'') = '' "
                        strSQL = strSQL & "	inner join PD_MERGEVOCH_MST c on b.mtaxno = c.mtaxyearmon +'-'+cast (cast (c.mtaxno as numeric)as varchar(10)) "
                        strSQL = strSQL & " where c.mtaxyearmon = '" & strMTAXYEARMON & "' AND c.mtaxno = " & strMTAXNO

                        strSQL = strSQL & "	;update a "
                        strSQL = strSQL & "	set a.VOCHNO = c.vochno "
                        strSQL = strSQL & "	from MD_TRUTAXOUTLIST_HDR a "
                        strSQL = strSQL & "	inner join MD_TRUTAX_HDR b on a.TAXYEARMON = b.TAXYEARMON AND A.TAXNO = B.TAXNO AND b.medflag in ('B','C') "
                        strSQL = strSQL & "	and isnull(b.mtaxno,'') <> ''  and isnull(a.vochno,'') = '' "
                        strSQL = strSQL & "	inner join PD_MERGEVOCH_MST c on b.mtaxno = c.mtaxyearmon +'-'+cast (cast (c.mtaxno as numeric)as varchar(10)) "
                        strSQL = strSQL & " where c.mtaxyearmon = '" & strMTAXYEARMON & "' AND c.mtaxno = " & strMTAXNO

                        strSQL = strSQL & "	;update a "
                        strSQL = strSQL & "	set a.tru_voch_no = c.vochno "
                        strSQL = strSQL & "	from md_booking_medium a "
                        strSQL = strSQL & "	inner join md_trutax_hdr b on a.tru_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag in ('B','C') "
                        strSQL = strSQL & "	and isnull(b.mtaxno,'') <> ''  and isnull(a.commi_voch_no,'') = '' "
                        strSQL = strSQL & "	inner join PD_MERGEVOCH_MST c on b.mtaxno = c.mtaxyearmon +'-'+cast (cast (c.mtaxno as numeric)as varchar(10)) "
                        strSQL = strSQL & " where c.mtaxyearmon = '" & strMTAXYEARMON & "' AND c.mtaxno = " & strMTAXNO

                        strSQL = strSQL & "	;update a "
                        strSQL = strSQL & "	set a.tru_voch_no = c.vochno "
                        strSQL = strSQL & "	from md_catv_medium a "
                        strSQL = strSQL & "	inner join md_trutax_hdr b on a.tru_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag = 'A2' "
                        strSQL = strSQL & "	and isnull(b.mtaxno,'') <> ''  and isnull(a.commi_voch_no,'') = '' "
                        strSQL = strSQL & "	inner join PD_MERGEVOCH_MST c on b.mtaxno = c.mtaxyearmon +'-'+cast (cast (c.mtaxno as numeric)as varchar(10)) "
                        strSQL = strSQL & " where c.mtaxyearmon = '" & strMTAXYEARMON & "' AND c.mtaxno = " & strMTAXNO

                        'AOR 대행매출 전표
                    ElseIf strMEDFLAG = "AOR" Then

                        strSQL = strSQL & " ;UPDATE A "
                        strSQL = strSQL & " SET A.VOCHNO = B.VOCHNO "
                        strSQL = strSQL & " FROM PD_MERGETAX_HDR A, "
                        strSQL = strSQL & " 	("
                        strSQL = strSQL & " 		SELECT MTAXYEARMON,"
                        strSQL = strSQL & " 		CAST (CAST (MTAXNO AS NUMERIC)AS VARCHAR(5)) MTAXNO,VOCHNO "
                        strSQL = strSQL & " 		FROM PD_MERGEVOCH_MST "
                        strSQL = strSQL & " 		WHERE ISNULL(ERRCODE,'') = '' "
                        strSQL = strSQL & " 		AND GBN = 'A' "
                        strSQL = strSQL & " 		AND MTAXYEARMON = '" & strMTAXYEARMON & "' "
                        strSQL = strSQL & " 		AND MTAXNO =  '" & strMTAXNO & "' "
                        strSQL = strSQL & " 	) B "
                        strSQL = strSQL & " WHERE A.MTAXYEARMON = B.MTAXYEARMON  "
                        strSQL = strSQL & " AND A.MTAXNO = B.MTAXNO AND ISNULL(A.VOCHNO,'') = '' "

                        strSQL = strSQL & " ;UPDATE A "
                        strSQL = strSQL & " SET A.VOCHNO = B.VOCHNO "
                        strSQL = strSQL & " FROM MD_COMMITAX_HDR A, "
                        strSQL = strSQL & " 	("
                        strSQL = strSQL & " 	SELECT MTAXYEARMON,"
                        strSQL = strSQL & " 	CAST (CAST (MTAXNO AS NUMERIC)AS VARCHAR(5)) MTAXNO,VOCHNO "
                        strSQL = strSQL & " 	FROM PD_MERGEVOCH_MST "
                        strSQL = strSQL & " 	WHERE ISNULL(ERRCODE,'') = '' "
                        strSQL = strSQL & " 	AND GBN = 'A' "
                        strSQL = strSQL & " 	AND MTAXYEARMON = '" & strMTAXYEARMON & "' "
                        strSQL = strSQL & " 	AND MTAXNO =  '" & strMTAXNO & "' "
                        strSQL = strSQL & " ) B "
                        strSQL = strSQL & " WHERE SUBSTRING(A.MTAXNO,1,6) = B.MTAXYEARMON "
                        strSQL = strSQL & " AND SUBSTRING(A.MTAXNO,8,LEN(A.MTAXNO)) = B.MTAXNO "
                        strSQL = strSQL & " AND ISNULL(A.VOCHNO,'') = ''"

                        strSQL = strSQL & " ;UPDATE A "
                        strSQL = strSQL & " SET A.COMMI_VOCH_NO = C.VOCHNO "
                        strSQL = strSQL & " FROM MD_AOR_MEDIUM A "
                        strSQL = strSQL & " INNER JOIN MD_COMMITAX_HDR B ON A.COMMI_TAX_NO = B.TAXYEARMON +'-'+CAST(B.TAXNO AS VARCHAR(10)) AND ISNULL(B.ATTR03,'') = 'AOR' "
                        strSQL = strSQL & " AND ISNULL(B.MTAXNO,'') <> '' "
                        strSQL = strSQL & " AND ISNULL(A.COMMI_VOCH_NO,'') = '' "
                        strSQL = strSQL & " INNER JOIN PD_MERGEVOCH_MST C ON B.MTAXNO = C.MTAXYEARMON +'-'+CAST (CAST (C.MTAXNO AS NUMERIC)AS VARCHAR(10)) "
                        strSQL = strSQL & " WHERE C.MTAXYEARMON = '" & strMTAXYEARMON & "' AND C.MTAXNO =  '" & strMTAXNO & "'"

                    End If

                End If
            End If

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_VOCHNO")
        End Try
    End Function



    Public Function Update_vochno(ByVal strMTAXYEARMON As String, _
                                  ByVal strMTAXNO As Double, _
                                  ByVal strMEDFLAG As String) As Integer
        Dim strSQL As String
        Dim strMTAX As String

        Try

            strMTAX = strMTAXYEARMON & "-" & strMTAXNO

            '통합청구
            If strMEDFLAG = "ALL" Then
                '인터넷
                strSQL = " update a set a.commi_voch_no = '' "
                strSQL = strSQL & " from md_internet_medium a "
                strSQL = strSQL & " inner join md_commitax_hdr b on a.commi_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag = 'O' "
                strSQL = strSQL & " where isnull(b.mtaxno,'') = '" & strMTAX & "' "

                '옥외
                strSQL = strSQL & " ;update a set a.commi_voch_no = '' "
                strSQL = strSQL & " from md_outdoor_medium a "
                strSQL = strSQL & " inner join md_commitax_hdr b on a.commi_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag = 'D' "
                strSQL = strSQL & " where isnull(b.mtaxno,'') = '" & strMTAX & "' "

                'CGV신탁
                strSQL = strSQL & " ;update a set a.commi_voch_no = '' "
                strSQL = strSQL & " from md_CLOUD_AMT a "
                strSQL = strSQL & " inner join md_commitax_hdr b on a.commi_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag = 'G' "
                strSQL = strSQL & " where isnull(b.mtaxno,'') = '" & strMTAX & "' "

                'CGV대행사/내부대행사
                strSQL = strSQL & " ;update a set a.VOCHNO = '' "
                strSQL = strSQL & " from md_CLOUD_OUT a "
                strSQL = strSQL & " inner join md_commitax_hdr b on a.commi_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag = 'G' "
                strSQL = strSQL & " where isnull(b.mtaxno,'') = '" & strMTAX & "' "

                'IFCMALL 신탁
                strSQL = strSQL & " ;update a set a.commi_voch_no = '' "
                strSQL = strSQL & " from MD_IFCMALL_AMT A "
                strSQL = strSQL & " inner join md_commitax_hdr b on a.commi_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag = 'Y' "
                strSQL = strSQL & " where isnull(b.mtaxno,'') = '" & strMTAX & "' "

                'POINT AD 신탁
                strSQL = strSQL & " ;update a set a.commi_voch_no = '' "
                strSQL = strSQL & " from MD_POINTAD_AMT A "
                strSQL = strSQL & " inner join md_commitax_hdr b on a.commi_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag = 'K' "
                strSQL = strSQL & " where isnull(b.mtaxno,'') = '" & strMTAX & "' "

                'PD는 신탁이없다.

                '인쇄 
                strSQL = strSQL & " ;update a set a.tru_voch_no = '' "
                strSQL = strSQL & " from md_booking_medium a "
                strSQL = strSQL & " inner join md_trutax_hdr b on a.tru_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag in ('B','C') "
                strSQL = strSQL & " where isnull(b.mtaxno,'') = '" & strMTAX & "' "

                '케이블
                strSQL = strSQL & " ;update a set a.tru_voch_no = '' "
                strSQL = strSQL & " from md_internet_medium a "
                strSQL = strSQL & " inner join md_trutax_hdr b on a.tru_tax_no = b.taxyearmon +'-'+cast(b.taxno as varchar(10)) and b.medflag in ('B','C') "
                strSQL = strSQL & " where isnull(b.mtaxno,'') = '" & strMTAX & "' "

                '인쇄 케이블 (일반)
                strSQL = strSQL & " ;update a set a.vochno = '' "
                strSQL = strSQL & " from MD_TRUTAXGENERAL_HDR a "
                strSQL = strSQL & " inner join md_trutax_hdr b on a.taxyearmon = b.taxyearmon and a.taxno = b.taxno and b.medflag in ('B','C','A2') "
                strSQL = strSQL & " where isnull(b.mtaxno,'') = '" & strMTAX & "' "

                strSQL = strSQL & " ;update a set a.vochno = '' "
                strSQL = strSQL & " from MD_TRUTAXOUTLIST_HDR a "
                strSQL = strSQL & " inner join md_trutax_hdr b on a.taxyearmon = b.taxyearmon and a.taxno = b.taxno and b.medflag in ('B','C','A2') "
                strSQL = strSQL & " where isnull(b.mtaxno,'') = '" & strMTAX & "' "

                'AOR 대행매출
            ElseIf strMEDFLAG = "AOR" Then

                strSQL = strSQL & " ;UPDATE A "
                strSQL = strSQL & " SET A.COMMI_VOCH_NO = '' "
                strSQL = strSQL & " FROM MD_AOR_MEDIUM A "
                strSQL = strSQL & " INNER JOIN MD_COMMITAX_HDR B ON A.COMMI_TAX_NO = B.TAXYEARMON +'-'+CAST(B.TAXNO AS VARCHAR(10)) "
                strSQL = strSQL & " AND ISNULL(B.ATTR03,'') = 'AOR' "
                strSQL = strSQL & " WHERE ISNULL(B.MTAXNO,'') = '" & strMTAX & "' "


            End If


            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_vochno")
        End Try
    End Function

    Public Function UpdateDelete(ByVal strMTAXYEARMON As String, _
                                 ByVal strMTAXNO As Double, _
                                 ByVal strMEDFLAG As String) As Integer
        Dim strSQL As String
        Try

            If strMEDFLAG = "ALL" Then
                '제작
                strSQL = "UPDATE PD_TAX_HDR SET VOCHNO ='' WHERE MTAXNO = '" & strMTAXYEARMON & "-" & strMTAXNO & "'"
                '인터넷 ,옥외
                strSQL = strSQL & ";UPDATE MD_COMMITAX_HDR "
                strSQL = strSQL & " SET VOCHNO ='' "
                strSQL = strSQL & " WHERE MTAXNO = '" & strMTAXYEARMON & "-" & strMTAXNO & "'"
                strSQL = strSQL & " AND ISNULL(ATTR03,'') <> 'AOR' "

                '인쇄 , 케이블
                strSQL = strSQL & ";UPDATE MD_TRUTAX_HDR "
                strSQL = strSQL & " SET VOCHNO ='' "
                strSQL = strSQL & " WHERE MTAXNO = '" & strMTAXYEARMON & "-" & strMTAXNO & "'"

                '통합청구
                strSQL = strSQL & ";UPDATE PD_MERGETAX_HDR "
                strSQL = strSQL & " SET VOCHNO ='' WHERE MTAXYEARMON = '" & strMTAXYEARMON & "' "
                strSQL = strSQL & " AND MTAXNO = '" & strMTAXNO & "'"

            ElseIf strMEDFLAG = "AOR" Then

                'AOR 대행매출
                strSQL = strSQL & " ;UPDATE MD_COMMITAX_HDR "
                strSQL = strSQL & " SET VOCHNO ='' "
                strSQL = strSQL & " WHERE MTAXNO = '" & strMTAXYEARMON & "-" & strMTAXNO & "' "
                strSQL = strSQL & " AND ISNULL(ATTR03,'') = 'AOR' "

                strSQL = strSQL & " ;UPDATE PD_MERGETAX_HDR "
                strSQL = strSQL & " SET VOCHNO ='' "
                strSQL = strSQL & " WHERE MTAXYEARMON = '" & strMTAXYEARMON & "' "
                strSQL = strSQL & " AND MTAXNO = '" & strMTAXNO & "'"


            End If


            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDelete")
        End Try
    End Function


    Public Function Delete_vochno(ByVal strMTAXYEARMON As String, _
                                  ByVal strVOCHNO As String) As Integer
        Dim strSQL As String
        Try

            strSQL = "DELETE FROM  PD_MERGEVOCH_MST "
            strSQL = strSQL & " WHERE MTAXYEARMON ='" & strMTAXYEARMON & "' "
            strSQL = strSQL & " AND VOCHNO = '" & strVOCHNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Delete_vochno")
        End Try
    End Function

    Public Function Delete_ERR(ByVal strMTAXYEARMON As String, _
                               ByVal strMTAXNO As Double) As Integer
        Dim strSQL As String
        Try
            strSQL = " DELETE FROM PD_MERGEVOCH_MST "
            strSQL = strSQL & " WHERE MTAXYEARMON ='" & strMTAXYEARMON & "' "
            strSQL = strSQL & " AND MTAXNO = " & strMTAXNO & ""

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Delete_ERR")
        End Try
    End Function

    Public Function UpdateRtn_DELETEERR(ByVal strGBN As String, _
                                        ByVal strTAXYEARMON As String, _
                                        ByVal intTAXNO As Integer, _
                                        ByVal strDOC_STATUS As String, _
                                        ByVal strDOC_MESSAGE As String, _
                                        ByVal strVOCHNO As String, _
                                        ByVal strMEDFLAG As String) As Integer
        Dim strSQL As String
        Try

            If strMEDFLAG = "ALL" Then
                '전표업데이트 시작
                strSQL = "UPDATE PD_MERGEVOCH_MST "
                strSQL = strSQL & " SET ERRCODE = '" & strDOC_STATUS & "', ERRMSG = '" & strDOC_MESSAGE & "'"
                strSQL = strSQL & " WHERE GBN = 'M' "
                strSQL = strSQL & " AND MTAXYEARMON = '" & strTAXYEARMON & "' "
                strSQL = strSQL & " AND MTAXNO = " & intTAXNO & " "
                strSQL = strSQL & " AND VOCHNO = '" & strVOCHNO & "'"

            ElseIf strMEDFLAG = "AOR" Then

                strSQL = "UPDATE PD_MERGEVOCH_MST "
                strSQL = strSQL & " SET ERRCODE = '" & strDOC_STATUS & "', ERRMSG = '" & strDOC_MESSAGE & "'"
                strSQL = strSQL & " WHERE GBN = 'A' "
                strSQL = strSQL & " AND MTAXYEARMON = '" & strTAXYEARMON & "' "
                strSQL = strSQL & " AND MTAXNO = " & intTAXNO & " "
                strSQL = strSQL & " AND VOCHNO = '" & strVOCHNO & "'"

            End If

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
        MyBase.EntityName = "PD_MERGEVOCH_MST"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region

End Class








