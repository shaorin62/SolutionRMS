'****************************************************************************************
'시스템구분    : 솔루션명 /시스템명/Server Control Class
'실행   환경    : COM+ Service Server Package
'프로그램명    : ccMDCMCUST_TRAN.vb
'기         능    : - 기능을 명시 합니다.
'특이  사항     : - 특이사항에 대해 표현
'                     -
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009-07-03 오전 10:32:13 By KTY
'****************************************************************************************

Imports System.Xml                  ' XML처리
Imports SCGLControl                 ' ControlClass의 Base Class 
Imports SCGLUtil.cbSCGLConfig       ' ConfigurationClass
Imports SCGLUtil.cbSCGLErr          '오류처리 클래스
Imports SCGLUtil.cbSCGLXml          'XML처리 클래스
Imports SCGLUtil.cbSCGLUtil         '기타유틸리티 클래스
Imports eSCCO '엔터티 추가

' 엔티티 클래스 사용시 해당 엔티티 클래스의 프로젝트를 참조한 후 Imports 하십시요. 
' Imports 엔티티프로젝트

Public Class ccSCCOCUSTLIST
    Inherits ccControl

#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccSCCOCUSTLIST"                  '자신의 클래스명
    Private mobjceSC_CUST_DTL As eSCCO.ceSC_CUST_DTL            '사용할 Entity 변수 선언
    Private mobjceSC_CUST_HDR As eSCCO.ceSC_CUST_HDR             '사용할 Entity 변수 선언
    Private mobjceSC_CUST_SAP As eSCCO.ceSC_CUST_SAP             '사용할 Entity 변수 선언
    Private mobjceSC_CUST_SAPBANK As eSCCO.ceSC_CUST_SAPBANK             '사용할 Entity 변수 선언

#End Region

#Region "GROUP BLOCK : Property 선언"
#End Region

#Region "GROUP BLOCK : Event 선언"
    Public Function Busino_Check(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, _
                                 ByRef intColCnt As Integer, _
                                 ByVal strBUSINO As String, _
                                 ByVal strMEDFLAG As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = "" : Con2 = ""

                If strBUSINO <> "" Then Con1 = String.Format(" AND (Ltrim(Rtrim(Replace(BUSINO,'-',''))) = '{0}')", strBUSINO)
                If strMEDFLAG <> "" Then Con2 = String.Format(" AND (MEDFLAG = '{0}')", strMEDFLAG)

                strWhere = BuildFields(" ", Con1, Con2)

                strFormet = "SELECT BUSINO FROM SC_CUST_HDR WHERE 1=1 {0}"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function SelectRtn_CountCheck(ByVal strInfoXML As String, _
                                         ByRef intRowCnt As Integer, _
                                         ByRef intColCnt As Integer, _
                                         ByVal strCUSTCODE As String, _
                                         ByVal strMEDFLAG As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormat, strSelFields, strWhere As String
        Dim Con1 As String
        Dim vntData As Object
        Dim strField

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""

                If strMEDFLAG = "G" Then        '대행사
                    If strCUSTCODE <> "" Then Con1 = String.Format(" AND (EXClIENTCODE = '{0}')", strCUSTCODE)
                    strField = "EXCLIENTCODE"
                ElseIf strMEDFLAG = "B" Then    '매체
                    If strCUSTCODE <> "" Then Con1 = String.Format(" AND (MEDCODE = '{0}')", strCUSTCODE)
                    strField = "MEDCODE"
                ElseIf strMEDFLAG = "K" Then    '크리조직
                    If strCUSTCODE <> "" Then Con1 = String.Format(" AND (EXCLIENTCODE = '{0}')", strCUSTCODE)
                    strField = "EXCLIENTCODE"
                ElseIf strMEDFLAG = "P" Then    'MPP
                    If strCUSTCODE <> "" Then Con1 = String.Format(" AND (MPP = '{0}')", strCUSTCODE)
                    strField = "MPP"
                ElseIf strMEDFLAG = "R" Then    '매체사
                    If strCUSTCODE <> "" Then Con1 = String.Format(" AND (REAL_MED_CODE = '{0}')", strCUSTCODE)
                    strField = "REAL_MED_CODE"
                End If


                strFormat = strFormat & "  SELECT MEDFLAG, COUNT(*) FROM ("
                strFormat = strFormat & "  	SELECT 'B' MEDFLAG, " & strField & " FROM MD_BOOKING_MEDIUM"
                strFormat = strFormat & "  	WHERE 1=1 {0}"
                strFormat = strFormat & "  	UNION ALL"
                strFormat = strFormat & "  	SELECT 'A2' MEDFLAG, " & strField & " FROM MD_CATV_MEDIUM"
                strFormat = strFormat & "  	WHERE 1=1 {0}"
                strFormat = strFormat & "  	UNION ALL"
                strFormat = strFormat & "  	SELECT 'A' MEDFLAG, " & strField & " FROM MD_ELECTRIC_MEDIUM"
                strFormat = strFormat & "  	WHERE 1=1 {0}"
                If strMEDFLAG <> "K" Then
                    strFormat = strFormat & "  	UNION ALL"
                    strFormat = strFormat & "  	SELECT 'O' MEDFLAG, " & strField & " FROM MD_INTERNET_MEDIUM"
                    strFormat = strFormat & "  	WHERE 1=1 {0}"
                End If
                If strMEDFLAG = "B" Or strMEDFLAG = "R" Then
                    strFormat = strFormat & "  	UNION ALL"
                    strFormat = strFormat & "  	SELECT 'D' MEDFLAG, " & strField & " FROM MD_OUTDOOR_MEDIUM"
                    strFormat = strFormat & "  	WHERE 1=1 {0}"
                End If

                strFormat = strFormat & "  ) AAA"
                strFormat = strFormat & "  GROUP BY MEDFLAG"


                strWhere = BuildFields(" ", Con1)

                strSQL = String.Format(strFormat, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function SelectRtn_PDCountCheck(ByVal strInfoXML As String, _
                                           ByRef intRowCnt As Integer, _
                                           ByRef intColCnt As Integer, _
                                           ByVal strCUSTCODE As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormat, strSelFields, strWhere As String
        Dim Con1 As String
        Dim vntData As Object
        Dim strField

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (OUTSCODE = '{0}')", strCUSTCODE)

                strFormat = strFormat & "  	SELECT COUNT(*) FROM PD_EXE_DTL"
                strFormat = strFormat & "  	WHERE 1=1 {0} GROUP BY OUTSCODE"

                strWhere = BuildFields(" ", Con1)

                strSQL = String.Format(strFormat, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_PDCountCheck")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function SelectRtn_MPPCountCheck(ByVal strInfoXML As String, _
                                            ByRef intRowCnt As Integer, _
                                            ByRef intColCnt As Integer, _
                                            ByVal strCUSTCODE As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormat, strSelFields, strWhere As String
        Dim Con1 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (MPP = '{0}')", strCUSTCODE)

                strFormat = strFormat & "  	SELECT COUNT(*) FROM MD_CATV_MEDIUM"
                strFormat = strFormat & "  	WHERE 1=1 {0} GROUP BY MPP"

                strWhere = BuildFields(" ", Con1)

                strSQL = String.Format(strFormat, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function SelectRtn_MEDCountCheck(ByVal strInfoXML As String, _
                                            ByRef intRowCnt As Integer, _
                                            ByRef intColCnt As Integer, _
                                            ByVal strHIGHCUSTCODE As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormat, strSelFields, strWhere As String
        Dim Con1 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""

                If strHIGHCUSTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE = '{0}')", strHIGHCUSTCODE)

                strFormat = strFormat & "  	SELECT COUNT(*) FROM SC_CUST_DTL"
                strFormat = strFormat & "  	WHERE 1=1 AND MEDFLAG = 'B' {0} GROUP BY HIGHCUSTCODE"

                strWhere = BuildFields(" ", Con1)

                strSQL = String.Format(strFormat, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : 외부에 공개 Method"
    ' =============== SelectRtn_CUSTHDR 광고주 헤더
    Public Function SelectRtn_CUSTHDR(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, _
                                      ByRef intColCnt As Integer, _
                                      ByVal strCUSTNAME As String, _
                                      ByVal strCOMPANYNAME As String, _
                                      ByVal strBUSINO As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""
                Con2 = ""
                Con3 = ""

                If strCUSTNAME <> "" Then Con1 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                If strCOMPANYNAME <> "" Then Con2 = String.Format(" AND (COMPANYNAME LIKE '%{0}%')", strCOMPANYNAME)
                If strBUSINO <> "" Then Con3 = String.Format(" AND (BUSINO LIKE '%{0}%')", strBUSINO)

                strWhere = BuildFields(" ", Con1, Con2, Con3)
                strSelFields = " BUSINO ,COMPANYNAME,CUSTNAME,HIGHCUSTCODE, CUSTOWNER , "
                strSelFields = strSelFields & " USE_FLAG, "
                strSelFields = strSelFields & " CASE CUSTTYPE WHEN '2' THEN '계열' ELSE '비계열' END AS CUSTTYPE, "
                strSelFields = strSelFields & " BUSISTAT,BUSITYPE, "
                strSelFields = strSelFields & " case len(isnull(ZIPCODE,'')) when 6 then  "
                strSelFields = strSelFields & " substring(isnull(ZIPCODE,''),1,3) + '-' + substring(isnull(ZIPCODE,''),4,3) else isnull(ZIPCODE,'') end as ZIPCODE, "
                strSelFields = strSelFields & " ADDRESS1, ADDRESS2, "
                strSelFields = strSelFields & " TEL, FAX,"
                strSelFields = strSelFields & " MEMO, "
                strSelFields = strSelFields & " KOBACOCUSTCODE, "
                strSelFields = strSelFields & " GREATCODE, DBO.SC_GET_GREATCUSTNAME_FUN(GREATCODE) GREATNAME, ISNULL(ATTR10,0) ATTR10, ISNULL(ATTR08,0) ATTR08,"
                strSelFields = strSelFields & " DBO.SC_EMPNAME_FUN(UUSER) UUSER "

                strFormet = "select {0} from SC_CUST_HDR where 1=1 AND MEDFLAG = 'A' {1} "

                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CUSTHDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function


    ' =============== SelectRtn_CUSTDTL 광고주디테일
    Public Function SelectRtn_CUSTDTL(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, _
                                      ByRef intColCnt As Integer, _
                                      ByRef strHIGHCUSTCODE As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""
                Con2 = ""

                If strHIGHCUSTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE = '{0}')", strHIGHCUSTCODE)

                strWhere = BuildFields(" ", Con1)

                strSelFields = " CASE GBNFLAG WHEN '0' THEN '팀' WHEN '1' THEN 'CIC/사업부' ELSE '' END GBNFLAG, "
                strSelFields = strSelFields & " CLIENTSUBCODE, '' BTN, DBO.SC_GET_CUSTNAME_FUN(CLIENTSUBCODE) AS CLIENTSUBNAME, "
                strSelFields = strSelFields & " CUSTNAME, CUSTCODE, "
                strSelFields = strSelFields & " HIGHCUSTCODE, '' BTNHIGH, DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) COMPANYNAME,"
                strSelFields = strSelFields & " USE_FLAG "

                strFormet = "select {0} from SC_CUST_DTL where 1=1 AND MEDFLAG = 'A' {1} "

                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CUSTDTL")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' =============== SelectRtn_MEDHDR 매체사 헤더 
    Public Function SelectRtn_MEDHDR(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByVal strMEDDIV As String, _
                                     ByVal strREAL_MED_NAME As String, _
                                     ByVal strMEDNAME As String, _
                                     ByVal strBUSINO As String) As Object

        Dim strSQL As String
        Dim Con1, Con2, Con3, Con4 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""
                Con2 = ""
                Con3 = ""
                Con4 = ""

                If strREAL_MED_NAME <> "" Then Con1 = String.Format(" AND (COMPANYNAME LIKE '%{0}%')", strREAL_MED_NAME)
                If strMEDNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strMEDNAME)
                If strBUSINO <> "" Then Con3 = String.Format(" AND (BUSINO LIKE '%{0}%')", strBUSINO)

                If strMEDDIV <> "" Then
                    If strMEDDIV = "MED_PAP" Then
                        Con4 = " AND ( MED_PAP = '1' OR MED_MAG = '1') "
                    Else
                        Con4 = String.Format(" AND ( {0} = '1')", strMEDDIV)
                    End If

                End If

                If strMEDDIV <> "" Then
                    strSQL = "  SELECT "
                    strSQL = strSQL & "  0 CHK, BUSINO ,COMPANYNAME,CUSTNAME,HIGHCUSTCODE, CUSTOWNER, "
                    strSQL = strSQL & "  USE_FLAG,"
                    strSQL = strSQL & "  CASE CUSTTYPE WHEN '2' THEN '계열' ELSE '비계열' END AS CUSTTYPE, "
                    strSQL = strSQL & "  BUSISTAT,BUSITYPE, "
                    strSQL = strSQL & "  ZIPCODE, ADDRESS1, ADDRESS2, "
                    strSQL = strSQL & "  TEL, FAX, "
                    strSQL = strSQL & "  MEMO, "
                    strSQL = strSQL & "  DBO.SC_EMPNAME_FUN(UUSER) UUSER "
                    strSQL = strSQL & "  FROM SC_CUST_HDR"
                    strSQL = strSQL & "  WHERE 1=1 AND MEDFLAG = 'B' "
                    strSQL = strSQL & "  " & Con1
                    strSQL = strSQL & "  " & Con3
                    strSQL = strSQL & "  AND HIGHCUSTCODE IN("
                    strSQL = strSQL & "       SELECT HIGHCUSTCODE FROM SC_CUST_DTL "
                    strSQL = strSQL & "       WHERE MEDFLAG = 'B' "
                    strSQL = strSQL & "       " & Con2
                    strSQL = strSQL & "       " & Con4
                    strSQL = strSQL & "       GROUP BY HIGHCUSTCODE )"
                Else
                    strSQL = "  SELECT "
                    strSQL = strSQL & "  0 CHK, BUSINO ,COMPANYNAME,CUSTNAME,HIGHCUSTCODE, CUSTOWNER, "
                    strSQL = strSQL & "  USE_FLAG,"
                    strSQL = strSQL & "  CASE CUSTTYPE WHEN '2' THEN '계열' ELSE '비계열' END AS CUSTTYPE, "
                    strSQL = strSQL & "  BUSISTAT,BUSITYPE, "
                    strSQL = strSQL & "  ZIPCODE, ADDRESS1, ADDRESS2, "
                    strSQL = strSQL & "  TEL, FAX, "
                    strSQL = strSQL & "  MEMO, "
                    strSQL = strSQL & "  DBO.SC_EMPNAME_FUN(UUSER) UUSER "
                    strSQL = strSQL & "  FROM SC_CUST_HDR"
                    strSQL = strSQL & "  WHERE 1=1 AND MEDFLAG = 'B' "
                    strSQL = strSQL & "  " & Con1
                    strSQL = strSQL & "  " & Con3

                End If



                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CUSTHDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' =============== SelectRtn_MEDDTL 매체사디테일
    Public Function SelectRtn_MEDDTL(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByRef strHIGHCUSTCODE As String, _
                                     ByRef strMEDNAME As String) As Object     'XML  데이터 조회시 

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""
                Con2 = ""

                If strHIGHCUSTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE = '{0}')", strHIGHCUSTCODE)
                If strMEDNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strMEDNAME)

                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = " 0 CHK, "
                strSelFields = strSelFields & " CUSTNAME, CUSTCODE, "
                strSelFields = strSelFields & " HIGHCUSTCODE, '' BTNHIGH, DBO.SC_GET_HIGHCOMPANYNAME_FUN(HIGHCUSTCODE) COMPANYNAME,"
                strSelFields = strSelFields & " MED_TV, MED_RD, "
                strSelFields = strSelFields & " MED_DMB, MED_CATV,MED_GEN, "
                strSelFields = strSelFields & " MED_PAP, MED_MAG, "
                strSelFields = strSelFields & " MED_NET, MED_OUT, MED_ETC,"
                strSelFields = strSelFields & " MPP, '' BTNMPP, "
                strSelFields = strSelFields & " DBO.SC_GET_HIGHCUSTNAME_FUN(MPP) MPPNAME, "
                strSelFields = strSelFields & " USE_FLAG "

                strFormet = "select {0} from SC_CUST_DTL where 1=1 AND MEDFLAG = 'B' {1} "

                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CUSTDTL")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' =============== SelectRtn_EXEHDR 대대행사 헤더
    Public Function SelectRtn_EXEHDR(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByVal strCUSTNAME As String, _
                                     ByVal strCOMPANYNAME As String, _
                                     ByVal strBUSINO As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""
                Con2 = ""
                Con3 = ""

                If strCUSTNAME <> "" Then Con1 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                If strCOMPANYNAME <> "" Then Con2 = String.Format(" AND (COMPANYNAME LIKE '%{0}%')", strCOMPANYNAME)
                If strBUSINO <> "" Then Con3 = String.Format(" AND (BUSINO LIKE '%{0}%')", strBUSINO)

                strWhere = BuildFields(" ", Con1, Con2, Con3)
                strSelFields = " 0 CHK, BUSINO ,COMPANYNAME,CUSTNAME,HIGHCUSTCODE, CUSTOWNER , "
                strSelFields = strSelFields & " USE_FLAG, "
                strSelFields = strSelFields & " CASE CUSTTYPE WHEN '2' THEN '계열' ELSE '비계열' END AS CUSTTYPE, "
                strSelFields = strSelFields & " BUSISTAT,BUSITYPE, "
                strSelFields = strSelFields & " ZIPCODE, ADDRESS1, ADDRESS2, "
                strSelFields = strSelFields & " TEL, FAX,"
                strSelFields = strSelFields & " MEMO, "
                strSelFields = strSelFields & " DBO.SC_EMPNAME_FUN(UUSER) UUSER "

                strFormet = "select {0} from SC_CUST_HDR where 1=1 AND MEDFLAG = 'G' {1} "

                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CUSTHDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' =============== SelectRtn_OUTHDR 외주처 헤더
    Public Function SelectRtn_OUTHDR(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByVal strCUSTNAME As String, _
                                     ByVal strCOMPANYNAME As String, _
                                     ByVal strBUSINO As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""
                Con2 = ""
                Con3 = ""

                If strCUSTNAME <> "" Then Con1 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                If strCOMPANYNAME <> "" Then Con2 = String.Format(" AND (COMPANYNAME LIKE '%{0}%')", strCOMPANYNAME)
                If strBUSINO <> "" Then Con3 = String.Format(" AND (BUSINO LIKE '%{0}%')", strBUSINO)

                strWhere = BuildFields(" ", Con1, Con2, Con3)
                strSelFields = " 0 CHK, BUSINO ,COMPANYNAME,CUSTNAME,HIGHCUSTCODE, CUSTOWNER , "
                strSelFields = strSelFields & " USE_FLAG, "
                strSelFields = strSelFields & " CASE CUSTTYPE WHEN '2' THEN '계열' ELSE '비계열' END AS CUSTTYPE, "
                strSelFields = strSelFields & " BUSISTAT,BUSITYPE, "
                strSelFields = strSelFields & " ZIPCODE, ADDRESS1, ADDRESS2, "
                strSelFields = strSelFields & " TEL, FAX,"
                strSelFields = strSelFields & " MEMO, "
                strSelFields = strSelFields & " DBO.SC_EMPNAME_FUN(UUSER) UUSER "

                strFormet = "select {0} from SC_CUST_HDR where 1=1 AND MEDFLAG = 'M' {1} "

                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CUSTHDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' =============== SelectRtn_MPPHDR MPP 헤더
    Public Function SelectRtn_MPPHDR(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByVal strCUSTNAME As String, _
                                     ByVal strBUSINO As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""
                Con2 = ""

                If strCUSTNAME <> "" Then Con1 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                If strBUSINO <> "" Then Con2 = String.Format(" AND (BUSINO LIKE '%{0}%')", strBUSINO)

                strWhere = BuildFields(" ", Con1, Con2)
                strSelFields = " 0 CHK, BUSINO , CUSTNAME,HIGHCUSTCODE, "
                strSelFields = strSelFields & " USE_FLAG, "
                strSelFields = strSelFields & " CASE CUSTTYPE WHEN '2' THEN '계열' ELSE '비계열' END AS CUSTTYPE,"
                strSelFields = strSelFields & " DBO.SC_EMPNAME_FUN(UUSER) UUSER"

                strFormet = "select {0} from SC_CUST_HDR where 1=1 AND MEDFLAG = 'P' {1} "

                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CUSTHDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' =============== SelectRtn_CREHDR 크리조직 헤더
    Public Function SelectRtn_CREHDR(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim Con1 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""

                If strCUSTNAME <> "" Then Con1 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)

                strWhere = BuildFields(" ", Con1)
                strSelFields = " 0 CHK, BUSINO ,COMPANYNAME,CUSTNAME,HIGHCUSTCODE, CUSTOWNER , "
                strSelFields = strSelFields & " USE_FLAG, "
                strSelFields = strSelFields & " CASE CUSTTYPE WHEN '2' THEN '계열' ELSE '비계열' END AS CUSTTYPE, "
                strSelFields = strSelFields & " BUSISTAT,BUSITYPE, "
                strSelFields = strSelFields & " ZIPCODE, ADDRESS1, ADDRESS2, "
                strSelFields = strSelFields & " TEL, FAX,"
                strSelFields = strSelFields & " MEMO, "
                strSelFields = strSelFields & " DBO.SC_EMPNAME_FUN(UUSER) UUSER "

                strFormet = "select {0} from SC_CUST_HDR where 1=1 AND MEDFLAG = 'K' {1} "

                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CREHDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function DeleteRtn(ByVal strInfoXML As String, ByVal strCUSTCODE As String) As Integer   '데이터 DELETE

        Dim intRtn_desc As Integer      'Return변수( 처리건수 또는 0 )
        Dim intRtn As Integer      'Return변수( 처리건수 또는 0 )

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjceSC_CUST_DTL = New ceSC_CUST_DTL(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjceSC_CUST_DTL.DeleteDo(strCUSTCODE)
                ' 트랜잭션 Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션 RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn")
            Finally
                'DB접속 종료
                .mobjSCGLSql.SQLDisconnect()
                '사용한 Entity(개체Dispose)
                mobjceSC_CUST_DTL.Dispose()
            End Try
        End With
    End Function

    Public Function DeleteRtn_EXE(ByVal strInfoXML As String, _
                                  ByVal strHIGHCUSTCODE As String) As Integer   '데이터 DELETE

        Dim intRtn_desc As Integer
        Dim intRtn As Integer
        Dim intRtn_DTL As Integer

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjceSC_CUST_HDR = New ceSC_CUST_HDR(mobjSCGLConfig)
                mobjceSC_CUST_DTL = New ceSC_CUST_DTL(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjceSC_CUST_HDR.DeleteDo(strHIGHCUSTCODE)
                intRtn_DTL = mobjceSC_CUST_DTL.DeleteDo(strHIGHCUSTCODE)
                ' 트랜잭션 Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션 RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn")
            Finally
                'DB접속 종료
                .mobjSCGLSql.SQLDisconnect()
                '사용한 Entity(개체Dispose)
                mobjceSC_CUST_HDR.Dispose()
                mobjceSC_CUST_DTL.Dispose()
            End Try
        End With
    End Function

    Public Function DeleteRtn_REAL(ByVal strInfoXML As String, _
                                   ByVal strHIGHCUSTCODE As String, _
                                   ByVal strMEDFLAG As String) As Integer   '데이터 DELETE

        Dim intRtn_desc As Integer
        Dim intRtn As Integer
        Dim intRtn_DTL As Integer

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjceSC_CUST_HDR = New ceSC_CUST_HDR(mobjSCGLConfig)
                mobjceSC_CUST_DTL = New ceSC_CUST_DTL(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                ' 엔티티 오브젝트의 Delete 메소드 호출
                If strMEDFLAG = "R" Then
                    intRtn = mobjceSC_CUST_HDR.DeleteDo(strHIGHCUSTCODE)
                ElseIf strMEDFLAG = "B" Then
                    intRtn_DTL = mobjceSC_CUST_DTL.DeleteDo(strHIGHCUSTCODE)
                End If


                ' 트랜잭션 Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션 RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn")
            Finally
                'DB접속 종료
                .mobjSCGLSql.SQLDisconnect()
                '사용한 Entity(개체Dispose)
                mobjceSC_CUST_HDR.Dispose()
                mobjceSC_CUST_DTL.Dispose()
            End Try
        End With
    End Function

    ' =============== ProcessRtn_CUSTHDR    거래처 해더 저장
    Public Function ProcessRtn_CUSTHDR(ByVal strInfoXML As String, _
                                       ByVal vntData As Object, _
                                       ByVal strMEDFLAG As String) As Object

        Dim intRtn As Integer
        Dim intRtn2 As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strHIGHCUSTCODE
        Dim strCUSTTYPE
        Dim strCUSTCODE

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_CUST_HDR = New ceSC_CUST_HDR(mobjSCGLConfig)
                    mobjceSC_CUST_DTL = New ceSC_CUST_DTL(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    For i = 1 To intRows
                        strHIGHCUSTCODE = ""

                        If GetElement(vntData, "CUSTTYPE", intColCnt, i, OPTIONAL_STR) = "계열" Then
                            strCUSTTYPE = "2"
                        ElseIf GetElement(vntData, "CUSTTYPE", intColCnt, i, OPTIONAL_STR) = "비계열" Then
                            strCUSTTYPE = "1"
                        End If

                        If GetElement(vntData, "HIGHCUSTCODE", intColCnt, i, OPTIONAL_STR) = "" Then
                            strHIGHCUSTCODE = SelectRtn_HIGHCUSTCODE(strMEDFLAG)
                            intRtn = InsertRtn_SC_CUST_HDR(vntData, intColCnt, i, strHIGHCUSTCODE, strMEDFLAG, strCUSTTYPE)
                            '신규 광고주 등록시 자동 팀 생성
                            strCUSTCODE = SelectRtn_CUSTCODE(strMEDFLAG)

                            If strMEDFLAG = "A" Then
                                intRtn2 = InsertRtn_SC_CUST_DTL_TIM(vntData, intColCnt, i, strHIGHCUSTCODE, strCUSTCODE, strMEDFLAG, "0")
                            Else
                                intRtn2 = InsertRtn_SC_CUST_DTL_TIM(vntData, intColCnt, i, strHIGHCUSTCODE, strCUSTCODE, strMEDFLAG, "")
                            End If
                        Else
                            strHIGHCUSTCODE = GetElement(vntData, "HIGHCUSTCODE", intColCnt, i, OPTIONAL_STR)
                            intRtn = UpdateRtn_SC_CUST_HDR(vntData, intColCnt, i, strHIGHCUSTCODE, strMEDFLAG, strCUSTTYPE)

                            If GetElement(vntData, "USE_FLAG", intColCnt, i, OPTIONAL_STR) = 0 Then
                                intRtn2 = mobjceSC_CUST_HDR.Update_USEFLAG_DTL(strHIGHCUSTCODE, strMEDFLAG)
                            End If
                        End If
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_CUSTHDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_CUST_HDR.Dispose()
                mobjceSC_CUST_DTL.Dispose()
            End Try
        End With
    End Function

    ' =============== ProcessRtn_CUSTDTL    거래처 디테일 저장
    Public Function ProcessRtn_CUSTDTL(ByVal strInfoXML As String, _
                                       ByVal vntData As Object, _
                                       ByVal strMEDFLAG As String) As Object

        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strCUSTCODE
        Dim strGBNFLAG

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_CUST_DTL = New ceSC_CUST_DTL(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    For i = 1 To intRows
                        strCUSTCODE = ""

                        strGBNFLAG = ""
                        If GetElement(vntData, "GBNFLAG", intColCnt, i, OPTIONAL_STR) = "팀" Then
                            strGBNFLAG = "0"
                        ElseIf GetElement(vntData, "GBNFLAG", intColCnt, i, OPTIONAL_STR) = "CIC/사업부" Then
                            strGBNFLAG = "1"
                        End If

                        If GetElement(vntData, "CUSTCODE", intColCnt, i, OPTIONAL_STR) = "" Then
                            strCUSTCODE = SelectRtn_CUSTCODE(strMEDFLAG)
                            intRtn = InsertRtn_SC_CUST_DTL(vntData, intColCnt, i, strCUSTCODE, strMEDFLAG, strGBNFLAG)
                        Else
                            strCUSTCODE = GetElement(vntData, "CUSTCODE", intColCnt, i, OPTIONAL_STR)
                            intRtn = UpdateRtn_SC_CUST_DTL(vntData, intColCnt, i, strCUSTCODE, strMEDFLAG, strGBNFLAG)
                        End If
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_CUSTDTL")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_CUST_DTL.Dispose()
            End Try
        End With
    End Function

    ' =============== ProcessRtn_MEDDTL    매체 디테일 저장
    Public Function ProcessRtn_MEDDTL(ByVal strInfoXML As String, _
                                      ByVal vntData As Object, _
                                      ByVal strMEDFLAG As String) As Object

        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strCUSTCODE

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_CUST_DTL = New ceSC_CUST_DTL(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    For i = 1 To intRows
                        strCUSTCODE = ""

                        If GetElement(vntData, "CUSTCODE", intColCnt, i, OPTIONAL_STR) = "" Then
                            strCUSTCODE = SelectRtn_CUSTCODE(strMEDFLAG)
                            intRtn = InsertRtnMED_SC_CUST_DTL(vntData, intColCnt, i, strCUSTCODE, strMEDFLAG, "")
                        Else
                            strCUSTCODE = GetElement(vntData, "CUSTCODE", intColCnt, i, OPTIONAL_STR)
                            intRtn = UpdateRtnMED_SC_CUST_DTL(vntData, intColCnt, i, strCUSTCODE, strMEDFLAG, "")
                        End If
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_MEDDTL")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_CUST_DTL.Dispose()
            End Try
        End With
    End Function

    ' =============== ProcessRtn_EXEHDR    대행사 해더/디테일 동시 저장
    Public Function ProcessRtn_EXEHDR(ByVal strInfoXML As String, _
                                      ByVal vntData As Object, _
                                      ByVal strMEDFLAG As String) As Object

        Dim intRtn As Integer
        Dim intRtn2 As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strHIGHCUSTCODE
        Dim strCUSTCODE
        Dim strCUSTTYPE

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_CUST_HDR = New ceSC_CUST_HDR(mobjSCGLConfig)
                    mobjceSC_CUST_DTL = New ceSC_CUST_DTL(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    For i = 1 To intRows
                        strHIGHCUSTCODE = ""

                        If GetElement(vntData, "CUSTTYPE", intColCnt, i, OPTIONAL_STR) = "계열" Then
                            strCUSTTYPE = "2"
                        ElseIf GetElement(vntData, "CUSTTYPE", intColCnt, i, OPTIONAL_STR) = "비계열" Then
                            strCUSTTYPE = "1"
                        End If

                        If GetElement(vntData, "HIGHCUSTCODE", intColCnt, i, OPTIONAL_STR) = "" Then
                            strHIGHCUSTCODE = SelectRtn_HIGHCUSTCODE(strMEDFLAG)
                            intRtn = InsertRtn_SC_EXCUST_HDR(vntData, intColCnt, i, strHIGHCUSTCODE, strMEDFLAG, strCUSTTYPE)

                            intRtn = InsertRtnEXE_SC_CUST_DTL(vntData, intColCnt, i, strHIGHCUSTCODE, strMEDFLAG, "")
                        Else
                            strHIGHCUSTCODE = GetElement(vntData, "HIGHCUSTCODE", intColCnt, i, OPTIONAL_STR)
                            intRtn = UpdateRtn_SC_CUST_HDR(vntData, intColCnt, i, strHIGHCUSTCODE, strMEDFLAG, strCUSTTYPE)

                            strCUSTCODE = GetElement(vntData, "HIGHCUSTCODE", intColCnt, i, OPTIONAL_STR)
                            intRtn = UpdateRtnEXE_SC_CUST_DTL(vntData, intColCnt, i, strCUSTCODE, strMEDFLAG, "")

                        End If

                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_EXEHDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_CUST_HDR.Dispose()
                mobjceSC_CUST_DTL.Dispose()
            End Try
        End With
    End Function

    ' =============== ProcessRtn_MPPHDR    MPP 해더/디테일 동시 저장
    Public Function ProcessRtn_MPPHDR(ByVal strInfoXML As String, _
                                      ByVal vntData As Object, _
                                      ByVal strMEDFLAG As String) As Object

        Dim intRtn As Integer
        Dim intRtn2 As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strHIGHCUSTCODE
        Dim strCUSTCODE
        Dim strCUSTTYPE

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_CUST_HDR = New ceSC_CUST_HDR(mobjSCGLConfig)
                    mobjceSC_CUST_DTL = New ceSC_CUST_DTL(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    For i = 1 To intRows
                        strHIGHCUSTCODE = ""

                        If GetElement(vntData, "CUSTTYPE", intColCnt, i, OPTIONAL_STR) = "계열" Then
                            strCUSTTYPE = "2"
                        ElseIf GetElement(vntData, "CUSTTYPE", intColCnt, i, OPTIONAL_STR) = "비계열" Then
                            strCUSTTYPE = "1"
                        End If

                        If GetElement(vntData, "HIGHCUSTCODE", intColCnt, i, OPTIONAL_STR) = "" Then
                            strHIGHCUSTCODE = SelectRtn_HIGHCUSTCODE(strMEDFLAG)
                            intRtn = InsertRtnMPP_SC_CUST_HDR(vntData, intColCnt, i, strHIGHCUSTCODE, strMEDFLAG, strCUSTTYPE)

                            intRtn = InsertRtnEXE_SC_CUST_DTL(vntData, intColCnt, i, strHIGHCUSTCODE, strMEDFLAG, "")
                        Else
                            strHIGHCUSTCODE = GetElement(vntData, "HIGHCUSTCODE", intColCnt, i, OPTIONAL_STR)
                            intRtn = UpdateRtnMPP_SC_CUST_HDR(vntData, intColCnt, i, strHIGHCUSTCODE, strMEDFLAG, strCUSTTYPE)

                            strCUSTCODE = GetElement(vntData, "HIGHCUSTCODE", intColCnt, i, OPTIONAL_STR)
                            intRtn = UpdateRtnEXE_SC_CUST_DTL(vntData, intColCnt, i, strCUSTCODE, strMEDFLAG, "")

                        End If

                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_EXEHDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_CUST_HDR.Dispose()
                mobjceSC_CUST_DTL.Dispose()
            End Try
        End With
    End Function

    '신규 CUSTCODE 생성
    Public Function SelectRtn_HIGHCUSTCODE(ByVal strMEDFLAG As String) As String

        Dim strSQL As String
        Dim strFormat As String
        Dim strRtn As String

        With mobjSCGLConfig '기본정보 Config 개체

            Try
                strSQL = String.Format("select '{0}' + dbo.lpad(isnull(Max(substring(highcustcode,2,6)),0)+1,5,0) From SC_CUST_HDR WHERE MEDFLAG =  '{1}'", strMEDFLAG, strMEDFLAG)
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_HIGHCUSTCODE")
            Finally
            End Try
        End With
    End Function

    '신규 CUSTCODE 생성
    Public Function SelectRtn_CUSTCODE(ByVal strMEDFLAG As String) As String

        Dim strSQL As String
        Dim strFormat As String
        Dim strRtn As String

        With mobjSCGLConfig '기본정보 Config 개체

            Try
                strSQL = String.Format("select '{0}' + dbo.lpad(isnull(Max(substring(custcode,2,6)),0)+1,5,0) From SC_CUST_DTL WHERE MEDFLAG =  '{1}'", strMEDFLAG, strMEDFLAG)
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CUSTCODE")
            Finally
            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : RFC 를 태우기 위한 사업자 정보 조회 및 이벤트"
    ' =============== 기본 사업자번호가져오기 
    Public Function SelectRtn_BUSINO(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByVal strBUSINO As String, _
                                     ByVal strMEDFLAG As String, _
                                     ByVal lngTO As String, _
                                     ByVal lngFROM As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = "" : Con2 = ""

                If strBUSINO <> "" Then Con1 = String.Format(" AND (replace(BUSINO,'-','') = '{0}')", strBUSINO)
                If strMEDFLAG <> "" Then Con2 = String.Format(" AND (MEDFLAG = '{0}')", strMEDFLAG)

                strWhere = BuildFields(" ", Con1, Con2)

                strFormet = " SELECT "
                strFormet = strFormet & " NO, BUSINO, CUSTNAME "
                strFormet = strFormet & " from ( "
                strFormet = strFormet & " 	SELECT "
                strFormet = strFormet & " 	ROW_NUMBER() OVER(ORDER BY BUSINO DESC) NO, "
                strFormet = strFormet & " 	BUSINO,CUSTNAME "
                strFormet = strFormet & " 	FROM SC_CUST_HDR "
                strFormet = strFormet & " 	WHERE 1=1 "
                strFormet = strFormet & " 	AND USE_FLAG = '1' "
                strFormet = strFormet & "   {0} "
                strFormet = strFormet & " 	GROUP BY BUSINO,CUSTNAME "
                strFormet = strFormet & " )A "
                strFormet = strFormet & " WHERE NO BETWEEN '" & lngTO & "' AND '" & lngFROM & "' "


                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_BUSINO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
    ' =============== 기본 사업자번호 정보 가져오기 
    Public Function SelectRtn_DTL(ByVal strInfoXML As String, _
                                  ByRef intRowCnt As Integer, _
                                  ByRef intColCnt As Integer, _
                                  ByVal strBUSINO As String, _
                                  ByVal strMEDFLAG As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정
                Con1 = ""

                If strBUSINO <> "" Then Con1 = String.Format(" AND (replace(BUSINO,'-','') = '{0}')", strBUSINO)
                If strMEDFLAG <> "" Then Con2 = String.Format(" AND (MEDFLAG = '{0}')", strMEDFLAG)

                strWhere = BuildFields(" ", Con1, Con2)

                strFormet = " SELECT  "
                strFormet = strFormet & " CUSTNAME,CUSTOWNER,BUSISTAT,BUSITYPE,ADDRESS1,ADDRESS2,TEL "
                strFormet = strFormet & " FROM SC_CUST_HDR"
                strFormet = strFormet & " WHERE 1=1 "
                strFormet = strFormet & " AND USE_FLAG = '1'"
                strFormet = strFormet & " {0}"
                strFormet = strFormet & " GROUP BY CUSTNAME,CUSTOWNER,BUSISTAT,BUSITYPE,ADDRESS1,ADDRESS2,TEL"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_DTL")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' =============== 기본 BANK_TYPE
    Public Function SelectRtn_BANK(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strBUSINO As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정
                Con1 = ""

                If strBUSINO <> "" Then Con1 = String.Format(" AND (BUSINO = '{0}')", strBUSINO)

                strWhere = BuildFields(" ", Con1)

                strFormet = " SELECT "
                strFormet = strFormet & " BANK_KEY,BANK_NUM,BANK_TYPE,BANK_USER"
                strFormet = strFormet & " FROM SC_BANKTYPE_MST"
                strFormet = strFormet & " WHERE 1=1 "
                strFormet = strFormet & " {0}"
                strFormet = strFormet & " AND USE_YN = 'Y'"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_BANK")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function ProcessRtnRFC(ByVal strInfoXML As String, _
                                  ByVal strBUSINO As String, _
                                  ByVal strBANKTYPE As String, _
                                  ByVal strMEDFLAG As String) As Integer


        Dim intRtn, intRtn2 As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strSQL

        '사업자 정보 관련 업데이트
        Dim firstArray_busino
        Dim secondArray_busino
        Dim strSAUPNO, strNAME1, strCNAME, strORT01, strSTRAS, strPSTLZ, strTELF1, strCEO, strJ_1KFTBUS
        Dim strJ_1KFTIND, strREGIO, strMCOD1, strCRTDAY, strCRTWHO, strNAME2

        'BANK_TYPE 관련 정보 업데이트
        Dim firstArray_bank
        Dim secondArray_bank
        Dim strSAUPNOBANK, strBVTYP, strBANKL, strBANKN, strKOINH

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                firstArray_busino = Split(strBUSINO, ":", -1, CompareMethod.Text)
                firstArray_bank = Split(strBANKTYPE, ":", -1, CompareMethod.Text)

                If strBUSINO <> "" Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_CUST_SAP = New ceSC_CUST_SAP(mobjSCGLConfig)
                    mobjceSC_CUST_SAPBANK = New ceSC_CUST_SAPBANK(mobjSCGLConfig)

                    '각각 temp 테이블 삭제 초기화
                    strSQL = " DELETE FROM SC_CUST_SAP "
                    strSQL = strSQL & " ;DELETE FROM SC_CUST_SAPBANK "

                    mobjSCGLConfig.mobjSCGLSql.SQLDo(strSQL)

                    '사업자 정보 업데이트 
                    For i = 0 To firstArray_busino.length - 1
                        strSAUPNO = "" : strNAME1 = "" : strCNAME = "" : strORT01 = "" : strSTRAS = "" : strPSTLZ = "" : strTELF1 = "" : strCEO = ""
                        strJ_1KFTBUS = "" : strJ_1KFTIND = "" : strREGIO = "" : strMCOD1 = "" : strCRTDAY = "" : strCRTWHO = "" : strNAME2 = ""

                        secondArray_busino = Split(firstArray_busino(i), "|", -1, CompareMethod.Text)

                        strSAUPNO = secondArray_busino(0)
                        strNAME1 = secondArray_busino(1)
                        strCNAME = secondArray_busino(2)
                        strORT01 = secondArray_busino(3)
                        strSTRAS = secondArray_busino(4)
                        strPSTLZ = secondArray_busino(5)
                        strTELF1 = secondArray_busino(6)
                        strCEO = secondArray_busino(7)
                        strJ_1KFTBUS = secondArray_busino(8)
                        strJ_1KFTIND = secondArray_busino(9)
                        strREGIO = secondArray_busino(10)
                        strMCOD1 = secondArray_busino(11)
                        strCRTDAY = secondArray_busino(12)
                        strCRTWHO = secondArray_busino(13)
                        strNAME2 = secondArray_busino(14)

                        intRtn = InsertRtnRFCBUSINO(strSAUPNO, strNAME1, strCNAME, strORT01, strSTRAS, strPSTLZ, strTELF1, strCEO, strJ_1KFTBUS, strJ_1KFTIND, strREGIO, strMCOD1, strCRTDAY, strCRTWHO, strNAME2)
                    Next

                    'BANK_TYPE 관련정보 업데이트
                    For i = 0 To firstArray_bank.length - 1

                        strSAUPNOBANK = "" : strBVTYP = "" : strBANKL = "" : strBANKN = "" : strKOINH = ""

                        secondArray_bank = Split(firstArray_bank(i), "|", -1, CompareMethod.Text)

                        strSAUPNOBANK = secondArray_bank(0)
                        strBVTYP = secondArray_bank(1)
                        strBANKL = secondArray_bank(2)
                        strBANKN = secondArray_bank(3)
                        strKOINH = secondArray_bank(4)

                        intRtn = InsertRtnRFCBANK(strSAUPNOBANK, strBVTYP, strBANKL, strBANKN, strKOINH)
                    Next

                    intRtn2 = UpdateRtn_busino(strInfoXML, strMEDFLAG)

                    '각각 temp 테이블 삭제 초기화
                    strSQL = " DELETE FROM SC_CUST_SAP"
                    strSQL = strSQL & " ;DELETE FROM SC_CUST_SAPBANK"

                    mobjSCGLConfig.mobjSCGLSql.SQLDo(strSQL)

                End If

                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtnRFC")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_CUST_SAPBANK.Dispose()
                mobjceSC_CUST_SAP.Dispose()
            End Try
        End With
    End Function

    Public Function UpdateRtn_busino(ByVal strInfoXML As String, _
                                     ByVal strMEDFLAG As String) As Integer


        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer

        Dim strSQL
        Dim strSQLBANK

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try

                '거래처 정보 업데이트 
                strSQL = " UPDATE A "
                strSQL = strSQL & " SET A.CUSTNAME = B.NAME1,"
                strSQL = strSQL & " A.COMPANYNAME = B.CNAME, "
                strSQL = strSQL & " A.ADDRESS1 = B.ORT01,"
                strSQL = strSQL & " A.ADDRESS2 = B.STRAS,"
                strSQL = strSQL & " A.ZIPCODE = B.PSTLZ,"
                strSQL = strSQL & " A.TEL = B.TELF1,"
                strSQL = strSQL & " A.CUSTOWNER = B.CEO,"
                strSQL = strSQL & " A.BUSISTAT = B.J_1KFTBUS,"
                strSQL = strSQL & " A.BUSITYPE = B.J_1KFTIND"
                strSQL = strSQL & " FROM SC_CUST_HDR A LEFT JOIN SC_CUST_SAP B"
                strSQL = strSQL & " ON REPLACE(A.BUSINO,'-','') = SAUPNO "
                strSQL = strSQL & " WHERE 1=1"
                strSQL = strSQL & " AND A.USE_FLAG = '1'"
                strSQL = strSQL & " AND A.MEDFLAG = '" & strMEDFLAG & "'"
                strSQL = strSQL & " AND ISNULL(B.SAUPNO,'') <> '' "

                '뱅크타입 DELETE & INSERT
                strSQL = strSQL & " ;DELETE FROM SC_BANKTYPE_MST"
                strSQL = strSQL & " WHERE BUSINO IN ("
                strSQL = strSQL & "	SELECT SAUPNO "
                strSQL = strSQL & "	FROM SC_CUST_SAPBANK"
                strSQL = strSQL & "	GROUP BY  SAUPNO"
                strSQL = strSQL & " )"

                mobjSCGLConfig.mobjSCGLSql.SQLDo(strSQL)

                strSQLBANK = " INSERT INTO SC_BANKTYPE_MST (BUSINO, BANK_KEY,BANK_NUM, BANK_TYPE,BANK_USER,USE_YN)"
                strSQLBANK = strSQLBANK & " SELECT SAUPNO, BANKL,BANKN, BVTYP,KOINH,'Y' USE_YN FROM SC_CUST_SAPBANK"
                strSQLBANK = strSQLBANK & " GROUP BY SAUPNO, BANKL,BANKN, BVTYP,KOINH "

                mobjSCGLConfig.mobjSCGLSql.SQLDo(strSQLBANK)

                Return intRtn
            Catch err As Exception
                ' .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_busino")
            Finally

            End Try
        End With
    End Function

#End Region

#Region "GROUP BLOCK : 외부에 비공개 Method"
    Private Function InsertRtn_SC_CUST_HDR(ByVal vntData As Object, _
                                           ByVal intColCnt As Integer, _
                                           ByVal intRow As Integer, _
                                           ByVal strHIGHCUSTCODE As String, _
                                           ByVal strMEDFLAG As String, _
                                           ByVal strCUSTTYPE As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjceSC_CUST_HDR.InsertCLIENT( _
                                       strHIGHCUSTCODE, _
                                       strHIGHCUSTCODE, _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       GetElement(vntData, "COMPANYNAME", intColCnt, intRow), _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       strCUSTTYPE, _
                                       strMEDFLAG, _
                                       GetElement(vntData, "CUSTOWNER", intColCnt, intRow), _
                                       GetElement(vntData, "BUSINO", intColCnt, intRow), _
                                       GetElement(vntData, "BUSISTAT", intColCnt, intRow), _
                                       GetElement(vntData, "BUSITYPE", intColCnt, intRow), _
                                       GetElement(vntData, "ZIPCODE", intColCnt, intRow), _
                                       GetElement(vntData, "ADDRESS1", intColCnt, intRow), _
                                       GetElement(vntData, "ADDRESS2", intColCnt, intRow), _
                                       GetElement(vntData, "TEL", intColCnt, intRow), _
                                       GetElement(vntData, "FAX", intColCnt, intRow), _
                                       "1", _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR06", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR07", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR08", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR09", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR10", intColCnt, intRow, NULL_NUM, True))
        Return intRtn
    End Function

    Private Function InsertRtn_SC_EXCUST_HDR(ByVal vntData As Object, _
                                             ByVal intColCnt As Integer, _
                                             ByVal intRow As Integer, _
                                             ByVal strHIGHCUSTCODE As String, _
                                             ByVal strMEDFLAG As String, _
                                             ByVal strCUSTTYPE As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjceSC_CUST_HDR.InsertEXCLIENT( _
                                       strHIGHCUSTCODE, _
                                       strHIGHCUSTCODE, _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       GetElement(vntData, "COMPANYNAME", intColCnt, intRow), _
                                       strCUSTTYPE, _
                                       strMEDFLAG, _
                                       GetElement(vntData, "CUSTOWNER", intColCnt, intRow), _
                                       GetElement(vntData, "BUSINO", intColCnt, intRow), _
                                       GetElement(vntData, "BUSISTAT", intColCnt, intRow), _
                                       GetElement(vntData, "BUSITYPE", intColCnt, intRow), _
                                       GetElement(vntData, "ZIPCODE", intColCnt, intRow), _
                                       GetElement(vntData, "ADDRESS1", intColCnt, intRow), _
                                       GetElement(vntData, "ADDRESS2", intColCnt, intRow), _
                                       GetElement(vntData, "TEL", intColCnt, intRow), _
                                       GetElement(vntData, "FAX", intColCnt, intRow), _
                                       "1", _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR06", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR07", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR08", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR09", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR10", intColCnt, intRow, NULL_NUM, True))
        Return intRtn
    End Function

    Private Function UpdateRtn_SC_CUST_HDR(ByVal vntData As Object, _
                                           ByVal intColCnt As Integer, _
                                           ByVal intRow As Integer, _
                                           ByVal strCUSTCODE As String, _
                                           ByVal strMEDFLAG As String, _
                                           ByVal strCUSTTYPE As String) As Integer
        Dim intRtn As Integer

        intRtn = mobjceSC_CUST_HDR.UpdateCLIENT( _
                                       GetElement(vntData, "HIGHCUSTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       GetElement(vntData, "COMPANYNAME", intColCnt, intRow), _
                                       strCUSTTYPE, _
                                       GetElement(vntData, "CUSTOWNER", intColCnt, intRow), _
                                       GetElement(vntData, "BUSISTAT", intColCnt, intRow), _
                                       GetElement(vntData, "BUSITYPE", intColCnt, intRow), _
                                       GetElement(vntData, "ZIPCODE", intColCnt, intRow), _
                                       GetElement(vntData, "ADDRESS1", intColCnt, intRow), _
                                       GetElement(vntData, "ADDRESS2", intColCnt, intRow), _
                                       GetElement(vntData, "TEL", intColCnt, intRow), _
                                       GetElement(vntData, "FAX", intColCnt, intRow), _
                                       GetElement(vntData, "USE_FLAG", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR08", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR10", intColCnt, intRow, NULL_NUM, True))

        Return intRtn
    End Function


    Private Function InsertRtn_SC_CUST_DTL(ByVal vntData As Object, _
                                           ByVal intColCnt As Integer, _
                                           ByVal intRow As Integer, _
                                           ByVal strCUSTCODE As String, _
                                           ByVal strMEDFLAG As String, _
                                           ByVal strGBNFLAG As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjceSC_CUST_DTL.InsertCLIENT( _
                                       strCUSTCODE, _
                                       GetElement(vntData, "HIGHCUSTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       strGBNFLAG, _
                                       GetElement(vntData, "CLIENTSUBCODE", intColCnt, intRow), _
                                       strMEDFLAG, _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       "1", _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR06", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR07", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR08", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR09", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR10", intColCnt, intRow, NULL_NUM, True))
        Return intRtn
    End Function

    Private Function InsertRtn_SC_CUST_DTL_TIM(ByVal vntData As Object, _
                                               ByVal intColCnt As Integer, _
                                               ByVal intRow As Integer, _
                                               ByVal strHIGHCUSTCODE As String, _
                                               ByVal strCUSTCODE As String, _
                                               ByVal strMEDFLAG As String, _
                                               ByVal strGBNFLAG As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjceSC_CUST_DTL.InsertCLIENT( _
                                       strCUSTCODE, _
                                       strHIGHCUSTCODE, _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       strGBNFLAG, _
                                       GetElement(vntData, "CLIENTSUBCODE", intColCnt, intRow), _
                                       strMEDFLAG, _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       "1", _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR06", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR07", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR08", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR09", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR10", intColCnt, intRow, NULL_NUM, True))
        Return intRtn
    End Function

    Private Function UpdateRtn_SC_CUST_DTL(ByVal vntData As Object, _
                                           ByVal intColCnt As Integer, _
                                           ByVal intRow As Integer, _
                                           ByVal strCUSTCODE As String, _
                                           ByVal strMEDFLAG As String, _
                                           ByVal strGBNFLAG As String) As Integer
        Dim intRtn As Integer

        intRtn = mobjceSC_CUST_DTL.UpdateCLIENT( _
                                       strCUSTCODE, _
                                       GetElement(vntData, "HIGHCUSTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       strGBNFLAG, _
                                       GetElement(vntData, "CLIENTSUBCODE", intColCnt, intRow), _
                                       strMEDFLAG, _
                                       GetElement(vntData, "USE_FLAG", intColCnt, intRow))

        Return intRtn
    End Function

    Private Function InsertRtnMED_SC_CUST_DTL(ByVal vntData As Object, _
                                              ByVal intColCnt As Integer, _
                                              ByVal intRow As Integer, _
                                              ByVal strCUSTCODE As String, _
                                              ByVal strMEDFLAG As String, _
                                              ByVal strGBNFLAG As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjceSC_CUST_DTL.InsertMED( _
                                       strCUSTCODE, _
                                       GetElement(vntData, "HIGHCUSTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       strMEDFLAG, _
                                       GetElement(vntData, "MED_TV", intColCnt, intRow), _
                                       GetElement(vntData, "MED_RD", intColCnt, intRow), _
                                       GetElement(vntData, "MED_DMB", intColCnt, intRow), _
                                       GetElement(vntData, "MED_CATV", intColCnt, intRow), _
                                       GetElement(vntData, "MED_GEN", intColCnt, intRow), _
                                       GetElement(vntData, "MED_PAP", intColCnt, intRow), _
                                       GetElement(vntData, "MED_MAG", intColCnt, intRow), _
                                       GetElement(vntData, "MED_MST", intColCnt, intRow), _
                                       GetElement(vntData, "MED_OUT", intColCnt, intRow), _
                                       GetElement(vntData, "MED_ETC", intColCnt, intRow), _
                                       GetElement(vntData, "MPP", intColCnt, intRow), _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       "1", _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR06", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR07", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR08", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR09", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR10", intColCnt, intRow, NULL_NUM, True))
        Return intRtn
    End Function

    Private Function UpdateRtnMED_SC_CUST_DTL(ByVal vntData As Object, _
                                              ByVal intColCnt As Integer, _
                                              ByVal intRow As Integer, _
                                              ByVal strCUSTCODE As String, _
                                              ByVal strMEDFLAG As String, _
                                              ByVal strGBNFLAG As String) As Integer
        Dim intRtn As Integer

        intRtn = mobjceSC_CUST_DTL.UpdateMED( _
                                       strCUSTCODE, _
                                       GetElement(vntData, "HIGHCUSTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       strMEDFLAG, _
                                       GetElement(vntData, "MED_TV", intColCnt, intRow), _
                                       GetElement(vntData, "MED_RD", intColCnt, intRow), _
                                       GetElement(vntData, "MED_DMB", intColCnt, intRow), _
                                       GetElement(vntData, "MED_CATV", intColCnt, intRow), _
                                       GetElement(vntData, "MED_GEN", intColCnt, intRow), _
                                       GetElement(vntData, "MED_PAP", intColCnt, intRow), _
                                       GetElement(vntData, "MED_MAG", intColCnt, intRow), _
                                       GetElement(vntData, "MED_NET", intColCnt, intRow), _
                                       GetElement(vntData, "MED_OUT", intColCnt, intRow), _
                                       GetElement(vntData, "MED_ETC", intColCnt, intRow), _
                                       GetElement(vntData, "MPP", intColCnt, intRow), _
                                       GetElement(vntData, "USE_FLAG", intColCnt, intRow))

        Return intRtn
    End Function

    Private Function InsertRtnEXE_SC_CUST_DTL(ByVal vntData As Object, _
                                              ByVal intColCnt As Integer, _
                                              ByVal intRow As Integer, _
                                              ByVal strCUSTCODE As String, _
                                              ByVal strMEDFLAG As String, _
                                              ByVal strGBNFLAG As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjceSC_CUST_DTL.InsertCLIENT( _
                                       strCUSTCODE, _
                                       strCUSTCODE, _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       strGBNFLAG, _
                                       GetElement(vntData, "CLIENTSUBCODE", intColCnt, intRow), _
                                       strMEDFLAG, _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       "1", _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR06", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR07", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR08", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR09", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR10", intColCnt, intRow, NULL_NUM, True))
        Return intRtn
    End Function

    Private Function UpdateRtnEXE_SC_CUST_DTL(ByVal vntData As Object, _
                                              ByVal intColCnt As Integer, _
                                              ByVal intRow As Integer, _
                                              ByVal strCUSTCODE As String, _
                                              ByVal strMEDFLAG As String, _
                                              ByVal strGBNFLAG As String) As Integer
        Dim intRtn As Integer

        intRtn = mobjceSC_CUST_DTL.UpdateEXE( _
                                       strCUSTCODE, _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       strGBNFLAG, _
                                       strMEDFLAG, _
                                       GetElement(vntData, "USE_FLAG", intColCnt, intRow))

        Return intRtn
    End Function


    Private Function InsertRtnMPP_SC_CUST_HDR(ByVal vntData As Object, _
                                              ByVal intColCnt As Integer, _
                                              ByVal intRow As Integer, _
                                              ByVal strHIGHCUSTCODE As String, _
                                              ByVal strMEDFLAG As String, _
                                              ByVal strCUSTTYPE As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjceSC_CUST_HDR.InsertMPPCLIENT( _
                                       strHIGHCUSTCODE, _
                                       strHIGHCUSTCODE, _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       strCUSTTYPE, _
                                       strMEDFLAG, _
                                       GetElement(vntData, "CUSTOWNER", intColCnt, intRow), _
                                       GetElement(vntData, "BUSINO", intColCnt, intRow), _
                                       GetElement(vntData, "BUSISTAT", intColCnt, intRow), _
                                       GetElement(vntData, "BUSITYPE", intColCnt, intRow), _
                                       GetElement(vntData, "ZIPCODE", intColCnt, intRow), _
                                       GetElement(vntData, "ADDRESS1", intColCnt, intRow), _
                                       GetElement(vntData, "ADDRESS2", intColCnt, intRow), _
                                       GetElement(vntData, "TEL", intColCnt, intRow), _
                                       GetElement(vntData, "FAX", intColCnt, intRow), _
                                       "1", _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR06", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR07", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR08", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR09", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR10", intColCnt, intRow, NULL_NUM, True))
        Return intRtn
    End Function

    Private Function UpdateRtnMPP_SC_CUST_HDR(ByVal vntData As Object, _
                                            ByVal intColCnt As Integer, _
                                            ByVal intRow As Integer, _
                                            ByVal strCUSTCODE As String, _
                                            ByVal strMEDFLAG As String, _
                                            ByVal strCUSTTYPE As String) As Integer
        Dim intRtn As Integer

        intRtn = mobjceSC_CUST_HDR.UpdateCLIENT( _
                                       GetElement(vntData, "HIGHCUSTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       GetElement(vntData, "CUSTNAME", intColCnt, intRow), _
                                       strCUSTTYPE, _
                                       GetElement(vntData, "CUSTOWNER", intColCnt, intRow), _
                                       GetElement(vntData, "BUSISTAT", intColCnt, intRow), _
                                       GetElement(vntData, "BUSITYPE", intColCnt, intRow), _
                                       GetElement(vntData, "ZIPCODE", intColCnt, intRow), _
                                       GetElement(vntData, "ADDRESS1", intColCnt, intRow), _
                                       GetElement(vntData, "ADDRESS2", intColCnt, intRow), _
                                       GetElement(vntData, "TEL", intColCnt, intRow), _
                                       GetElement(vntData, "FAX", intColCnt, intRow), _
                                       GetElement(vntData, "USE_FLAG", intColCnt, intRow))

        Return intRtn
    End Function

    Private Function InsertRtnRFCBUSINO(ByVal strSAUPNO As String, _
                                        ByVal strNAME1 As String, _
                                        ByVal strCNAME As String, _
                                        ByVal strORT01 As String, _
                                        ByVal strSTRAS As String, _
                                        ByVal strPSTLZ As String, _
                                        ByVal strTELF1 As String, _
                                        ByVal strCEO As String, _
                                        ByVal strJ_1KFTBUS As String, _
                                        ByVal strJ_1KFTIND As String, _
                                        ByVal strREGIO As String, _
                                        ByVal strMCOD1 As String, _
                                        ByVal strCRTDAY As String, _
                                        ByVal strCRTWHO As String, _
                                        ByVal strNAME2 As String) As Integer


        Dim intRtn As Integer
        intRtn = mobjceSC_CUST_SAP.InsertDo( _
                                       strSAUPNO, _
                                       strNAME1, _
                                       strCNAME, _
                                       strORT01, _
                                       strSTRAS, _
                                       strPSTLZ, _
                                       strTELF1, _
                                       strCEO, _
                                       strJ_1KFTBUS, _
                                       strJ_1KFTIND, _
                                       strREGIO, _
                                       strMCOD1, _
                                       strCRTDAY, _
                                       strCRTWHO, _
                                       strNAME2)

        Return intRtn
    End Function

    Private Function InsertRtnRFCBANK(ByVal strSAUPNO As String, _
                                      ByVal strBVTYP As String, _
                                      ByVal strBANKL As String, _
                                      ByVal strBANKN As String, _
                                      ByVal strKOINH As String) As Integer


        Dim intRtn As Integer
        intRtn = mobjceSC_CUST_SAPBANK.InsertDo( _
                                       strSAUPNO, _
                                       strBVTYP, _
                                       strBANKL, _
                                       strBANKN, _
                                       strKOINH)

        Return intRtn
    End Function
#End Region
End Class



