'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 컨트롤 클래스 메이커
'시스템구분    : 솔루션명 /시스템명/Server Control Class
'실행   환경    : COM+ Service Server Package
'프로그램명    : ccMDCMDEPTMST.vb
'기         능    : - 기능을 명시 합니다.
'특이  사항     : - 특이사항에 대해 표현
'----------------------------------------------------------------------------------------
'HISTORY    :1) 
'            2) 
'****************************************************************************************

Imports System.Xml                  ' XML처리
Imports SCGLControl                 ' ControlClass의 Base Class
Imports SCGLUtil.cbSCGLConfig       ' ConfigurationClass
Imports SCGLUtil.cbSCGLErr          '오류처리 클래스
Imports SCGLUtil.cbSCGLXml          'XML처리 클래스
Imports SCGLUtil.cbSCGLUtil         '기타유틸리티 클래스
Imports eMDCO                       '엔터티 추가

' 엔티티 클래스 사용시 해당 엔티티 클래스의 프로젝트를 참조한 후 Imports 하십시요. 
' Imports 엔티티프로젝트

Public Class ccMDSCAORVOCH
    Inherits ccControl
#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccMDSCAORVOCH"                  '자신의 클래스명
    Private mobjcePD_MERGEVOCH_MST As eMDCO.cePD_MERGEVOCH_MST   '사용할 Entity 변수 선언
#End Region

#Region "GROUP BLOCK : 외부에비공개"
    Public Function Get_COMBO_VALUE(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, _
                                    ByRef intColCnt As Integer, _
                                    ByVal strCODE As String) As Object

        Dim strSQL, strFormat, strSelFields As String
        Dim vntData As Object

        SetConfig(strInfoXML)   '기본정보 설정					

        '조회 필드 설정					
        strSelFields = "CODE, CODE_NAME"

        'SQL문 생성
        strFormat = "SELECT {0} " & _
                    "FROM SC_CODE " & _
                    "WHERE CLASS_CODE = '" & strCODE & "' and use_yn = 'Y'" & _
                    "ORDER BY SORT_SEQ "

        With mobjSCGLConfig
            strSQL = String.Format(strFormat, strSelFields)

            ''데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_COMBO_VALUE")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function SelectRtn_SUSU(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strYEARMON As String, _
                                   ByVal strCLIENTCODE As String, _
                                   ByVal strCLIENTNAME As String, _
                                   ByVal strGBN As String) As Object

        Dim strSQL, strFormat, strSelFields, strKeys As String
        Dim strCondition As String
        Dim strCondition2 As String
        Dim strChkDate As String = ""
        Dim vntData As Object
        Dim Con1, Con2, Con3, Con4 As String
        Dim strWhere As String

        Con1 = "" : Con2 = "" : Con3 = "" : Con4 = ""

        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            If strYEARMON <> "" Then Con1 = String.Format(" AND (A.DEMANDDAY LIKE '%{0}%')", strYEARMON)
            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND A.REAL_MED_CODE = '{0}'", strCLIENTCODE)
            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND DBO.SC_GET_HIGHCUSTNAME_FUN(A.REAL_MED_CODE) LIKE '%{0}%'", strCLIENTNAME)

            If strGBN <> "" Then
                If strGBN = "rdT" Then
                    Con4 = " AND ISNULL(A.VOCHNO,'') <> '' "
                ElseIf strGBN = "rdF" Then
                    Con4 = " AND ISNULL(A.VOCHNO,'') = '' AND ISNULL(B.ERRMSG,'') = '' "
                ElseIf strGBN = "rdE" Then
                    Con4 = " AND ISNULL(B.ERRMSG,'') <> '' "
                End If
            End If

            strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

            strSelFields = strSelFields & " SELECT "
            strSelFields = strSelFields & " 'AOR' MEDFLAGNAME,"
            strSelFields = strSelFields & " CASE A.DEMANDDAY WHEN '' THEN B.POSTINGDATE ELSE A.DEMANDDAY END  POSTINGDATE,"
            strSelFields = strSelFields & " replace(dbo.SC_GET_BUSINO_FUN(REAL_MED_CODE),'-','') customercode,"
            strSelFields = strSelFields & " dbo.sc_get_highcompanyname_fun(REAL_MED_CODE) CUSTNAME,"
            strSelFields = strSelFields & " A.AMT,"
            strSelFields = strSelFields & " A.VAT,"
            strSelFields = strSelFields & " A.TAXYEARMON,"
            strSelFields = strSelFields & " A.TAXNO,"
            strSelFields = strSelFields & " A.VOCHNO,"
            strSelFields = strSelFields & " B.RMSNO,"
            strSelFields = strSelFields & " B.ERRCODE,"
            strSelFields = strSelFields & " B.ERRMSG "
            strSelFields = strSelFields & " FROM MD_COMMITAX_HDR  A "
            strSelFields = strSelFields & " LEFT JOIN PD_MERGEVOCH_MST B "
            strSelFields = strSelFields & " ON A.TAXYEARMON = B.MTAXYEARMON AND A.TAXNO = B.MTAXNO AND ISNULL(B.MTAXNOSEQ, 0) = '0' and B.GBN = 'M' "
            strSelFields = strSelFields & " WHERE isnull(A.attr10,0) <> 999999"
            strSelFields = strSelFields & " AND ISNULL(A.ATTR03,'') = 'AOR'"
            strSelFields = strSelFields & " {0} "
            strSelFields = strSelFields & " ORDER BY  A.TAXYEARMON, A.TAXNO"

            strSQL = String.Format(strSelFields, strWhere)
            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_HDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function SelectRtn_SUSUDTL(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, _
                                      ByRef intColCnt As Integer, _
                                      ByVal strTAXYEARMON As String, _
                                      ByVal strTAXNO As String) As Object

        Dim strSQL, strFormat, strSelFields, strKeys As String
        Dim strCondition As String
        Dim strCondition2 As String
        Dim strChkDate As String = ""
        Dim vntData As Object
        Dim Con1, Con2, Con3 As String
        Dim strWhere As String

        Con1 = "" : Con2 = ""

        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            If strTAXYEARMON <> "" Then Con1 = String.Format(" AND (A.TAXYEARMON =  '{0}')", strTAXYEARMON)
            If strTAXNO <> "" Then Con2 = String.Format(" AND (A.TAXNO = '{0}')", strTAXNO)

            strWhere = BuildFields(" ", Con1, Con2)

            strFormat = strFormat & " SELECT "
            strFormat = strFormat & " '1' GFLAG, "
            strFormat = strFormat & " '1' GUBUN, "
            strFormat = strFormat & " '통합' MEDFLAGNAME,"
            strFormat = strFormat & " CASE A.DEMANDDAY WHEN '' THEN B.POSTINGDATE ELSE A.DEMANDDAY END  POSTINGDATE,"
            strFormat = strFormat & " REPLACE(DBO.SC_GET_BUSINO_FUN(REAL_MED_CODE),'-','') CUSTOMERCODE,"
            strFormat = strFormat & " DBO.SC_GET_HIGHCOMPANYNAME_FUN(REAL_MED_CODE) CUSTNAME,"
            strFormat = strFormat & " CASE ISNULL(B.SUMM,'') WHEN '' "
            strFormat = strFormat & " THEN  CONVERT(CHAR(11),RTRIM(LTRIM(DBO.SC_GET_HIGHCOMPANYNAME_FUN(REAL_MED_CODE))))+'AOR 대행매출' "
            strFormat = strFormat & " ELSE B.SUMM END AS SUMM,"
            strFormat = strFormat & " CASE ISNULL(B.BA,'') WHEN '' THEN DBO.SC_GET_BA(A.DEPT_CD) ELSE B.BA END BA,"
            strFormat = strFormat & " CASE ISNULL(B.COSTCENTER,'') WHEN '' THEN DBO.SC_GET_CCTR(A.DEPT_CD) ELSE B.COSTCENTER END AS COSTCENTER,"
            strFormat = strFormat & " CASE ISNULL(B.COSTCENTER,'') WHEN '' THEN DBO.SC_DEPT_NAME_FUN(A.DEPT_CD) ELSE DBO.SC_DEPT_NAME_FUN(DBO.SC_GET_CCTR_DEPT_CD(B.COSTCENTER)) END AS DEPT_NAME,"
            strFormat = strFormat & " ISNULL(A.AMT,0) AMT,"
            strFormat = strFormat & " ISNULL(A.VAT,0) VAT ,"
            strFormat = strFormat & " 'B0' SEMU,"
            strFormat = strFormat & " '7040' BP,"
            strFormat = strFormat & " CASE ISNULL(B.DEMANDDAY,'') WHEN '' THEN CONVERT(CHAR(8),DATEADD(DD,90,A.DEMANDDAY ),112) ELSE B.DEMANDDAY END AS DEMANDDAY, "
            strFormat = strFormat & " CASE ISNULL(B.DUEDATE,'') WHEN '' THEN CONVERT(CHAR(8),DATEADD(DD,60,A.DEMANDDAY),112) ELSE B.DUEDATE END AS DUEDATE, "
            strFormat = strFormat & " REPLACE(DBO.SC_GET_BUSINO_FUN(REAL_MED_CODE),'-','') VENDOR,"
            strFormat = strFormat & " 'M' GBN,"
            strFormat = strFormat & " '200400' ACCOUNT,"
            strFormat = strFormat & " '' DEBTOR,"
            strFormat = strFormat & " CASE A.DEMANDDAY WHEN '' THEN B.DOCUMENTDATE ELSE A.DEMANDDAY END DOCUMENTDATE,"
            strFormat = strFormat & " '' PREPAYMENT,"
            strFormat = strFormat & " '' FROMDATE,"
            strFormat = strFormat & " '' TODATE,"
            strFormat = strFormat & " CASE ISNULL(B.SUMMTEXT,'') WHEN '' "
            strFormat = strFormat & " THEN  CONVERT(CHAR(11),RTRIM(LTRIM(DBO.SC_GET_HIGHCOMPANYNAME_FUN(REAL_MED_CODE))))+'AOR 대행매출' "
            strFormat = strFormat & " ELSE B.SUMM END AS SUMMTEXT,"
            strFormat = strFormat & " CASE WHEN ISNULL(A.AMT,0) < 0 THEN '-' ELSE '+' END AMTGBN,"
            strFormat = strFormat & " A.TAXYEARMON,"
            strFormat = strFormat & " A.TAXNO,"
            strFormat = strFormat & " 0 TAXNOSEQ,"
            strFormat = strFormat & " A.VOCHNO,"
            strFormat = strFormat & " 'AOR' RMSNO,"
            strFormat = strFormat & " B.ERRCODE,"
            strFormat = strFormat & " B.ERRMSG,"
            strFormat = strFormat & " 'AOR' MEDFLAG"
            strFormat = strFormat & " FROM MD_COMMITAX_HDR  A "
            strFormat = strFormat & " LEFT JOIN PD_MERGEVOCH_MST B "
            strFormat = strFormat & " ON A.TAXYEARMON = B.MTAXYEARMON AND A.TAXNO = B.MTAXNO AND B.GBN = 'M' AND ISNULL(B.MTAXNOSEQ, 0) = '0'"
            strFormat = strFormat & " WHERE(ISNULL(A.ATTR10, 0) <> 999999)"
            strFormat = strFormat & " AND ISNULL(A.ATTR03,'') = 'AOR'"
            strFormat = strFormat & " {0}"
            strFormat = strFormat & " UNION ALL"
            strFormat = strFormat & " SELECT "
            strFormat = strFormat & " '1' GFLAG, "
            strFormat = strFormat & " '2' GUBUN, "
            strFormat = strFormat & " CASE A.MEDFLAG WHEN 'A' THEN '공중파'"
            strFormat = strFormat & " WHEN 'A2' THEN '케이블'"
            strFormat = strFormat & " WHEN 'T' THEN '종편'"
            strFormat = strFormat & " WHEN 'B' THEN '신문'"
            strFormat = strFormat & " WHEN 'C' THEN '잡지'"
            strFormat = strFormat & " WHEN 'O' THEN '인터넷'"
            strFormat = strFormat & " WHEN 'D' THEN '옥외' END MED_FLAGNAME,"
            strFormat = strFormat & " CASE DBO.MD_COMMITAXNO_DEMANDDAY_FUN(A.TAXYEARMON,A.TAXNO) WHEN '' THEN B.POSTINGDATE ELSE DBO.MD_COMMITAXNO_DEMANDDAY_FUN(A.TAXYEARMON,A.TAXNO) END  POSTINGDATE,"
            strFormat = strFormat & " REPLACE(DBO.SC_GET_BUSINO_FUN(REAL_MED_CODE),'-','') CUSTOMERCODE,"
            strFormat = strFormat & " DBO.SC_GET_HIGHCOMPANYNAME_FUN(REAL_MED_CODE) CUSTNAME,"
            strFormat = strFormat & " CASE ISNULL(B.SUMM,'') WHEN '' "
            strFormat = strFormat & " THEN  CONVERT(CHAR(11),RTRIM(LTRIM(DBO.SC_GET_HIGHCOMPANYNAME_FUN(REAL_MED_CODE))))+'AOR 대행매출' "
            strFormat = strFormat & " ELSE B.SUMM END AS SUMM,"
            strFormat = strFormat & " CASE ISNULL(B.BA,'') WHEN '' THEN DBO.SC_GET_BA(A.DEPT_CD) ELSE B.BA END BA,"
            strFormat = strFormat & " CASE ISNULL(B.COSTCENTER,'') WHEN '' THEN DBO.SC_GET_CCTR(A.DEPT_CD) ELSE B.COSTCENTER END AS COSTCENTER,"
            strFormat = strFormat & " CASE ISNULL(B.COSTCENTER,'') WHEN '' THEN DBO.SC_DEPT_NAME_FUN(A.DEPT_CD) "
            strFormat = strFormat & " ELSE DBO.SC_DEPT_NAME_FUN(DBO.SC_GET_CCTR_DEPT_CD(B.COSTCENTER)) END AS DEPT_NAME,"
            strFormat = strFormat & " ISNULL(A.AMT,0) AMT,"
            strFormat = strFormat & " 0 VAT,"
            strFormat = strFormat & " 'B0' SEMU,"
            strFormat = strFormat & " '7040' BP,"
            strFormat = strFormat & " CASE ISNULL(B.DEMANDDAY,'') WHEN '' THEN CONVERT(CHAR(8),DATEADD(DD,90,DBO.MD_COMMITAXNO_DEMANDDAY_FUN(A.TAXYEARMON,A.TAXNO)),112) ELSE B.DEMANDDAY END AS DEMANDDAY, "
            strFormat = strFormat & " CASE ISNULL(B.DUEDATE,'') WHEN '' THEN CONVERT(CHAR(8),DATEADD(DD,60,DBO.MD_COMMITAXNO_DEMANDDAY_FUN(A.TAXYEARMON,A.TAXNO)),112) ELSE B.DUEDATE END AS DUEDATE, "
            strFormat = strFormat & " REPLACE(DBO.SC_GET_BUSINO_FUN(REAL_MED_CODE),'-','') VENDOR,"
            strFormat = strFormat & " 'M' GBN,"
            strFormat = strFormat & " '' ACCOUNT,"
            strFormat = strFormat & " CASE ISNULL(B.DEBTOR,'') WHEN '' THEN '506940' ELSE B.DEBTOR END AS  DEBTOR,"
            strFormat = strFormat & " CASE DBO.MD_COMMITAXNO_DEMANDDAY_FUN(A.TAXYEARMON,A.TAXNO) WHEN '' THEN B.DOCUMENTDATE ELSE DBO.MD_COMMITAXNO_DEMANDDAY_FUN(A.TAXYEARMON,A.TAXNO) END DOCUMENTDATE,"
            strFormat = strFormat & " CASE ISNULL(B.PREPAYMENT,'') WHEN '' THEN '' ELSE B.PREPAYMENT END AS  PREPAYMENT,"
            strFormat = strFormat & " CASE ISNULL(B.FROMDATE,'') WHEN '' THEN '' ELSE B.FROMDATE END AS  FROMDATE,"
            strFormat = strFormat & " CASE ISNULL(B.TODATE,'') WHEN '' THEN '' ELSE B.TODATE END AS  TODATE,"
            strFormat = strFormat & " CASE ISNULL(B.SUMMTEXT,'') WHEN '' "
            strFormat = strFormat & " THEN  CONVERT(CHAR(11),RTRIM(LTRIM(DBO.SC_GET_HIGHCOMPANYNAME_FUN(REAL_MED_CODE))))+'합산청구' "
            strFormat = strFormat & " ELSE B.SUMM END AS SUMMTEXT,"
            strFormat = strFormat & " CASE WHEN ISNULL(A.AMT,0) < 0 THEN '-' ELSE '+' END AMTGBN,"
            strFormat = strFormat & " A.TAXYEARMON,"
            strFormat = strFormat & " A.TAXNO,"
            strFormat = strFormat & " A.TAXNOSEQ,"
            strFormat = strFormat & " B.VOCHNO,"
            strFormat = strFormat & " 'AOR' RMSNO,"
            strFormat = strFormat & " B.ERRCODE,"
            strFormat = strFormat & " B.ERRMSG,"
            strFormat = strFormat & " A.MEDFLAG"
            strFormat = strFormat & " FROM MD_COMMITAX_DTL  A "
            strFormat = strFormat & " LEFT JOIN PD_MERGEVOCH_MST B "
            strFormat = strFormat & " ON A.TAXYEARMON = B.MTAXYEARMON AND A.TAXNO = B.MTAXNO AND A.TAXNOSEQ = B.MTAXNOSEQ AND  B.GBN = 'M' AND B.RMSNO = 'AOR' "
            strFormat = strFormat & " WHERE ISNULL(A.ATTR10,0) <> 999999"
            strFormat = strFormat & " AND ISNULL(A.ATTR03,'') = 'AOR'"
            strFormat = strFormat & " {0}"
            strFormat = strFormat & " ORDER BY A.TAXYEARMON, A.TAXNO ,GUBUN, MEDFLAGNAME, DEBTOR DESC"

            strSQL = String.Format(strFormat, strWhere)
            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_SUSUDTL")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object, _
                               ByVal strRETURNLIST As String, _
                               ByVal strGBN As String) As Integer

        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strPOSTINGDATE As String
        Dim strDEMANDDAY As String
        Dim strDOCUMENTDATE As String
        Dim strDUEDATE As String
        Dim strFROMDATE As String
        Dim strTODATE As String

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjcePD_MERGEVOCH_MST = New cePD_MERGEVOCH_MST(mobjSCGLConfig)

                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    If strGBN = "M" Then
                        For i = 1 To intRows
                            If GetElement(vntData, "POSTINGDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strPOSTINGDATE = Replace(GetElement(vntData, "POSTINGDATE", intColCnt, i, OPTIONAL_STR), "-", "")
                            If GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strDEMANDDAY = Replace(GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR), "-", "")
                            If GetElement(vntData, "DOCUMENTDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strDOCUMENTDATE = Replace(GetElement(vntData, "DOCUMENTDATE", intColCnt, i, OPTIONAL_STR), "-", "")
                            If GetElement(vntData, "FROMDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strFROMDATE = Replace(GetElement(vntData, "FROMDATE", intColCnt, i, OPTIONAL_STR), "-", "")
                            If GetElement(vntData, "TODATE", intColCnt, i, OPTIONAL_STR) <> "" Then strTODATE = Replace(GetElement(vntData, "TODATE", intColCnt, i, OPTIONAL_STR), "-", "")
                            If GetElement(vntData, "DUEDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strDUEDATE = Replace(GetElement(vntData, "DUEDATE", intColCnt, i, OPTIONAL_STR), "-", "")

                            intRtn = InsertRtn(vntData, intColCnt, i, strPOSTINGDATE, strDEMANDDAY, strFROMDATE, strTODATE, strDOCUMENTDATE, strDUEDATE)
                        Next
                        '추후에 매입이 발생할경우 사용. 현재 사용안함.
                    ElseIf strGBN = "B" Then
                        For i = 1 To intRows
                            If GetElement(vntData, "CHK", intColCnt, i) = "1" Then
                                If GetElement(vntData, "POSTINGDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strPOSTINGDATE = Replace(GetElement(vntData, "POSTINGDATE", intColCnt, i, OPTIONAL_STR), "-", "")
                                If GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strDEMANDDAY = Replace(GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR), "-", "")
                                If GetElement(vntData, "DOCUMENTDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strDOCUMENTDATE = Replace(GetElement(vntData, "DOCUMENTDATE", intColCnt, i, OPTIONAL_STR), "-", "")
                                If GetElement(vntData, "FROMDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strFROMDATE = Replace(GetElement(vntData, "FROMDATE", intColCnt, i, OPTIONAL_STR), "-", "")
                                If GetElement(vntData, "TODATE", intColCnt, i, OPTIONAL_STR) <> "" Then strTODATE = Replace(GetElement(vntData, "TODATE", intColCnt, i, OPTIONAL_STR), "-", "")
                                If GetElement(vntData, "DUEDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strDUEDATE = Replace(GetElement(vntData, "DUEDATE", intColCnt, i, OPTIONAL_STR), "-", "")

                                intRtn = InsertRtn(vntData, intColCnt, i, strPOSTINGDATE, strDEMANDDAY, strFROMDATE, strTODATE, strDOCUMENTDATE, strDUEDATE)
                            End If
                        Next
                    End If
                    intRtn = UpdateRtn(strInfoXML, strRETURNLIST, strGBN)
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjcePD_MERGEVOCH_MST.Dispose()
            End Try
        End With
    End Function

    Public Function UpdateRtn(ByVal strInfoXML As String, _
                              ByVal strRETURNLIST As String, _
                              ByVal strGBN As String) As Integer

        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim firstArray
        Dim secondArray

        Dim strTAXYEARMON As String
        Dim intTAXNO As Integer
        Dim strDOC_STATUS As String
        Dim strDOC_MESSAGE As String
        Dim strVOCHNO As String
        Dim strIF_KEY As String

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                '.mobjSCGLSql.SQLConnect(.DBConnStr)
                '.mobjSCGLSql.SQLBeginTrans()

                firstArray = Split(strRETURNLIST, ":", -1, CompareMethod.Text)

                If strRETURNLIST <> "" Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjcePD_MERGEVOCH_MST = New cePD_MERGEVOCH_MST(mobjSCGLConfig)

                    For i = 0 To firstArray.length - 1
                        secondArray = Split(firstArray(i), "|", -1, CompareMethod.Text)

                        strTAXYEARMON = secondArray(0)
                        intTAXNO = CInt(secondArray(1))
                        strDOC_STATUS = secondArray(2)
                        strDOC_MESSAGE = Replace(secondArray(3), "'", "''")
                        strVOCHNO = secondArray(4)
                        strIF_KEY = secondArray(5)

                        If strDOC_STATUS = "S" Then
                            strDOC_STATUS = ""
                            strDOC_MESSAGE = ""
                        Else
                            strDOC_STATUS = "1"
                        End If

                        intRtn = UpdateRtn_VOCHNO(strGBN, strTAXYEARMON, intTAXNO, strDOC_STATUS, strDOC_MESSAGE, strVOCHNO, strIF_KEY)
                    Next

                End If
                '.mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                ' .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn")
            Finally
                '.mobjSCGLSql.SQLDisconnect()
                mobjcePD_MERGEVOCH_MST.Dispose()
            End Try
        End With
    End Function

    Public Function VOCHDELL(ByVal strInfoXML As String, _
                             ByVal strRETURNLIST As String, _
                             ByVal strGBN As String) As Integer
        Dim intRtn As Integer

        Dim i, intColCnt, intRows As Integer
        Dim firstArray
        Dim secondArray
        Dim strTAXYEARMON As String
        Dim intTAXNO As Integer
        Dim strDOC_STATUS As String
        Dim strDOC_MESSAGE As String
        Dim strVOCHNO As String

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                firstArray = Split(strRETURNLIST, ":", -1, CompareMethod.Text)
                If strRETURNLIST <> "" Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjcePD_MERGEVOCH_MST = New cePD_MERGEVOCH_MST(mobjSCGLConfig)

                    For i = 0 To firstArray.length - 1
                        secondArray = Split(firstArray(i), "|", -1, CompareMethod.Text)

                        strTAXYEARMON = secondArray(0)
                        intTAXNO = CInt(secondArray(1))
                        strDOC_STATUS = secondArray(2)
                        strDOC_MESSAGE = Replace(secondArray(3), "'", "''")
                        strVOCHNO = secondArray(4)

                        If strDOC_STATUS = "S" Then
                            '전표내역 삭제
                            intRtn = mobjcePD_MERGEVOCH_MST.Delete_vochno(strTAXYEARMON, strVOCHNO)
                            intRtn = mobjcePD_MERGEVOCH_MST.Update_AORVOCHNO(strTAXYEARMON, intTAXNO)
                            intRtn = mobjcePD_MERGEVOCH_MST.UpdateDelete(strTAXYEARMON, intTAXNO)

                        ElseIf strDOC_STATUS = "E" Then
                            '전표삭제시 SAP오류는 오류전표로 보내기위함.
                            strDOC_STATUS = "1"
                            intRtn = mobjcePD_MERGEVOCH_MST.UpdateRtn_DELETEERR(strGBN, strTAXYEARMON, intTAXNO, strDOC_STATUS, strDOC_MESSAGE, strVOCHNO)
                        End If
                    Next
                End If

                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".VOCHDELL")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjcePD_MERGEVOCH_MST.Dispose()
            End Try
        End With
    End Function

    Public Function DeleteRtn(ByVal strInfoXML As String, _
                              ByVal vntData As Object) As Integer
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim dblID As Double '자동 ID 를사용할 때만 사용
        Dim strSC_EMP_STATUS As String
        Dim vntData2 As Object
        Dim intSEQ As Double

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjcePD_MERGEVOCH_MST = New cePD_MERGEVOCH_MST(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력

                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    '''해당하는Row 만큼 Loop
                    For i = 1 To intRows
                        If GetElement(vntData, "CHK", intColCnt, i) = 1 And GetElement(vntData, "ERRCODE", intColCnt, i) = 1 Then
                            intRtn = DeleteRtn_ERR(vntData, intColCnt, i)
                        End If
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".DeleteRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjcePD_MERGEVOCH_MST.Dispose()
            End Try
        End With
    End Function

    Public Function DeleteRtn_GANG(ByVal strInfoXML As String, _
                                   ByVal strTAXYEARMON As String, _
                                   ByVal strTAXNO As String, _
                                   ByVal strVOCHNO As String, _
                                   ByVal strGBN As String) As Integer

        ' strGBN 변수는 추후에 통합전표에서 매입 전표를 해달라고 했을때를 위해 매출과 매입을 구분하는 용도로 사용할것 20110429 KTY
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim dblID As Double '자동 ID 를사용할 때만 사용
        Dim strSC_EMP_STATUS As String
        Dim vntData2 As Object
        Dim intSEQ As Double

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                mobjcePD_MERGEVOCH_MST = New cePD_MERGEVOCH_MST(mobjSCGLConfig)

                intRtn = mobjcePD_MERGEVOCH_MST.Delete_vochno(strTAXYEARMON, strVOCHNO)

                intRtn = mobjcePD_MERGEVOCH_MST.Update_AORVOCHNO(strTAXYEARMON, strTAXNO)

                intRtn = mobjcePD_MERGEVOCH_MST.UpdateDelete(strTAXYEARMON, strTAXNO)

                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".DeleteRtn_GANG")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjcePD_MERGEVOCH_MST.Dispose()
            End Try
        End With
    End Function

#End Region

#Region "GROUP BLOCK : 외부에 비공개 Method"
    Private Function InsertRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer, _
                               ByVal strPOSTINGDATE As String, _
                               ByVal strDEMANDDAY As String, _
                               ByVal strFROMDATE As String, _
                               ByVal strTODATE As String, _
                               ByVal strDOCUMENTDATE As String, _
                               ByVal strDUEDATE As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjcePD_MERGEVOCH_MST.InsertDo( _
                                       strPOSTINGDATE, _
                                       GetElement(vntData, "CUSTOMERCODE", intColCnt, intRow), _
                                       GetElement(vntData, "SUMM", intColCnt, intRow), _
                                       GetElement(vntData, "BA", intColCnt, intRow), _
                                       GetElement(vntData, "COSTCENTER", intColCnt, intRow), _
                                       GetElement(vntData, "AMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "VAT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "SEMU", intColCnt, intRow), _
                                       GetElement(vntData, "BP", intColCnt, intRow), _
                                       strDEMANDDAY, _
                                       strDUEDATE, _
                                       GetElement(vntData, "TAXYEARMON", intColCnt, intRow), _
                                       GetElement(vntData, "TAXNO", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "TAXNOSEQ", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "GBN", intColCnt, intRow), _
                                       GetElement(vntData, "VOCHNO", intColCnt, intRow), _
                                       GetElement(vntData, "RMSNO", intColCnt, intRow), _
                                       strDOCUMENTDATE, _
                                       GetElement(vntData, "PAYCODE", intColCnt, intRow), _
                                       GetElement(vntData, "PREPAYMENT", intColCnt, intRow), _
                                       strFROMDATE, _
                                       strTODATE, _
                                       GetElement(vntData, "SUMMTEXT", intColCnt, intRow), _
                                       GetElement(vntData, "ACCOUNT", intColCnt, intRow), _
                                       GetElement(vntData, "DEBTOR", intColCnt, intRow), _
                                       GetElement(vntData, "MEDFLAG", intColCnt, intRow))
        Return intRtn
    End Function

    Private Function DeleteRtn_ERR(ByVal vntData As Object, _
                                   ByVal intColCnt As Integer, _
                                   ByVal intRow As Integer) As Integer

        Dim intRtn As Integer

        intRtn = mobjcePD_MERGEVOCH_MST.Delete_ERR( _
                                       GetElement(vntData, "TAXYEARMON", intColCnt, intRow), _
                                       GetElement(vntData, "TAXNO", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "RMSNO", intColCnt, intRow))
        Return intRtn
    End Function

    Private Function UpdateRtn_VOCHNO(ByVal strGBN As String, _
                                      ByVal strTAXYEARMON As String, _
                                      ByVal intTAXNO As Integer, _
                                      ByVal strDOC_STATUS As String, _
                                      ByVal strDOC_MESSAGE As String, _
                                      ByVal strVOCHNO As String, _
                                      ByVal strIF_KEY As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjcePD_MERGEVOCH_MST.UpdateRtn_AORVOCHNO( _
                                       strGBN, _
                                       strTAXYEARMON, _
                                       intTAXNO, _
                                       strDOC_STATUS, _
                                       strDOC_MESSAGE, _
                                       strVOCHNO, _
                                       strIF_KEY)

        Return intRtn
    End Function
#End Region
End Class
