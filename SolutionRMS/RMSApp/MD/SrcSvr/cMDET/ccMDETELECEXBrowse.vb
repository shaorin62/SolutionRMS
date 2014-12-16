'****************************************************************************************
'Generated By :::BELLWORLD Making Entity Bean Ver1.0:::
'시스템구분 : SFAR/시스템명/Server단 Entity Base Component
'실행  환경 : COM+ Service Server Package
'프로그램명 : ccSCEXMain.vb (Table한글명 Entity 처리 Class)
'기      능 : Table한글명(영문명) Entity에 대해Insert/Update/Delete/Select를 처리
'특이  사항 : -특이사항에 대해 표현
'             -
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003-06-19 오후 2:35:59 By 작성자
'            2) 2003-06-19 오후 2:35:59 By 작성자 : 변경 내용 설명
'****************************************************************************************

Imports System.Xml                  'XML처리
Imports SCGLControl                 'ControlClass의 Base Class
Imports SCGLUtil.cbSCGLConfig       'ConfigurationClass
Imports SCGLUtil.cbSCGLErr          '오류처리클래스
Imports SCGLUtil.cbSCGLXml          'XML처리클래스
Imports SCGLUtil.cbSCGLUtil         '기타유틸리티 클래스
Imports eMDCO

Public Class ccMDETELECEXBrowse
    Inherits ccControl

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccMDETELECEXBrowse"         '자신의클래스명
    Dim mobjceSC_USER_COL As eMDCO.ceSC_USER_COL
    Dim mobjceSC_USER_TAB As eMDCO.ceSC_USER_TAB
    'Private Const .DBConnStr = "Provider=SQLOLEDB;Data Source=10.110.10.86;Initial Catalog=MCDEV;DSN=MCDEV;UID=devadmin;Pwd=password" '커넥션Setting

#End Region

#Region "GROUP BLOCk : Property 선언"
#End Region

#Region "GROUP BLOCk : Event 선언"
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
    Public Function SelectUserTab(ByVal strInfoXml As String, ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                  Optional ByVal strTabName As String = "", Optional ByVal dblTabID As Double = OPTIONAL_NUM) As Object

        SetConfig(strInfoXml)   '모든Public 함수의 처음오는 기본정보 설정

        Dim vntData As Object        '''Array Object
        Dim strFields As String      '''{0}번째 필드명
        Dim strKeys As String        '''{1}번째조건절
        Dim strFormat As String      '''Format을 맞추어 주는 String
        Dim strSQL As String         '''SQL 문을 만들어 주는 String

        ''''{0} 번째는 필드명으로한다.
        strFields = "A.TAB_ID, A.TAB_NAME, A.TAB_USER_NAME, A.TAB_TYPE, A.TAB_DESC "


        ''''{1} 번째는 WHERE 조건으로 한다.          
        strKeys = "(" & BuildFields("OR", _
                  GetFieldNameValue(" A.TAB_USER_NAME ", "%" & strTabName & "%", "like"), _
                  GetFieldNameValue(" A.TAB_NAME ", "%" & strTabName & "%", "like")) & ")"

        If dblTabID <> OPTIONAL_NUM Then strKeys &= " AND A.TAB_ID = " & dblTabID

        ''''{1} 이후 또는 전에 필수적으로 들어가야할 SQL문장을 추가주거나 ORDER BY 절을 넣어준다.

        strFormat = "SELECT {0} FROM SC_USER_TAB A " & _
                    "WHERE {1} ORDER BY A.TAB_ID "
        ''''strSQL 문을 만들어 준다.

        strSQL = String.Format(strFormat, strFields, strKeys)

        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectUserTab")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With

    End Function

#End Region

#Region "GROUP BLOCk : 외부에 비공개 Method"

#End Region
End Class
