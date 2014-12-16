'****************************************************************************************
'Generated By: JNF BY Solution
'�ý��۱��� : SolutionOLD/MD/ceMD_BOOKING_MEDIUM Class
'����  ȯ�� : GAC(Global Assembly Cache)
'���α׷��� : ceMD_BOOKING_MEDIUM.vb ( MD_BOOKING_MEDIUM Entity ó�� Class)
'��      �� : MD_BOOKING_MEDIUM Entity�� ����Insert/Update/Delete/Select�� ó��
'             - �θ�ƼƼ ��ü�� SCGLUtil.ceEntity�� ���
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-08-18 ���� 2:40:17 By Kim Tae Ho
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '���� ��ƿ��Ƽ ��ü
Imports SCGLUtil.cbSCGLErr      '���� ����ó�� ��ü
Imports SCGLEntity              '��ƼƼ ��ü�� �θ� ��ü

Public Class ceMD_REPORT_MSTTOTAL
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "ceMD_REPORT_MSTTOTAL"    '�ڽ��� Ŭ������
#End Region

#Region "SQL Insert/Update/Delete/Select"
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************

    'YEARMON | GBN | AMTGBN | A | B | O | D | R | P | E | ATTR1 | ATTR2 | ATTR3 | CUSER | CDATE | UUSER | UDATE 

    Public Function InsertDo(ByVal strYEARMON As String, _
            Optional ByVal strGBN As String = OPTIONAL_STR, _
            Optional ByVal strAMTGBN As String = OPTIONAL_STR, _
            Optional ByVal strA As Double = OPTIONAL_NUM, _
            Optional ByVal strB As Double = OPTIONAL_NUM, _
            Optional ByVal strO As Double = OPTIONAL_NUM, _
            Optional ByVal strD As Double = OPTIONAL_NUM, _
            Optional ByVal strR As Double = OPTIONAL_NUM, _
            Optional ByVal strP As Double = OPTIONAL_NUM, _
            Optional ByVal strE As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR)


        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            BuildNameValues(",", "YEARMON", strYEARMON, strFields, strValues)
            BuildNameValues(",", "GBN", strGBN, strFields, strValues)
            BuildNameValues(",", "AMTGBN", strAMTGBN, strFields, strValues)
            BuildNameValues(",", "A", strA, strFields, strValues)
            BuildNameValues(",", "B", strB, strFields, strValues)
            BuildNameValues(",", "O", strO, strFields, strValues)
            BuildNameValues(",", "D", strD, strFields, strValues)
            BuildNameValues(",", "R", strR, strFields, strValues)
            BuildNameValues(",", "P", strP, strFields, strValues)
            BuildNameValues(",", "E", strE, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)
            BuildNameValues(",", "ATTR02", strATTR02, strFields, strValues)
            BuildNameValues(",", "ATTR03", strATTR03, strFields, strValues)
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
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Update ó��
    '���� : Key ���ǰ� Value Field����������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function UpdateDo(ByVal strYEARMON As String, _
            Optional ByVal strGBN As String = OPTIONAL_STR, _
            Optional ByVal strAMTGBN As String = OPTIONAL_STR, _
            Optional ByVal strA As Double = OPTIONAL_NUM, _
            Optional ByVal strB As Double = OPTIONAL_NUM, _
            Optional ByVal strO As Double = OPTIONAL_NUM, _
            Optional ByVal strD As Double = OPTIONAL_NUM, _
            Optional ByVal strR As Double = OPTIONAL_NUM, _
            Optional ByVal strP As Double = OPTIONAL_NUM, _
            Optional ByVal strE As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("GBN", strGBN), _
                        GetFieldNameValue("AMTGBN", strAMTGBN), _
                        GetFieldNameValue("A", strA), _
                        GetFieldNameValue("B", strB), _
                        GetFieldNameValue("O", strO), _
                        GetFieldNameValue("D", strD), _
                        GetFieldNameValue("R", strR), _
                        GetFieldNameValue("P", strP), _
                        GetFieldNameValue("E", strE), _
                        GetFieldNameValue("ATTR01", strATTR01), _
                        GetFieldNameValue("ATTR02", strATTR02), _
                        GetFieldNameValue("ATTR03", strATTR03), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("YEARMON", strYEARMON)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Delete ó��
    '���� : Key ������ ��������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function DeleteDo(Optional ByVal strYEARMON As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("YEARMON", strYEARMON)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function


#End Region

#Region "��ü ����/����"
    '*****************************************************************
    '�Է� : strInfoXML = ����⺻������ ���� XML
    'objSCGLSql = DB ó�� ��ü �ν��Ͻ� ����    '��ȯ : ����
    '��� : DB ó���� ���� ����⺻���� ����
    '*****************************************************************
    Public Sub New(Optional ByVal objSCGLConfig As SCGLUtil.cbSCGLConfig = Nothing, Optional ByVal strInfoXML As String = "")
        MyBase.SetConfig(objSCGLConfig, strInfoXML)
        MyBase.EntityName = "MD_REPORT_MSTTOTAL"     'Entity Name ����
    End Sub

    '���� ����� Base Class���� �����Ǿ� ����
#End Region

End Class