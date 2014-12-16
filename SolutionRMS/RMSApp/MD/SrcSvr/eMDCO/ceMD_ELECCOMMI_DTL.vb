'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - ��ƼƼ Ŭ���� ����Ŀ - ��ȭ S&C
'�ý��۱��� : �ַ�Ǹ�/�ý��۸�/Server Entity Class
'����  ȯ�� : GAC(Global Assembly Cache)
'���α׷��� : ceMD_ELECCOMMI_DTL.vb ( MD_ELECCOMMI_DTL Entity ó�� Class)
'��      �� : MD_ELECCOMMI_DTL Entity�� ����Insert/Update/Delete/Select�� ó��
'             - �θ�ƼƼ ��ü�� SCGLUtil.ceEntity�� ���
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2007-12-12 ���� 10:44:26 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '���� ��ƿ��Ƽ ��ü
Imports SCGLUtil.cbSCGLErr      '���� ����ó�� ��ü
Imports SCGLEntity              '��ƼƼ ��ü�� �θ� ��ü

Public Class ceMD_ELECCOMMI_DTL
    Inherits ceEntity

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "ceMD_ELECCOMMI_DTL"    '�ڽ��� Ŭ������
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Insert ó��
    '*****************************************************************
    'TRANSYEARMON,TRANSNO,SEQ,CLIENTCODE,MEDCODE,REAL_MED_CODE,DEPT_CD,DEMANDDAY,PRINTDAY,PRICE,CNT,AMT,SUSU,SUSURATE,TRU_TAX_FLAG,VAT,TRUST_YEARMON,TRUST_SEQ,MEMO,MED_FLAG
    Public Function InsertDo(ByVal strTRANSYEARMON As String, _
            ByVal dblTRANSNO As Double, _
            ByVal dblSEQ As Double, _
            ByVal strCLIENTCODE As String, _
            ByVal strMEDCODE As String, _
            ByVal strREAL_MED_CODE As String, _
            Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
            Optional ByVal strPRINTDAY As String = OPTIONAL_STR, _
            Optional ByVal dblCNT As Double = OPTIONAL_NUM, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUSU As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUSURATE As Double = OPTIONAL_NUM, _
            Optional ByVal strTRU_TAX_FLAG As String = OPTIONAL_STR, _
            Optional ByVal dblVAT As Double = OPTIONAL_NUM, _
            Optional ByVal strTRUST_YEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTRUST_SEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strMED_FLAG As String = OPTIONAL_STR, _
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
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            BuildNameValues(",", "TRANSYEARMON", strTRANSYEARMON, strFields, strValues)
            BuildNameValues(",", "TRANSNO", dblTRANSNO, strFields, strValues)
            BuildNameValues(",", "SEQ", dblSEQ, strFields, strValues)
            BuildNameValues(",", "CLIENTCODE", strCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "MEDCODE", strMEDCODE, strFields, strValues)
            BuildNameValues(",", "REAL_MED_CODE", strREAL_MED_CODE, strFields, strValues)
            BuildNameValues(",", "DEPT_CD", strDEPT_CD, strFields, strValues)
            BuildNameValues(",", "DEMANDDAY", strDEMANDDAY, strFields, strValues)
            BuildNameValues(",", "PRINTDAY", strPRINTDAY, strFields, strValues)
            BuildNameValues(",", "CNT", dblCNT, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "SUSU", dblSUSU, strFields, strValues)
            BuildNameValues(",", "SUSURATE", dblSUSURATE, strFields, strValues)
            BuildNameValues(",", "TRU_TAX_FLAG", strTRU_TAX_FLAG, strFields, strValues)
            BuildNameValues(",", "VAT", dblVAT, strFields, strValues)
            BuildNameValues(",", "TRUST_YEARMON", strTRUST_YEARMON, strFields, strValues)
            BuildNameValues(",", "TRUST_SEQ", dblTRUST_SEQ, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "MED_FLAG", strMED_FLAG, strFields, strValues)
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
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Update ó��
    '���� : Key ���ǰ� Value Field����������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function UpdateDo(ByVal strTRANSYEARMON As String, _
            ByVal dblTRANSNO As Double, _
            ByVal dblSEQ As Double, _
            ByVal strCLIENTCODE As String, _
            ByVal strMEDCODE As String, _
            ByVal strREAL_MED_CODE As String, _
            Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
            Optional ByVal strPRINTDAY As String = OPTIONAL_STR, _
            Optional ByVal dblPRICE As Double = OPTIONAL_NUM, _
            Optional ByVal dblCNT As Double = OPTIONAL_NUM, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUSU As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUSURATE As Double = OPTIONAL_NUM, _
            Optional ByVal strTRU_TAX_FLAG As String = OPTIONAL_STR, _
            Optional ByVal dblVAT As Double = OPTIONAL_NUM, _
            Optional ByVal strTRUST_YEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTRUST_SEQ As Double = OPTIONAL_NUM, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strMED_FLAG As String = OPTIONAL_STR, _
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
        Dim strNOW As String '����Ʈ���� ó���� ������ �޾� �ؽ�Ʈ�� ó�� �Ѵ�.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("CLIENTCODE", strCLIENTCODE), _
                        GetFieldNameValue("MEDCODE", strMEDCODE), _
                        GetFieldNameValue("REAL_MED_CODE", strREAL_MED_CODE), _
                        GetFieldNameValue("DEPT_CD", strDEPT_CD), _
                        GetFieldNameValue("DEMANDDAY", strDEMANDDAY), _
                        GetFieldNameValue("PRINTDAY", strPRINTDAY), _
                        GetFieldNameValue("PRICE", dblPRICE), _
                        GetFieldNameValue("CNT", dblCNT), _
                        GetFieldNameValue("AMT", dblAMT), _
                        GetFieldNameValue("SUSU", dblSUSU), _
                        GetFieldNameValue("SUSURATE", dblSUSURATE), _
                        GetFieldNameValue("TRU_TAX_FLAG", strTRU_TAX_FLAG), _
                        GetFieldNameValue("VAT", dblVAT), _
                        GetFieldNameValue("TRUST_SEQ", dblTRUST_SEQ), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("MED_FLAG", strMED_FLAG), _
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
                        GetFieldNameValue("TRANSYEARMON", strTRANSYEARMON), GetFieldNameValue("TRANSNO", dblTRANSNO), GetFieldNameValue("SEQ", dblSEQ)))

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
    Public Function DeleteDo(Optional ByVal strTRANSYEARMON As String = OPTIONAL_STR, Optional ByVal dblTRANSNO As Double = OPTIONAL_NUM, Optional ByVal dblSEQ As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("TRANSYEARMON", strTRANSYEARMON), GetFieldNameValue("TRANSNO", dblTRANSNO), GetFieldNameValue("SEQ", dblSEQ)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Select ó��
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


    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : ���ݰ�꼭 ��ȣ�� �ŷ����� ������ ATTR02�� �ִ´�.
    '���� : Key ������ ��������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function TRUST_InsertRtn(ByVal strSQL As String)
        Try
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".TRUST_InsertRtn")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL �� By KTH
    '��ȯ : ó���Ǽ�
    '��� : �ŷ����� ������ �� ���ݰ�꼭 ��ȣ ������Ʈ
    '*****************************************************************
    Public Function Update_CommiTax(ByVal strTRANSYEARMON As String, _
                                    ByVal lngTRANSNO As Double, _
                                    ByVal lngSEQ As Double, _
                                    ByVal strTAXYEARMON As String, _
                                    ByVal intTAXNO As Double) As Integer
        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_ELECCOMMI_DTL SET ATTR03 = 'Y' , ATTR02 = '" & strTAXYEARMON & "-" & intTAXNO & "'"
            strSQL = strSQL & "  WHERE TRANSYEARMON = '" & strTRANSYEARMON & "' AND TRANSNO =" & lngTRANSNO & " AND SEQ =" & lngSEQ
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_CommiTax")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL �� By KTH
    '��ȯ : ó���Ǽ�
    '��� : ������ �ŷ����� ������ �� ���ݰ�꼭 ��ȣ ����
    '*****************************************************************
    Public Function CommiTaxDeleteUpdateDo(ByVal strTAXYEARMON As String, _
                                           ByVal intTAXNO As Double) As Integer

        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_ELECCOMMI_DTL SET ATTR02 = ''"
            strSQL = strSQL & "  WHERE ATTR02 = '" & strTAXYEARMON & "-" & intTAXNO & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".CommiTaxDeleteUpdateDo")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ش� Entity�� Delete ó��
    '���� : Key ������ ��������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function UPDATE_COMMITAXNO_DO(ByVal strCOMMIYEARMON As String, _
                                         ByVal strMED_FLAG As String, _
                                         ByVal strCUSTCODE As String, _
                                         ByVal strREAL_MED_CODE As String, _
                                         ByVal strTRANSRANK As String, _
                                         ByVal strCOMMITAXNO As String) As Integer
        'strCOMMIYEARMON, strMED_FLAG, strCUSTCODE, strREAL_MED_CODE,strTRANSRANK
        Dim strSQL As String

        Try
            strSQL = "UPDATE MD_ELECTRIC_SUSUTEMP SET ATTR02 = '" & strCOMMITAXNO & "' "
            strSQL = strSQL & " WHERE YEARMON = '" & strCOMMIYEARMON & "' "
            strSQL = strSQL & " AND REAL_MED_CODE = '" & strREAL_MED_CODE & "' "



            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UPDATE_COMMITAXNO_DO")
        End Try
    End Function

    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ŷ������� ������ ����������TEMP �� �ŷ������� ��ȣ�� ����
    '���� : Key ������ ��������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function UPDATE_COMMI_DO(ByVal strCOMMIYEARMON As String, _
                                    ByVal strCLIENTCODE As String, _
                                    ByVal strREAL_MED_CODE As String, _
                                    ByVal strMED_FLAG As String, _
                                    ByVal strTRANSRANK As String, _
                                    ByVal strCOMMTRANSNO As String) As Integer
        'strCOMMIYEARMON, strCLIENTCODE, strREAL_MED_CODE, strMED_FLAG, strTRANSRANK, lngSUSU_DTL
        Dim strSQL As String

        Try

            strSQL = "UPDATE MD_ELECTRIC_SUSUTEMP SET ATTR01 = '" & strCOMMTRANSNO & "'"
            strSQL = strSQL & " WHERE YEARMON = '" & strCOMMIYEARMON & "' AND CLIENTCODE = '" & strCLIENTCODE & "'"
            strSQL = strSQL & " AND REAL_MED_CODE ='" & strREAL_MED_CODE & "' "
            strSQL = strSQL & " AND INPUT_MEDFLAG ='" & strMED_FLAG & "' "
            strSQL = strSQL & " AND TRANSRANK = '" & strTRANSRANK & "' "



            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UPDATE_TRANS_DO")
        End Try
    End Function
    'TRUST_UpdateRtn
    '*****************************************************************
    '�Է� : strSQL = SQL ��
    '��ȯ : ó���Ǽ�
    '��� : �ŷ������� ������ ��Ź �� �ŷ������� ��ȣ�� ����
    '���� : Key ������ ��������(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function TRUST_UpdateRtn(ByVal strCOMMIYEARMON As String) As Integer
        'strCOMMIYEARMON, strCLIENTCODE, strREAL_MED_CODE, strMED_FLAG, strTRANSRANK, lngSUSU_DTL
        Dim strSQL As String

        Try
            strSQL = "UPDATE MD_ELECTRIC_MEDIUM SET MD_ELECTRIC_MEDIUM.COMMI_TRANS_NO = B.ATTR01"
            strSQL = strSQL & " FROM MD_ELECTRIC_SUSUTEMP B"
            strSQL = strSQL & " WHERE MD_ELECTRIC_MEDIUM.YEARMON = B.YEARMON"
            strSQL = strSQL & " AND MD_ELECTRIC_MEDIUM.CLIENTCODE = B.CLIENTCODE"
            strSQL = strSQL & " AND MD_ELECTRIC_MEDIUM.REAL_MED_CODE = B.REAL_MED_CODE"
            strSQL = strSQL & " AND CASE MD_ELECTRIC_MEDIUM.INPUT_MEDFLAG WHEN '03' THEN '02' WHEN '20' THEN '10' ELSE MD_ELECTRIC_MEDIUM.INPUT_MEDFLAG END = B.INPUT_MEDFLAG"
            strSQL = strSQL & " AND MD_ELECTRIC_MEDIUM.YEARMON = '" & strCOMMIYEARMON & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".TRUST_UpdateRtn")
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
        MyBase.EntityName = "MD_ELECCOMMI_DTL"     'Entity Name ����
    End Sub

    '���� ����� Base Class���� �����Ǿ� ����
#End Region
#End Region
End Class

'------->>��ƼƼ INSERT/UPDATE �����Դϴ�. �ݵ�� �ڽ��� ȯ�濡 ���߾ �����Ͻñ� �ٶ��ϴ�.
'=========================================================
'       'vntData Array�� ����� �� Insert/Update �Դϴ�.
'=========================================================
'        Dim intRtn As Integer
'        intRtn = mobjceMD_ELECCOMMI_DTL.InsertDo( _
'                                       GetElement(vntData,"TRANSYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TRANSNO", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CLIENTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MEDCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"REAL_MED_CODE", intColCnt, intRow), _
'                                       GetElement(vntData,"DEPT_CD", intColCnt, intRow), _
'                                       GetElement(vntData,"DEMANDDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PRINTDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PRICE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CNT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"AMT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"SUSU", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"SUSURATE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TRU_TAX_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"VAT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TRUST_SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"MEMO", intColCnt, intRow), _
'                                       GetElement(vntData,"MED_FLAG", intColCnt, intRow), _
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
'        intRtn = mobjceMD_ELECCOMMI_DTL.UpdateDo( _
'                                       GetElement(vntData,"TRANSYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TRANSNO", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CLIENTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MEDCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"REAL_MED_CODE", intColCnt, intRow), _
'                                       GetElement(vntData,"DEPT_CD", intColCnt, intRow), _
'                                       GetElement(vntData,"DEMANDDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PRINTDAY", intColCnt, intRow), _
'                                       GetElement(vntData,"PRICE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CNT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"AMT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"SUSU", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"SUSURATE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TRU_TAX_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"VAT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TRUST_SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"MEMO", intColCnt, intRow), _
'                                       GetElement(vntData,"MED_FLAG", intColCnt, intRow), _
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
'       'XmlData �� ����� �� Insert/Update �Դϴ�.
'=========================================================
'        Dim intRtn As Integer
'        intRtn = mobjceMD_ELECCOMMI_DTL.InsertDo( _
'                                       XMLGetElement(xmlRoot,"TRANSYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TRANSNO", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CLIENTCODE"), _
'                                       XMLGetElement(xmlRoot,"MEDCODE"), _
'                                       XMLGetElement(xmlRoot,"REAL_MED_CODE"), _
'                                       XMLGetElement(xmlRoot,"DEPT_CD"), _
'                                       XMLGetElement(xmlRoot,"DEMANDDAY"), _
'                                       XMLGetElement(xmlRoot,"PRINTDAY"), _
'                                       XMLGetElement(xmlRoot,"PRICE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CNT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"AMT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"SUSU", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"SUSURATE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TRU_TAX_FLAG"), _
'                                       XMLGetElement(xmlRoot,"VAT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TRUST_SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"MEMO"), _
'                                       XMLGetElement(xmlRoot,"MED_FLAG"), _
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
'        intRtn = mobjceMD_ELECCOMMI_DTL.UpdateDo( _
'                                       XMLGetElement(xmlRoot,"TRANSYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TRANSNO", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CLIENTCODE"), _
'                                       XMLGetElement(xmlRoot,"MEDCODE"), _
'                                       XMLGetElement(xmlRoot,"REAL_MED_CODE"), _
'                                       XMLGetElement(xmlRoot,"DEPT_CD"), _
'                                       XMLGetElement(xmlRoot,"DEMANDDAY"), _
'                                       XMLGetElement(xmlRoot,"PRINTDAY"), _
'                                       XMLGetElement(xmlRoot,"PRICE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CNT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"AMT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"SUSU", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"SUSURATE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TRU_TAX_FLAG"), _
'                                       XMLGetElement(xmlRoot,"VAT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TRUST_SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"MEMO"), _
'                                       XMLGetElement(xmlRoot,"MED_FLAG"), _
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

