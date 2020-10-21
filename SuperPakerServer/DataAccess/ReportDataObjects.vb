Module ReportDataObjects

#Region "Struktura Widoku QWHV_RPT_DOC_OUT w Oracle"
    'DT.CODE AS DT_CODE, varchar2(25)
    'Did.D_DOC_ID, number not null
    'Did.D_DOC_NR, varchar2(25) not null
    'Did.D_DOC_TYPE_NR, varchar2(25) not null
    'Did.D_DATE_EMITTED, date not null
    '    Did.D_FIRM_ID, number
    '    Did.D_ADR_STREET, varchar2(100)
    '    Did.D_ADR_ZIPCODE, varchar2(100)
    '    Did.D_ADR_CITY, varchar2(100)
    '    Did.D_ADR_POSTOFFICE, varchar2(100)
    '    Did.D_C_ORDER_NR, varchar2(100)
    'Did.DATE_EXPIRE , date
    '    Did.D_INFO, varchar2(250)
    '    Did.D_STATUS, varchar2(2)
    '    'SH' as DI_OBJECT_TYPE, varchar2(2)
    'did.DI_QUANTITY, number not null
    'WH.WH_NR AS WH_WH_NR, varchar2(25) not null
    'WH.NAME AS WH_NAME, varchar2(50) not null
    'F.FIRM_ID AS F_FIRM_ID, nubmer not null
    'F.NAME AS F_NAME, varchar2 (50) not null
    'SU.INFO AS SU_INFO, varchar2 (250)
    'PS.SERIAL_NR AS PS_SERIAL_NR, varchar2(50) not null
    'PS.INFO AS PS_INFO, varchar2 (250)
    'SHh.SHIPMENT_NR as sh_shipment_nr, varchar2(25)
    'MU.ABBREV AS MU_ABBREV, varchar2(10) not null
    'P.PRODUCT_NR AS P_PRODUCT_NR, varchar2(25) not null
    'P.NAME AS P_NAME, varchar2(100) not null
    'PC.CODE_VALUE AS PC_CODE_VALUE, varchar2(25) not null
    'SA.SA_NR AS SA_SA_NR, varchar2(25)
    'SA.NAME AS SA_NAME, varchar2(50) not null
    '    did.object_id, number
    'dt.info as dt_info varchar2(250)
#End Region

    Public Function getQWHV_RPT_DOC_OUT_Tbl() As DataTable

        Dim QWHVRPTDOCOUT_Table As DataTable = New DataTable

        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.DT_CODE), GetType(String)) 'varchar2(25)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.D_DOC_ID), GetType(Integer)) 'number not null
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.D_DOC_NR), GetType(String)) 'varchar2(25) not null
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.D_DOC_TYPE_NR), GetType(String)) 'varchar2(25)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.D_DATE_EMITTED), GetType(DateTime)) 'date not null)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.D_FIRM_ID), GetType(Integer)) 'number
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.D_ADR_STREET), GetType(String)) 'varchar2(100)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.D_ADR_ZIPCODE), GetType(String)) 'varchar2(100)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.D_ADR_CITY), GetType(String)) 'varchar2(100)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.D_ADR_POSTOFFICE), GetType(String)) 'varchar2(100)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.D_C_ORDER_NR), GetType(String)) 'varchar2(100)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.DATE_EXPIRE), GetType(DateTime)) 'date
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.D_INFO), GetType(String)) 'varchar2(250)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.D_STATUS), GetType(String)) 'varchar2(2)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.DI_OBJECT_TYPE), GetType(String)) 'varchar2(2)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.DI_QUANTITY), GetType(Integer)) 'number not null
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.WH_WH_NR), GetType(String)) 'varchar2(25) not null
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.WH_NAME), GetType(String)) 'varchar2(50) not null
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.F_FIRM_ID), GetType(Integer)) 'number not null
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.F_NAME), GetType(String)) 'varchar2(50) not null
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.SU_INFO), GetType(String)) 'varchar2(250)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.PS_SERIAL_NR), GetType(String)) 'varchar2(250) not null
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.PS_INFO), GetType(String)) 'varchar2(250)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.SH_SHIPMENT_NR), GetType(String)) 'varchar2(25)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.MU_ABBREV), GetType(String)) 'varchar2(10) not null
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.P_PRODUCT_NR), GetType(String)) 'varchar2(25) not null
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.P_NAME), GetType(String)) 'varchar2(100) not null
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.PC_CODE_VALUE), GetType(String)) 'varchar2(25) not null
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.SA_SA_NR), GetType(String)) 'varchar2(25)
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.SA_NAME), GetType(String)) 'varchar2(50) not null
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.OBJECT_ID), GetType(Integer)) 'number
        QWHVRPTDOCOUT_Table.Columns.Add([Enum].GetName(GetType(QWHV_RPT_DOC_OUT_DBColumnsEnum), QWHV_RPT_DOC_OUT_DBColumnsEnum.DT_INFO), GetType(String)) 'varchar2(250)

        Return QWHVRPTDOCOUT_Table

    End Function

    'VIEW "QGUARADM"."QWHV_RPT_DOC_OUT" 
    Public Enum QWHV_RPT_DOC_OUT_DBColumnsEnum
        DT_CODE
        D_DOC_ID
        D_DOC_NR
        D_DOC_TYPE_NR
        D_DATE_EMITTED
        D_FIRM_ID
        D_ADR_STREET
        D_ADR_ZIPCODE
        D_ADR_CITY
        D_ADR_POSTOFFICE
        D_C_ORDER_NR
        DATE_EXPIRE
        D_INFO
        D_STATUS
        DI_OBJECT_TYPE
        DI_QUANTITY
        WH_WH_NR
        WH_NAME
        F_FIRM_ID
        F_NAME
        SU_INFO
        PS_SERIAL_NR
        PS_INFO
        SH_SHIPMENT_NR
        MU_ABBREV
        P_PRODUCT_NR
        P_NAME
        PC_CODE_VALUE
        SA_SA_NR
        SA_NAME
        OBJECT_ID
        DT_INFO
    End Enum

End Module
