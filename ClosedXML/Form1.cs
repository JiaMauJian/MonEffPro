using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace ClosedXML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("test");
            worksheet.Cell("A1").Value = "Hello World!";
            workbook.SaveAs("HelloWork.xlsx");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var svr = new MmsRepSvr.Service();

            string sqlCmd = @"SELECT A.*,W.QV_NO,V.VENDOR_MSERIES,TX_BOOKINTIME,TX_BOOKOUTTIME,DF_BOOKINTIME,DF_BOOKOUTTIME,HF_BOOKINTIME,HF_BOOKOUTTIME,HF_MACHID,PC_BOOKINTIME,PC_BOOKOUTTIME,PT_BOOKOUTTIME,PT_MACHID,CT_BOOKOUTTIME,CT_MACHID,P.PROC_CODE_NAME,D.PCS_ZIR2,D.PCS_ZIR5,D.PCS_TOTAL,D.PCS_ABZ_RSH_30,W.OEM_SUPPLIER,E.PASTE_TYPE,W.TTP_QV,ipa_BOOKINTIME,ipa_BOOKOUTTIME,ipa_MACHID,ald_BOOKINTIME,ald_BOOKOUTTIME,ald_MACHID,lsr_BOOKINTIME,lsr_BOOKOUTTIME,lsr_MACHID,il_BOOKINTIME,il_BOOKOUTTIME,il_MACHID
FROM 

(SELECT LOT_ID,
        ipa_BOOKINTIME,
        ipa_BOOKOUTTIME,
        ipa_FAB||ipa_MACHID_LIST AS ipa_MACHID
 FROM EDWADM.MEDA_MMS_ipa_P@DBLINK_EDWUSER_28
 WHERE LOT_ID IS NOT NULL
AND TIMESTAMP > SYSDATE - 7 )ipa,
(SELECT LOT_ID,
        ald_BOOKINTIME,
        ald_BOOKOUTTIME,
        ald_FAB||ald_MACHID_LIST AS ald_MACHID
 FROM EDWADM.MEDA_MMS_ald_P@DBLINK_EDWUSER_28
 WHERE LOT_ID IS NOT NULL
AND TIMESTAMP > SYSDATE - 7 )ald,
(SELECT LOT_ID,
        TX_BOOKINTIME,
        TX_BOOKOUTTIME,
        TX_FAB||TX_MACHID_LIST AS TX_MACHID
 FROM EDWADM.MEDA_MMS_TX_P@DBLINK_EDWUSER_28
 WHERE LOT_ID IS NOT NULL
AND TIMESTAMP > SYSDATE - 7 )TX,
 (SELECT LOT_ID,
        DF_BOOKINTIME,
        DF_BOOKOUTTIME,
        DF_FAB||DF_MACHID_LIST AS DF_MACHID
 FROM EDWADM.MEDA_MMS_DF_P@DBLINK_EDWUSER_28
 WHERE LOT_ID IS NOT NULL
AND TIMESTAMP > SYSDATE - 7 )DF,
 (SELECT LOT_ID,
        PC_BOOKINTIME,
        PC_BOOKOUTTIME,
        PC_FAB||PC_MACHID_LIST AS PC_MACHID
 FROM EDWADM.MEDA_MMS_PC_P@DBLINK_EDWUSER_28
 WHERE LOT_ID IS NOT NULL
AND TIMESTAMP > SYSDATE - 7 )PC,
(SELECT LOT_ID,
        HF_BOOKINTIME,
        HF_BOOKOUTTIME,
        HF_FAB||HF_MACHID_LIST AS HF_MACHID
 FROM EDWADM.MEDA_MMS_HF_P@DBLINK_EDWUSER_28
 WHERE LOT_ID IS NOT NULL
AND TIMESTAMP > SYSDATE - 7 )HF,
 (SELECT LOT_ID,
         PT_BOOKINTIME,
         PT_BOOKOUTTIME,
        PT_FAB||PT_MACHID_LIST AS PT_MACHID
 FROM EDWADM.MEDA_MMS_PT_P@DBLINK_EDWUSER_28
 WHERE LOT_ID IS NOT NULL
AND TIMESTAMP > SYSDATE - 7 )PT,
(SELECT LOT_ID,
        CT_BOOKOUTTIME,
        CT_FAB||CT_MACHID_LIST AS CT_MACHID
 FROM EDWADM.MEDA_MMS_CT_P@DBLINK_EDWUSER_28
 WHERE LOT_ID IS NOT NULL
AND TIMESTAMP > SYSDATE - 7 )CT,
 (SELECT LOT_ID,
         lsr_BOOKINTIME,
         lsr_BOOKOUTTIME,
        lsr_FAB||lsr_MACHID_LIST AS lsr_MACHID
 FROM EDWADM.MEDA_MMS_lsr_P@DBLINK_EDWUSER_28
 WHERE LOT_ID IS NOT NULL
AND TIMESTAMP > SYSDATE - 7 )lsr,
 (SELECT LOT_ID,
         il_BOOKINTIME,
         il_BOOKOUTTIME,
        il_FAB||il_MACHID_LIST AS il_MACHID
 FROM EDWADM.MEDA_MMS_il_P@DBLINK_EDWUSER_28
 WHERE LOT_ID IS NOT NULL
AND TIMESTAMP > SYSDATE - 7 )il,
 (SELECT LOTNO,
         SUM(PCS_ZIR2) AS PCS_ZIR2,
         SUM(PCS_ZIR5) AS PCS_ZIR5,
         SUM(PCS_TOTAL) AS PCS_TOTAL,
         SUM(PCS_ABZ_RSH_30) AS PCS_ABZ_RSH_30
 FROM EDAADM.BERGER_IREV_RSH  D
 WHERE LOTNO IS  NOT NULL
 AND CT_DATE > SYSDATE-7
 GROUP BY LOTNO) D,
 edwadm.REPORT_MORNING_MEETING_EFF@DBLINK_EDWUSER_28 A, EDWADM.EDW_DIM_WO@DBLINK_EDWUSER_28 W,EDWADM.EDW_DIM_VENDOR@DBLINK_EDWUSER_28 V,edw_dim_proc_code@dblink_EDWADM_28 P,EDWADM.EDW_DIM_LOTINFO_CELL@DBLINK_EDWUSER_28 E
WHERE  a.DAYKEY BETWEEN TO_CHAR(SYSDATE-7,'YYYYMMDD') AND TO_CHAR(SYSDATE-0,'YYYYMMDD')
AND A.WORKORDER=W.WO_ID(+)
AND A.LOTNO=TX.LOT_ID(+)
AND A.LOTNO=PC.LOT_ID(+)
AND A.LOTNO=DF.LOT_ID(+)
AND A.LOTNO=HF.LOT_ID(+)
AND A.LOTNO=PT.LOT_ID(+)
AND A.LOTNO=CT.LOT_ID(+)
AND A.LOTNO=ipa.LOT_ID(+)
AND A.LOTNO=ald.LOT_ID(+)
AND A.LOTNO=lsr.LOT_ID(+)
AND A.LOTNO=il.LOT_ID(+)
AND A.LOTNO=D.LOTNO(+)
AND A.LOTNO=E.LOTNO(+)
AND A.REGION=W.REGION
AND A.PROC_CODE=P.PROC_CODE(+)
--AND W.WO_CODE IN ('P1','P2','P7','P9','2C','ZP01','ZP02','ZE01','ZR01','ZR06') AND SUBSTR(W.PART_ITEM_NO,12,1) <>'S'
--AND substr(A.VENDERID,1,2) NOT IN ('MD','MB')
AND A.EFF_U<>'0'
AND A.VENDERID=V.VENDOR
AND A.FAB NOT IN 'F03'
--and a.ag not in 'NA'
AND SUBSTR(A.LOTNO,18,1)='0'--重工不抓取
AND SUBSTR(A.LOTNO,4,1)<>'A'-- a CELL LOT不抓取
--AND a.output >=500
And a.venderid<> 'RE'";

            var dt = NoOraConn.Query(sqlCmd);
            dt.TableName = "Monitor_EFF_2018";
            var workbook = new XLWorkbook();
            workbook.Worksheets.Add(dt);
            workbook.SaveAs("DataTableToExcel.xlsx");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var workbook = new XLWorkbook("DataTableToExcelPivot.xlsx");
            var sheet = workbook.Worksheet(1);
            var range = sheet.Range(sheet.FirstCellUsed(), sheet.LastCellUsed());

            var ptSheet = workbook.Worksheets.Add("PivotTalbe");
            var pt = ptSheet.PivotTables.AddNew("PivotTable", ptSheet.Cell(1, 1), range);

            pt.RowLabels.Add("VENDERID");
            pt.ColumnLabels.Add("OP");

            pt.Values.Add("EFF_U").SetSummaryFormula(XLPivotSummary.Average);

            // https://github.com/ClosedXML/ClosedXML/issues/759
            workbook.SaveAs("Pivot.xlsx");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Excel -> DataTable
            DataTable dtImport = ImportExcel("DataTableToExcel.xlsx");

            // DataTable -> Class
            List<MonEff> monEffList = dtImport.DataTableToList<MonEff>();

            // 按時間排序
            var monEffListOrderByTime = from m in monEffList orderby m.START_TIME select m;

            // 取得欄位條件組合
            var conditions = (from m in monEffListOrderByTime
                              orderby m.PASTE_TYPE
                              where m.FAB == "F02"
                              group m by new
                              {
                                  m.FAB, m.PASTE_TYPE, m.VENDERID, m.OEM_SUPPLIER, m.BUS, m.CT
                              } into m
                              select new MonEff
                              {
                                  FAB = m.Key.FAB,
                                  PASTE_TYPE = m.Key.PASTE_TYPE,
                                  VENDERID = m.Key.VENDERID,
                                  OEM_SUPPLIER = m.Key.OEM_SUPPLIER,
                                  BUS = m.Key.BUS,
                                  CT = m.Key.CT,
                              }).ToList();

            // debug
            foreach (var c in conditions)
            {
                Console.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t", c.FAB, c.PASTE_TYPE, c.VENDERID, c.OEM_SUPPLIER, c.BUS, c.CT);
            }

            // 連續三點下降的組合
            List<MonEff> results = new List<MonEff>();
            foreach (var c in conditions)
            {
                if (c.CT.Length == 0) continue;

                var filter = (from m in monEffListOrderByTime
                              where m.FAB == c.FAB
                              where m.PASTE_TYPE == c.PASTE_TYPE
                              where m.VENDERID == c.VENDERID
                              where m.OEM_SUPPLIER == c.OEM_SUPPLIER
                              where m.BUS == c.BUS
                              where m.CT == c.CT
                              select m).ToList();

                if (filter.Count() >= 3)
                {
                    int l = filter.Count();
                    if ((filter[l - 1].EFF_U < filter[l - 2].EFF_U) && (filter[l - 2].EFF_U < filter[l - 3].EFF_U))
                    {
                        Console.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t", filter[0].FAB, filter[0].PASTE_TYPE, filter[0].VENDERID, filter[0].OEM_SUPPLIER, filter[0].BUS, filter[0].CT);
                        results.Add(filter[0]);
                    }
                }
            }

            // Save Filter Results to Excel
            var workbook = new XLWorkbook("DataTableToExcel.xlsx");
            
            // Conditions Sheet
            var condSheet = workbook.Worksheets.Add("Conditions");
            condSheet.Cells("A1").Value = "FAB";
            condSheet.Cells("B1").Value = "PASTE_TYPE";
            condSheet.Cells("C1").Value = "VENDERID";
            condSheet.Cells("D1").Value = "OEM_SUPPLIER";
            condSheet.Cells("E1").Value = "BUS";
            condSheet.Cells("F1").Value = "CT";
            int i = 2;
            foreach (var result in results)
            {
                condSheet.Cells("A" + i.ToString()).Value = result.FAB;
                condSheet.Cells("B" + i.ToString()).Value = result.PASTE_TYPE;
                condSheet.Cells("C" + i.ToString()).Value = result.VENDERID;
                condSheet.Cells("D" + i.ToString()).Value = result.OEM_SUPPLIER;
                condSheet.Cells("E" + i.ToString()).Value = result.BUS;
                condSheet.Cells("F" + i.ToString()).Value = result.CT;
                i++;
            }

            // By Each Condition Sheet
            i = 1;
            foreach (var result in results)
            {
                var filter = from m in monEffListOrderByTime
                             where m.FAB == result.FAB
                             where m.PASTE_TYPE == result.PASTE_TYPE
                             where m.VENDERID == result.VENDERID
                             where m.OEM_SUPPLIER == result.OEM_SUPPLIER
                             where m.BUS == result.BUS
                             where m.CT == result.CT                             
                             select m;

                // Class -> DataTable -> Excel
                DataTable dtFilter = filter.CreateDataTable<MonEff>();
                dtFilter.TableName = "c" + i.ToString();
                workbook.Worksheets.Add(dtFilter);
                i++;
            }
            workbook.SaveAs("DataTableToExcelWithFilter.xlsx");            
        }

        public DataTable ImportExcel(string filePath)
        {
            //Open the Excel file using ClosedXML.
            using (XLWorkbook workbook = new XLWorkbook(filePath))
            {
                IXLWorksheet sheet = workbook.Worksheet(1);

                DataTable dt = new DataTable();

                bool firstRow = true;
                foreach (IXLRow row in sheet.Rows())
                {
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        dt.Rows.Add();
                        int i = 0;

                        // 避免空值跳位 row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber
                        foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }

                return dt;
            }
        }
    }
}
