using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using CompareRepairOrder.Model;
using System.Linq;

namespace CompareExcelItem.Service
{
    public class ImportFileService
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="txtFilePath"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="columns">讀取幾個欄位</param>
        /// <returns></returns>
        public DataTable ExcelToDataTable(string txtFilePath, int sheetIndex, int columns)
        {
            DataTable dt = new DataTable();
            FileInfo filePath = new FileInfo(txtFilePath);
            ExcelPackage ep = new ExcelPackage(filePath);
            ExcelWorksheet sheet = ep.Workbook.Worksheets[sheetIndex + 1];
            int startRowNumber = sheet.Dimension.Start.Row + 1;//起始列編號，從1算起
            int endRowNumber = sheet.Dimension.End.Row;//結束列編號，從1算起
            //int startColumn = 1; //sheet.Dimension.Start.Column;//開始欄編號，從1算起
            //int endColumn = 2; //sheet.Dimension.End.Column;//結束欄編號，

            //建立欄位名稱
            for (int k = 1; k <= columns; k++)
            {
                dt.Columns.Add(sheet.Cells[1, k].Value.ToString());
            }

            try
            {
                //寫入資料到資料列
                for (int currentRow = startRowNumber; currentRow <= endRowNumber; currentRow++)
                {
                    dt.NewRow();
                    object[] cell = new object[columns];
                    int idx = 0;
                    for (int i = 1; i <= columns; i++)
                    {
                        cell[idx] = sheet.Cells[currentRow, i].Value;
                        idx++;
                    }
                    dt.Rows.Add(cell);
                }

            }
            catch (Exception ex)
            {
                throw;
            }

            return dt;
        }

        public List<RepairOrder> ExcelToList(string excleFilePath, int columns)
        {
            FileInfo filePath = new FileInfo(excleFilePath);
            ExcelPackage ep = new ExcelPackage(filePath);
            List<RepairOrder> repairOrder = new List<RepairOrder>();
            List<RepairOrder> repairOrderList = new List<RepairOrder>();

            try
            {
                foreach (var sheet in ep.Workbook.Worksheets)
                {
                    repairOrder = SheetToList(ep, sheet.Index, columns);
                    repairOrderList.AddRange(repairOrder);
                }
               
            }
            catch (Exception )
            {
                throw;
            }

            return repairOrderList;
        }


        public List<RepairOrder> SheetToList(ExcelPackage ep, int sheetIndex, int columns)
        {
            ExcelWorksheet sheet = ep.Workbook.Worksheets[sheetIndex];
            int startRowNumber = sheet.Dimension.Start.Row + 1;//起始列編號，從1算起
            int endRowNumber = sheet.Dimension.End.Row;//結束列編號，從1算起
            //int startColumn = 1; //sheet.Dimension.Start.Column;//開始欄編號，從1算起
            //int endColumn = 2; //sheet.Dimension.End.Column;//結束欄編號，

            DataTable dt = new DataTable();
            //建立欄位名稱
            for (int k = 1; k <= columns; k++)
            {
                dt.Columns.Add(sheet.Cells[1, k].Value.ToString());
            }

            try
            {
                //寫入資料到資料列
                for (int currentRow = startRowNumber; currentRow <= endRowNumber; currentRow++)
                {
                    dt.NewRow();
                    object[] cell = new object[columns];
                    int idx = 0;
                    for (int i = 1; i <= columns; i++)
                    {
                        cell[idx] = sheet.Cells[currentRow, i].Value;
                        idx++;
                    }
                    dt.Rows.Add(cell);
                }

                List<RepairOrder> repairOrders = new List<RepairOrder>();
                foreach (DataRow dr in dt.Rows)
                {
                    RepairOrder data = new RepairOrder();

                    if (dr["整盒新品料號"].ToString().Length > 2  && dr["維修單號"].ToString().Length > 2)
                    {
                        data.NewPartNumber = dr["整盒新品料號"].ToString();
                        data.RepairOrderNumber = dr["維修單號"].ToString();
                        repairOrders.Add(data);
                    }
                }

                return repairOrders;
            }
            catch (Exception ex)
            {
                throw;
            }

            
        }

            //public DataTable CompareRevisionAndReturn(DataTable dtRevision, List<RepairOrder> returnList)
            //{
            //    try
            //    {
            //        for (int i = 0; i < dtRevision.Rows.Count; i++)
            //        {
            //            //去退料檔找序號
            //            List<RepairOrder> findList = returnList.Where(x => x.Model == dtRevision.Rows[i]["品號"].ToString().Replace("-C", "")).ToList();
            //
            //            dtRevision.Rows[i][3] = findList.Count;
            //
            //            int serialIndex = 4;
            //            string strIndex = "\"" + serialIndex.ToString() + "\"";
            //            foreach (var item in findList)
            //            {
            //                dtRevision.Rows[i][serialIndex] = item.SerialNumber;
            //                serialIndex++;
            //            }
            //
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        throw ex;
            //    }
            //
            //    return dtRevision;
            //}


        }
}
