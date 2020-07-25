using CompareRepairOrder.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompareRepairOrder.Service
{
    public class CommonService
    {

        public DataTable RepairOrderListToDataTable(List<RepairOrder> newPartList, string NewPartNumberColName, string RepairOrderColName)
        {
            DataTable dt = new DataTable();

            try
            {
                //整盒新品料號 在維修單號  最多出現幾次 maxColumns，用來顯示Excel有幾欄
                int repairOrderColumns = newPartList.Max(x => x.Quantity);
                int maxColumns = repairOrderColumns +2;
                
                dt.Columns.Add(NewPartNumberColName);
                dt.Columns.Add("Quantity");//數量

                //建立欄位名稱
                for (int k = 1; k <= repairOrderColumns; k++)
                    dt.Columns.Add(RepairOrderColName + "-" + k.ToString());
               
                //整盒新品 newPartList
                foreach (var newPart  in newPartList)
                {
                    //寫入資料到資料列
                    dt.NewRow();
                    object[] cell = new object[maxColumns];
                    cell[0] = newPart.NewPartNumber;
                    cell[1] = newPart.RepairOrderNumberList.Count();//整盒新品料號 在 維修單號 出現幾次
                    int idx = 2;
                    foreach (string orderNumber in newPart.RepairOrderNumberList)
                    {
                        cell[idx] = orderNumber;
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
    }
}
