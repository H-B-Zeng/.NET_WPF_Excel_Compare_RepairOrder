using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompareRepairOrder.Model
{
    public class RepairOrder
    {
        /// <summary>
        /// 整盒新品料號
        /// </summary>
        public string NewPartNumber { get; set; }

        /// <summary>
        /// 維修單號
        /// </summary>
        public string RepairOrderNumber { get; set; }

        /// <summary>
        /// 數量
        /// </summary>
        public string Quantity { get; set; }


    }
}
