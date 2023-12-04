using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScriptsExecutionUtility.Models
{
    public class BillPeriodDto
    {
        public string BillCycleID { get; set; }
        public string BillCycleStartDate { get; set; }
        public string BillCycleEndDate { get; set; }
        public string FreezeDate { get; set; }
        public string BillDate { get; set; }
    }
}
