using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOM_Version.Model
{
    public class BomLines
    {
        public String ItemCode { get; set; }
        public int Quantity { get; set; }
        public String WhsCode { get; set; }
    }
}
