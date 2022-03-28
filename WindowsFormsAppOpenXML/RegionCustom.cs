using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsAppOpenXML
{
    public class RegionCustom
    {
        public int Codigo { get; set; }
        public string Region { get; set; }
        public List<Pessoa> pessoas { get; set; }

        public RegionCustom()
        {
            pessoas = new List<Pessoa>();
        }

        public override string ToString()
        {
            return Region;
        }
    }
}
