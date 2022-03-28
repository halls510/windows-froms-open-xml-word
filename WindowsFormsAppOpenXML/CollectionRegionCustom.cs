using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsAppOpenXML
{
    public class CollectionRegionCustom
    {
        private List<RegionCustom> collection;

        public CollectionRegionCustom()
        {
            collection = new List<RegionCustom>();
        }

        public void Set(RegionCustom regionCustom)
        {
            collection.Add(regionCustom);
        }

        public void AddPessoa(int Codigo, Pessoa pessoa)
        {
            var c = collection.Where(w => w.Codigo == Codigo).FirstOrDefault();
            if(c != null)
            {
                c.pessoas.Add(pessoa);
            }
        }

        public List<RegionCustom> Get()
        {
            return collection;
        }
    }
}
