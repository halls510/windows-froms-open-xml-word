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
        public List<Pessoa> Pessoas { get; set; }
        public List<ParagraphCustom> ParaghapsCustom { get; set; }

        public RegionCustom()
        {
            Pessoas = new List<Pessoa>();
            ParaghapsCustom = new List<ParagraphCustom>();
        }

        public void AddParagraph(ParagraphCustom paragraphCustom)
        {
            ParagraphCustom p = new ParagraphCustom();
            p.Clone(paragraphCustom.Key, paragraphCustom.Paragraph);
            ParaghapsCustom.Add(p);
        }

        public void AddParagraphRange(List<ParagraphCustom> paragraphsCustom)
        {
            ParaghapsCustom.Clear();
            foreach (var item in paragraphsCustom)
            {
                ParagraphCustom p = new ParagraphCustom();
                p.Clone(item.Key, item.Paragraph);
                ParaghapsCustom.Add(p);
            }            
        }

        public override string ToString()
        {
            return Region;
        }
    }
}
