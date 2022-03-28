using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WindowsFormsAppOpenXML
{
    public class RegionCustom
    {
        public int Codigo { get; set; }
        public string Region { get; set; }
        public int PositionInitial { get; set; }
        public List<Pessoa> Pessoas { get; set; }
        public List<ParagraphCustom> ParagraphCustomBefore { get; set; }
        public List<ParagraphCustom> ParagraphCustomFull { get; set; }
        public List<ParagraphCustom> ParagraphCustomMiddle { get; set; }
        public List<ParagraphCustom> ParagraphCustomAfter { get; set; }
        public List<ParagraphCustom> ParaghapsCustom { get; set; }

        public RegionCustom()
        {
            Pessoas = new List<Pessoa>();
            ParaghapsCustom = new List<ParagraphCustom>();
            ParagraphCustomBefore = new List<ParagraphCustom>();
            ParagraphCustomFull = new List<ParagraphCustom>();
            ParagraphCustomMiddle = new List<ParagraphCustom>();
            ParagraphCustomAfter = new List<ParagraphCustom>();
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

        public void AddParagraphRangeComplex(List<ParagraphCustom> paragraphsCustom)
        {
            ParagraphCustomBefore.Clear();
            ParagraphCustomFull.Clear();
            ParagraphCustomMiddle.Clear();
            ParagraphCustomAfter.Clear();

            if(paragraphsCustom.Count > 0)
            {
                PositionInitial = paragraphsCustom[0].Key;
            }
            
            string openTag = "<r-" + Region + ">";
            string closeTag = "</r-" + Region + ">";
            int LenghtCloseTag = closeTag.Length;

            bool pbefore_removeAfter = false;
            bool pmiddle_removeBefore = true;
            bool pmiddle_removeAfter = false;
            bool pafter_removeBefore = true;

            foreach (var item in paragraphsCustom)
            {

                // add Custom Before
                {
                    ParagraphCustom pbefore = new ParagraphCustom();
                    pbefore.Clone(item.Key, item.Paragraph);

                    //bool removeAfter = false;

                    var texts = pbefore.Paragraph.Descendants<Text>();
                    foreach (var text in texts)
                    {                        
                        if (pbefore_removeAfter == true)
                        {
                            text.Remove();
                        }

                        if (text.Text.Contains(closeTag))
                        {
                            pbefore_removeAfter = true;
                            int firstCloseTage = text.Text.IndexOf(closeTag) + LenghtCloseTag;
                            // remove no próprio texto palavras depois da tag da região
                            text.Text = text.Text.Substring(0, firstCloseTage);
                        }
                    }

                    ParagraphCustomBefore.Add(pbefore);
                }

                // add Custom Full
                {
                    ParagraphCustom pfull = new ParagraphCustom();
                    pfull.Clone(item.Key, item.Paragraph);
                    ParagraphCustomFull.Add(pfull);
                }

                // add Custom Middle
                {
                    ParagraphCustom pmiddle = new ParagraphCustom();
                    pmiddle.Clone(item.Key, item.Paragraph);

                    //bool removeBefore = true;
                    //bool removeAfter = false;

                    var texts = pmiddle.Paragraph.Descendants<Text>();
                    foreach (var text in texts)
                    {
                        if (text.Text.Contains(openTag))
                        {
                            pmiddle_removeBefore = false;
                            int firstOpenTag = text.Text.IndexOf(openTag);
                            int lenghtText = text.Text.Count() - firstOpenTag;
                            // remove no próprio texto palavras antes da tag da região
                            text.Text = text.Text.Substring(firstOpenTag, lenghtText);                           
                        }

                        if (pmiddle_removeBefore == true)
                        {
                            text.Remove();
                        }

                        if (pmiddle_removeAfter == true)
                        {
                            text.Remove();
                        }

                        if (text.Text.Contains(closeTag))
                        {
                            pmiddle_removeAfter = true;                     
                            int firstCloseTage = text.Text.IndexOf(closeTag) + LenghtCloseTag;
                            // remove no próprio texto palavras depois da tag da região
                            text.Text = text.Text.Substring(0, firstCloseTage); 
                        }
                    }

                    ParagraphCustomMiddle.Add(pmiddle);
                }

                // add Custom After
                {
                    ParagraphCustom pafter = new ParagraphCustom();
                    pafter.Clone(item.Key, item.Paragraph);

                    //bool removeBefore = true;

                    var texts = pafter.Paragraph.Descendants<Text>();
                    foreach (var text in texts)
                    {
                        if (text.Text.Contains(openTag))
                        {
                            pafter_removeBefore = false;
                            int firstOpenTag = text.Text.IndexOf(openTag);
                            int lenghtText = text.Text.Count() - firstOpenTag;
                            // remove no próprio texto palavras antes da tag da região
                            text.Text = text.Text.Substring(firstOpenTag, lenghtText);
                        }

                        if (pafter_removeBefore == true)
                        {
                            text.Remove();
                        }
                    }
                    ParagraphCustomAfter.Add(pafter);
                }
            }
        }

        public override string ToString()
        {
            return Region;
        }
    }
}
