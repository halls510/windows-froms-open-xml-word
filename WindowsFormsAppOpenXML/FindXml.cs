using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsAppOpenXML
{

    public class ParagraphCustom
    {
        public int Key { get; set; }
        public Paragraph Paragraph { get; set; }

        public ParagraphCustom()
        {
        }

        public ParagraphCustom(int key, Paragraph paragraph)
        {
            Key = key;
            Paragraph = paragraph;
        }

        public void Clone(int key, Paragraph paragraph)
        {
            Key = key;
            Paragraph = (Paragraph)paragraph.Clone();
        }
    }

    public class TextCustom
    {
        public int Key { get; set; }
        public Text Text { get; set; }
    }

    public class CollectionTextCustom
    {
        private List<TextCustom> collection;

        public CollectionTextCustom()
        {
            collection = new List<TextCustom>();
        }

        public void Set(int key, Text text)
        {
            collection.Add(new TextCustom { Key = key, Text = text });
        }

        public List<TextCustom> Get()
        {
            return collection;
        }

    }

    public static class FindXml
    {
        public static Document FindReplace(this Document document, string oldValue, string newValue)
        {
            foreach (Text text in document.Descendants<Text>())
            {
                if (text.Text.Contains(oldValue))
                {
                    text.Text = text.Text.Replace(oldValue, newValue);
                }
            }
            return document;
        }

        public static CollectionTextCustom FindRegion(this Document document, string search)
        {
            CollectionTextCustom collection = new CollectionTextCustom();

            string openTag = "<" + search + ">";
            string closeTag = "</" + search + ">";

            bool start = false;
            int i = 0;
            foreach (Text text in document.Descendants<Text>())
            {
                if (text.Text.Contains(openTag) == true && start == false)
                {
                    start = true;
                    collection.Set(i, text);
                }
                else if (start == true && !text.Text.Contains(closeTag))
                {
                    collection.Set(i, text);
                }
                else if (start == true && text.Text.Contains(closeTag))
                {
                    collection.Set(i, text);
                    start = false;
                }

                i++;
            }
            return collection;
        }

        public static List<ParagraphCustom> FindRegionTest(this Document document, string search)
        {
            List<ParagraphCustom> collection = new List<ParagraphCustom>();

            string openTag = "<r-" + search + ">";
            string closeTag = "</r-" + search + ">";

            int countStart = 0;
            int countFinish = 0;
            int i = 0;

            foreach (Paragraph paragraph in document.Descendants<Paragraph>())
            {
                var texts = paragraph.Descendants<Text>();
                foreach (var text in texts)
                {
                    if (text.Text.Contains(openTag))
                    {
                        countStart++;
                    }

                    if (text.Text.Contains(closeTag))
                    {
                        countFinish++;
                    }
                }

                if (countStart == 1)
                {
                    ParagraphCustom paragraphCustom = new ParagraphCustom();
                    paragraphCustom.Clone(i, paragraph);

                    collection.Add(paragraphCustom);
                }

                if (countFinish == 1)
                {
                    break;
                }

                i++;
            }
            return collection;
        }

        #region RegionComplex

        public static List<ParagraphCustom> FindRegionAndRemoveParagraphComplex(this Document document, string search)
        {
            List<ParagraphCustom> collection = new List<ParagraphCustom>();

            string openTag = "<r-" + search + ">";
            string closeTag = "</r-" + search + ">";

            int countStart = 0;
            int countFinish = 0;
            int i = 0;

            List<int> removes = new List<int>();

            foreach (Paragraph paragraph in document.Descendants<Paragraph>())
            {
                var texts = paragraph.Descendants<Text>();
                foreach (var text in texts)
                {
                    if (text.Text.Contains(openTag))
                    {
                        countStart++;
                    }

                    if (text.Text.Contains(closeTag))
                    {
                        countFinish++;
                    }
                }

                if (countStart == 1)
                {
                    ParagraphCustom paragraphCustom = new ParagraphCustom();
                    paragraphCustom.Clone(i, paragraph);
                    collection.Add(paragraphCustom);
                    removes.Add(i);
                }

                if (countFinish == 1)
                {
                    break;
                }
                i++;
            }

            var paragraphsDocuments = document.Descendants<Paragraph>().ToList();
            foreach (var remove in removes)
            {
                paragraphsDocuments[remove].Remove();
            }

            return collection;
        }

        #endregion

        public static List<ParagraphCustom> RemoveTextNoRegion(this RegionCustom regionCustom)
        {
            List<ParagraphCustom> collection = new List<ParagraphCustom>();

            string openTag = "<r-" + regionCustom.Region + ">";
            string closeTag = "</r-" + regionCustom.Region + ">";

            bool removeBefore = true;
            bool removeAfter = false;

            foreach (ParagraphCustom pc in regionCustom.ParaghapsCustom)
            {
                var texts = pc.Paragraph.Descendants<Text>();
                foreach (var text in texts)
                {
                    if (text.Text.Contains(openTag))
                    {
                        removeBefore = false;
                    }

                    if (removeBefore == true)
                    {
                        text.Remove();

                    }

                    if (removeAfter == true)
                    {
                        text.Remove();
                    }

                    if (text.Text.Contains(closeTag))
                    {
                        removeAfter = true;
                    }
                }

                ParagraphCustom paragraphCustom = new ParagraphCustom();
                paragraphCustom.Clone(pc.Key, pc.Paragraph);
                collection.Add(paragraphCustom);

            }
            return collection;
        }



        public static int? FindRegionReplaceTrue(this Document document, string search)
        {
            int? result = null;

            string closeTag = "</r-" + search + " replace=\"true\">";

            int countFinish = 0;
            int i = 0;

            foreach (Paragraph paragraph in document.Descendants<Paragraph>())
            {
                var texts = paragraph.Descendants<Text>();
                foreach (var text in texts)
                {
                    if (text.Text.Contains(closeTag))
                    {
                        countFinish = 1;
                    }
                }

                if (countFinish == 1)
                {
                    result = i;
                    countFinish = 0;
                }

                i++;
            }
            return result;
        }

        public static Document FindReplaceRegion(this Document document, CollectionTextCustom textsCustom, string oldValue, string newValue)
        {
            int countTextsCustom = textsCustom.Get().Count();
            if (countTextsCustom > 0)
            {

                int i = 0;
                foreach (Text text in document.Descendants<Text>())
                {
                    Text textCustom = textsCustom.Get().Where(w => w.Key == i).Select(s => s.Text).FirstOrDefault();
                    if (textCustom != null)
                    {
                        if (text.Text == textCustom.Text)
                        {
                            if (text.Text.Contains(oldValue))
                            {
                                text.Text = text.Text.Replace(oldValue, newValue);
                            }
                        }
                    }
                    i++;
                }
            }
            return document;
        }


        public static Document FindReplaceRegionContentTest(this Document document, List<ParagraphCustom> paragraphsCustom, string oldValue, string newValue, string searchRegion)
        {
            int countParagraphsCustom = paragraphsCustom.Count();
            if (countParagraphsCustom > 0)
            {
                string openOldTag = "<r-" + searchRegion + ">";
                string openTag = "<r-" + searchRegion + " replace=\"true\">";
                string closeOldTag = "</r-" + searchRegion + ">";
                string closeTag = "</r-" + searchRegion + " replace=\"true\">";

                var paragraphsDocument = document.Descendants<Paragraph>().ToList();
                int i = 0;
                foreach (ParagraphCustom paragraphCustom in paragraphsCustom)
                {
                    var p = paragraphsDocument[paragraphCustom.Key];

                    foreach (var text in p.Descendants<Text>())
                    {
                        if (text.Text.Contains(openOldTag))
                        {
                            text.Text = text.Text.Replace(openOldTag, openTag);
                        }

                        if (text.Text.Contains(closeOldTag))
                        {
                            text.Text = text.Text.Replace(closeOldTag, closeTag);
                        }

                        if (text.Text.Contains(oldValue))
                        {
                            text.Text = text.Text.Replace(oldValue, newValue);
                        }
                    }
                    i++;
                }
            }
            return document;
        }

        public static Document FindReplaceRegionTest(this Document document, string searchRegion)
        {
            int countText = document.Descendants<Text>().Count();
            if (countText > 0)
            {
                string openOldTag = "<r-" + searchRegion + ">";
                string openTag = "<r-" + searchRegion + " replace=\"true\">";
                string closeOldTag = "</r-" + searchRegion + ">";
                string closeTag = "</r-" + searchRegion + " replace=\"true\">";

                foreach (Text text in document.Descendants<Text>())
                {
                    if (text.Text.Contains(openOldTag))
                    {
                        text.Text = text.Text.Replace(openTag, "");
                    }

                    if (text.Text.Contains(closeOldTag))
                    {
                        text.Text = text.Text.Replace(closeOldTag, "");
                    }

                    if (text.Text.Contains(openTag))
                    {
                        text.Text = text.Text.Replace(openTag, "");
                    }

                    if (text.Text.Contains(closeTag))
                    {
                        text.Text = text.Text.Replace(closeTag, "");
                    }
                }
            }
            return document;
        }


        public static List<ParagraphCustom> FindReplaceRegionContentTest(List<ParagraphCustom> paramParagraphsCustom, string oldValue, string newValue, string searchRegion)
        {
            List<ParagraphCustom> collecion = new List<ParagraphCustom>();

            foreach (var item in paramParagraphsCustom)
            {
                ParagraphCustom paragraphCustom = new ParagraphCustom();
                paragraphCustom.Clone(item.Key, item.Paragraph);
                collecion.Add(paragraphCustom);
            }

            int countCollection = collecion.Count();
            if (countCollection > 0)
            {
                string openOldTag = "<r-" + searchRegion + ">";
                string openTag = "<r-" + searchRegion + " replace=\"true\">";
                string closeOldTag = "</r-" + searchRegion + ">";
                string closeTag = "</r-" + searchRegion + " replace=\"true\">";

                int i = 0;
                foreach (ParagraphCustom paragraphCustom in collecion)
                {
                    foreach (var text in paragraphCustom.Paragraph.Descendants<Text>())
                    {
                        if (text.Text.Contains(openOldTag))
                        {
                            text.Text = text.Text.Replace(openOldTag, openTag);
                        }

                        if (text.Text.Contains(closeOldTag))
                        {
                            text.Text = text.Text.Replace(closeOldTag, closeTag);
                        }

                        if (text.Text.Contains(oldValue))
                        {
                            text.Text = text.Text.Replace(oldValue, newValue);
                        }
                    }
                    i++;
                }
            }
            return collecion;
        }

        public static List<ParagraphCustom> CloneParagraph(this Document document)
        {
            List<ParagraphCustom> collecion = new List<ParagraphCustom>();

            var paragraphs = document.Descendants<Paragraph>().ToList();

            int i = 0;
            foreach (var item in paragraphs)
            {
                ParagraphCustom paragraphCustom = new ParagraphCustom();
                paragraphCustom.Clone(i, item);
                collecion.Add(paragraphCustom);
                i++;
            }

            return collecion;
        }

        public static List<ParagraphCustom> AddChildParagraphCustom(this List<ParagraphCustom> paramParagraphsCustom, List<ParagraphCustom> paragraphsNovos, int? startPosition)
        {
            List<ParagraphCustom> collecion = new List<ParagraphCustom>();

            if (startPosition != null)
            {
                var countParamParagraphsCustom = paramParagraphsCustom.Count();
                var countParagraphsNovos = paragraphsNovos.Count();

                //int total = countParamParagraphsCustom + countParagraphsNovos;

                int g = 0;
                for (int i = 0; i < countParamParagraphsCustom; i++)
                {
                    //var item = paramParagraphsCustom[i];
                    //ParagraphCustom paragraphCustom = new ParagraphCustom();
                    //paragraphCustom.Clone(g, item.Paragraph);
                    //collecion.Add(paragraphCustom);

                    if (i == startPosition)
                    {
                        for (int j = 0; j < countParagraphsNovos; j++)
                        {
                            g++;
                            var item_j = paragraphsNovos[j];
                            ParagraphCustom paragraphCustom_j = new ParagraphCustom();
                            paragraphCustom_j.Clone(g, item_j.Paragraph);
                            collecion.Add(paragraphCustom_j);
                        }
                    }

                    var item = paramParagraphsCustom[i];
                    ParagraphCustom paragraphCustom = new ParagraphCustom();
                    paragraphCustom.Clone(g, item.Paragraph);
                    collecion.Add(paragraphCustom);

                    g++;
                }

                //foreach (var item in paramParagraphsCustom)
                //{
                //    ParagraphCustom paragraphCustom = new ParagraphCustom();
                //    paragraphCustom.Clone(item.Key, item.Paragraph);
                //    collecion.Add(paragraphCustom);
                //}
                return collecion;
            }
            else
            {
                return paramParagraphsCustom;
            }
        }


        public static Document RemoveParagraphs(this Document document)
        {
            var paragraphs = document.Descendants<Paragraph>().ToList();

            foreach (var item in paragraphs)
            {
                item.Remove();
            }

            return document;
        }

        public static Document AddParagraphs(this Document document, List<ParagraphCustom> paramParagraphsCustom)
        {
            foreach (var item in paramParagraphsCustom)
            {
                ParagraphCustom paragraphCustom = new ParagraphCustom();
                paragraphCustom.Clone(item.Key, item.Paragraph);
                document.Body.AppendChild(paragraphCustom.Paragraph);
            }

            return document;
        }


        public static string AllText(this Document document)
        {
            string result = "";
            foreach (Text text in document.Descendants<Text>())
            {
                result += " \r\n " + text.InnerText;
            }
            return result;
        }
    }
}
