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

            var docs = document.Descendants<Paragraph>();
            var t = docs.Count();
            //for (int k = 0; k <= docs.Count(); k++)
            //{
            //    var c = removes.Contains(k);
            //    if (c == true)
            //    {
            //        docs.ElementAt<Paragraph>(k).Remove();
            //    }
            //}

            for (int l = docs.Count() - 1; l >= 0; l--)
            {
                var c = removes.Contains(l);
                if (c == true)
                {
                    docs.ElementAt<Paragraph>(l).Remove();
                }
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


        public static int? FindPositionRegionReplaceTrue(this Document document, string search)
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
                    result = i + 1;
                    countFinish = 0;
                }

                i++;
            }
            return result;
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

        public static int? FindPositionRegion(this List<ParagraphCustom> paragraphCustoms, string search)
        {
            int? result = null;

            string openTag = "<r-" + search;

            int i = 0;

            var ps = paragraphCustoms.Select(s => s.Paragraph).ToList();

            foreach (Paragraph paragraph in ps)
            {
                var texts = paragraph.Descendants<Text>();
                foreach (var text in texts)
                {
                    if (text.Text.Contains(openTag))
                    {
                        result = i;
                        break;
                    }
                }

                i++;
            }
            return result;
        }
        public static int? FindPositionRegion(this Document document, string search)
        {
            int? result = null;

            string openTag = "<r-" + search;

            int i = 0;

            foreach (Paragraph paragraph in document.Descendants<Paragraph>())
            {
                var texts = paragraph.Descendants<Text>();
                foreach (var text in texts)
                {
                    if (text.Text.Contains(openTag))
                    {
                        result = i;
                    }
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

        public static List<ParagraphCustom> AddChildParagraphCustom(this List<ParagraphCustom> paramParagraphsCustom, List<ParagraphCustom> paragraphsNovos, int? startPosition, string tipoParagraph, RegionCustom regionCustom)
        {
            List<ParagraphCustom> collecion = new List<ParagraphCustom>();

            if (startPosition != null)
            {
                int ParseStartPosition = (int)startPosition;
                var countParamParagraphsCustom = paramParagraphsCustom.Count();
                var countParagraphsNovos = paragraphsNovos.Count();

                //bool check_paragraph_unico = (countParagraphsNovos == 1) ? true : false;
                //bool check_paragraph_unico = (countParagraphsNovos == 1) ? true : false;
                bool check_paragraph_unico = false;

                check_paragraph_unico = regionCustom.MesmoParagrafo;

                if (regionCustom.MesmoParagrafo == true && countParagraphsNovos > 1)
                {
                    int? position = paramParagraphsCustom.FindPositionRegion(regionCustom.Region);
                    if (position != null)
                    {
                        ParseStartPosition =(int)position;
                    }
                }

                int g = 0;
                for (int i = 0; i <= countParamParagraphsCustom; i++)
                {
                    if (i == ParseStartPosition)
                    {
                        var positionParagraphCustomBefore = (ParseStartPosition - 1);
                        var ParagraphsCustomBefore = paramParagraphsCustom.ElementAtOrDefault(positionParagraphCustomBefore);

                        for (int j = 0; j < countParagraphsNovos; j++)
                        {
                            var t_paragraphsNovos = paragraphsNovos.ElementAtOrDefault<ParagraphCustom>(j);
                            if (t_paragraphsNovos != null)
                            {
                                var item_j = t_paragraphsNovos;
                                ParagraphCustom paragraphCustom_j = new ParagraphCustom();
                                paragraphCustom_j.Clone(g, item_j.Paragraph);
                                var runs_item_j = paragraphCustom_j.Paragraph.Descendants<Run>().ToList();
                                if (check_paragraph_unico == true)
                                {
                                    switch (tipoParagraph)
                                    {
                                        case "full":
                                            // add o paragrafo
                                            collecion.Add(paragraphCustom_j);
                                            g++;
                                            break;
                                        case "before":
                                            // add o paragrafo
                                            if (countParagraphsNovos > 1)
                                            {
                                                var ParagraphsCustomParseStartPosition = collecion.ElementAtOrDefault(ParseStartPosition);

                                                if (ParagraphsCustomParseStartPosition != null)
                                                {
                                                    ParagraphCustom pcj = new ParagraphCustom();
                                                    pcj.Clone(ParagraphsCustomParseStartPosition.Key, ParagraphsCustomParseStartPosition.Paragraph);
                                                    foreach (var item_run in runs_item_j)
                                                    {
                                                        var item_run_clone = (Run)item_run.Clone();
                                                        pcj.Paragraph.AppendChild<Run>(item_run_clone);
                                                    }
                                                    var el = collecion.ElementAtOrDefault<ParagraphCustom>(ParseStartPosition);
                                                    if (el != null)
                                                    {
                                                        el.Key = pcj.Key;
                                                        el.Paragraph = pcj.Paragraph;
                                                    }
                                                }
                                                else
                                                {
                                                    collecion.Add(paragraphCustom_j);
                                                    g++;
                                                }
                                            }
                                            else
                                            {
                                                collecion.Add(paragraphCustom_j);
                                                g++;
                                            }
                                            break;
                                        case "middle":
                                            // adiciona o run com text
                                            if (countParagraphsNovos > 1)
                                            {
                                                //var ParagraphsCustomParseStartPosition = collecion.ElementAtOrDefault(ParseStartPosition);
                                                var ParagraphsCustomParseStartPosition = paramParagraphsCustom.ElementAtOrDefault(ParseStartPosition);

                                                if (ParagraphsCustomParseStartPosition != null)
                                                {
                                                    ParagraphCustom pcj = new ParagraphCustom();
                                                    pcj.Clone(ParagraphsCustomParseStartPosition.Key, ParagraphsCustomParseStartPosition.Paragraph);
                                                    foreach (var item_run in runs_item_j)
                                                    {
                                                        var item_run_clone = (Run)item_run.Clone();
                                                        pcj.Paragraph.AppendChild<Run>(item_run_clone);
                                                    }
                                                    var el = collecion.ElementAtOrDefault<ParagraphCustom>(ParseStartPosition);
                                                    if (el != null)
                                                    {

                                                        el.Key = pcj.Key;
                                                        el.Paragraph = pcj.Paragraph;
                                                    }
                                                    else
                                                    {
                                                        collecion.Add(pcj);
                                                        g++;
                                                    }
                                                }
                                                else
                                                {
                                                    collecion.Add(paragraphCustom_j);
                                                    g++;
                                                }
                                            }
                                            else
                                            {
                                                if (ParagraphsCustomBefore != null)
                                                {
                                                    ParagraphCustom pcj = new ParagraphCustom();
                                                    pcj.Clone(ParagraphsCustomBefore.Key, ParagraphsCustomBefore.Paragraph);
                                                    foreach (var item_run in runs_item_j)
                                                    {
                                                        var item_run_clone = (Run)item_run.Clone();
                                                        pcj.Paragraph.AppendChild<Run>(item_run_clone);
                                                    }
                                                    var el = collecion.ElementAtOrDefault<ParagraphCustom>(positionParagraphCustomBefore);
                                                    if (el != null)
                                                    {

                                                        el.Key = pcj.Key;
                                                        el.Paragraph = pcj.Paragraph;
                                                    }
                                                    g++;
                                                }
                                                else
                                                {
                                                    collecion.Add(paragraphCustom_j);
                                                    g++;
                                                }
                                            }
                                            break;
                                        case "after":
                                            // adiciona o run com text
                                            if (countParagraphsNovos > 1)
                                            {
                                                //var ParagraphsCustomParseStartPosition = collecion.ElementAtOrDefault(ParseStartPosition);
                                                var ParagraphsCustomParseStartPosition = paramParagraphsCustom.ElementAtOrDefault(ParseStartPosition);

                                                if (ParagraphsCustomParseStartPosition != null)
                                                {
                                                    ParagraphCustom pcj = new ParagraphCustom();
                                                    pcj.Clone(ParagraphsCustomParseStartPosition.Key, ParagraphsCustomParseStartPosition.Paragraph);
                                                    foreach (var item_run in runs_item_j)
                                                    {
                                                        var item_run_clone = (Run)item_run.Clone();
                                                        pcj.Paragraph.AppendChild<Run>(item_run_clone);
                                                    }
                                                    var el = collecion.ElementAtOrDefault<ParagraphCustom>(ParseStartPosition);
                                                    if (el != null)
                                                    {

                                                        el.Key = pcj.Key;
                                                        el.Paragraph = pcj.Paragraph;
                                                    }
                                                    else
                                                    {
                                                        collecion.Add(pcj);
                                                        g++;
                                                    }
                                                }
                                                else
                                                {
                                                    collecion.Add(paragraphCustom_j);
                                                    g++;
                                                }
                                            }
                                            else
                                            {
                                                if (ParagraphsCustomBefore != null)
                                                {
                                                    ParagraphCustom pcj = new ParagraphCustom();
                                                    pcj.Clone(ParagraphsCustomBefore.Key, ParagraphsCustomBefore.Paragraph);
                                                    foreach (var item_run in runs_item_j)
                                                    {
                                                        var item_run_clone = (Run)item_run.Clone();
                                                        pcj.Paragraph.AppendChild<Run>(item_run_clone);
                                                    }
                                                    var el = collecion.ElementAtOrDefault<ParagraphCustom>(positionParagraphCustomBefore);
                                                    if (el != null)
                                                    {

                                                        el.Key = pcj.Key;
                                                        el.Paragraph = pcj.Paragraph;
                                                    }

                                                    g++;
                                                }
                                                else
                                                {
                                                    collecion.Add(paragraphCustom_j);
                                                    g++;
                                                }
                                            }
                                            break;
                                    }
                                }
                                else
                                {
                                    collecion.Add(paragraphCustom_j);
                                    g++;
                                }
                            }
                        }
                    }

                    var t_paramParagraphsCustom = paramParagraphsCustom.ElementAtOrDefault<ParagraphCustom>(i);
                    if (t_paramParagraphsCustom != null)
                    {
                        var item = t_paramParagraphsCustom;
                        ParagraphCustom paragraphCustom = new ParagraphCustom();
                        paragraphCustom.Clone(g, item.Paragraph);
                        collecion.Add(paragraphCustom);
                    }
                    g++;
                }

                return collecion;
            }
            else
            {
                return paramParagraphsCustom;
            }
        }

        //public static List<ParagraphCustom> AddChildParagraphCustom(this List<ParagraphCustom> paramParagraphsCustom, List<ParagraphCustom> paragraphsNovos, int? startPosition)
        //{
        //    List<ParagraphCustom> collecion = new List<ParagraphCustom>();

        //    if (startPosition != null)
        //    {
        //        var countParamParagraphsCustom = paramParagraphsCustom.Count();
        //        var countParagraphsNovos = paragraphsNovos.Count();

        //        int g = 0;
        //        for (int i = 0; i <= countParamParagraphsCustom; i++)
        //        {                  

        //            if (i == startPosition)
        //            {
        //                for (int j = 0; j < countParagraphsNovos; j++)
        //                {
        //                    var t_paragraphsNovos = paragraphsNovos.ElementAtOrDefault<ParagraphCustom>(j);
        //                    if (t_paragraphsNovos != null)
        //                    {
        //                        var item_j = t_paragraphsNovos;
        //                        ParagraphCustom paragraphCustom_j = new ParagraphCustom();
        //                        paragraphCustom_j.Clone(g, item_j.Paragraph);
        //                        collecion.Add(paragraphCustom_j);
        //                        g++;
        //                    }
        //                }
        //            }

        //            var t_paramParagraphsCustom = paramParagraphsCustom.ElementAtOrDefault<ParagraphCustom>(i);
        //            if (t_paramParagraphsCustom != null)
        //            {
        //                var item = t_paramParagraphsCustom;
        //                ParagraphCustom paragraphCustom = new ParagraphCustom();
        //                paragraphCustom.Clone(g, item.Paragraph);
        //                collecion.Add(paragraphCustom);
        //            }
        //            g++;                   
        //        }

        //        return collecion;
        //    }
        //    else
        //    {
        //        return paramParagraphsCustom;
        //    }
        //}


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
