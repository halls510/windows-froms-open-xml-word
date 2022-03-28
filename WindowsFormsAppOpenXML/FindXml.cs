using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsAppOpenXML
{
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
                    collection.Set(i,text);
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
