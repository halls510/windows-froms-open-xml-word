using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Xml.Serialization;
using System.Xml;
using System.Security;

namespace WindowsFormsAppOpenXML
{

    public partial class Form1 : Form
    {
        CollectionRegionCustom _regionCustom;
        public Form1()
        {
            InitializeComponent();
            InitializeComponentCustom();
        }

        private void InitializeComponentCustom()
        {
            Rtx_Conteudo.Text = "";

            _regionCustom = new CollectionRegionCustom();
        }

        private void CloneDocument(string FilePathOrigin, string FilePathDestination)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(FilePathOrigin, false))
            {
                var newDoc = (WordprocessingDocument)wordDocument.Clone(FilePathDestination, true);
                newDoc.Save();
                newDoc.Close();
            }
        }

        private void ReplaceDocument()
        {
            string namedestiantion = Txt_FilePathDestination.Text + "\\" + Txt_NameFileDestination.Text + ".docx";
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(namedestiantion, true))
            {
                wordDocument.MainDocumentPart.Document.FindReplace("<name>", Txt_Name.Text);

                wordDocument.Save();
                wordDocument.Close();
            }
        }

        private void ReplaceDocumentRegion()
        {
            string namedestiantion = Txt_FilePathDestination.Text + "\\" + Txt_NameFileDestination.Text + ".docx";
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(namedestiantion, true))
            {
                int i = 0;
                foreach (var regiao in _regionCustom.Get())
                {

                    CollectionTextCustom texts = new CollectionTextCustom();

                    if (i == 0)
                    {
                        texts = wordDocument.MainDocumentPart.Document.FindRegion(regiao.Region);
                    }
                    if (i > 0)
                    {
                        Body body = new Body();
                        foreach (var item in texts.Get())
                        {
                            body.AppendChild(new Paragraph(new Run ( new Text (item.Text.Text))));
                        }

                        wordDocument.MainDocumentPart.Document.AppendChild(body);
                        wordDocument.Save();

                        texts = wordDocument.MainDocumentPart.Document.FindRegion(regiao.Region);
                    }

                    foreach (var pessoa in regiao.pessoas)
                    {
                        wordDocument.MainDocumentPart.Document.FindReplaceRegion(texts, "<name>", pessoa.name);
                    }

                    i++;
                }   
               

                wordDocument.Save();
                wordDocument.Close();
            }
        }

        //private List<Text> FindRegion()
        //{
        //    List<Text> collectionText = new List<Text>();
        //    string namedestiantion = Txt_FilePathDestination.Text + "\\" + Txt_NameFileDestination.Text + ".docx";
        //    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(namedestiantion, true))
        //    {

        //        collectionText = wordDocument.MainDocumentPart.Document.FindRegion(Txt_Region.Text);
        //    }

        //    return collectionText;
        //}

        private void ViewDocument(string filePath)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false))
            {
                var alltext = wordDocument.MainDocumentPart.Document.AllText();
                Rtx_Conteudo.Text = alltext;
                wordDocument.Close();
            }
        }

        private void Btn_Aplicar_Click(object sender, EventArgs e)
        {
            try
            {
                SetClone();
                ReplaceDocumentRegion();
                MessageBox.Show("Aplicado");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro");
            }
        }

        private void Btn_Visualizar_Click(object sender, EventArgs e)
        {
            if (Txt_FilePathDestination.Text != "" && Txt_NameFileDestination.Text != "")
            {
                string namedestiantion = Txt_FilePathDestination.Text + "\\" + Txt_NameFileDestination.Text + ".docx";

                ViewDocument(namedestiantion);
            }
            else if (Txt_NameFileOrigin.Text != "")
            {
                ViewDocument(Txt_NameFileOrigin.Text);
            }
        }

        private void SetClone()
        {
            Rtx_Conteudo.Text = "";
            string namedestiantion = Txt_FilePathDestination.Text + "\\" + Txt_NameFileDestination.Text + ".docx";
            CloneDocument(Txt_NameFileOrigin.Text, namedestiantion);
        }

        private void SetFileNameOrigin(string text)
        {
            Txt_NameFileOrigin.Text = text;
        }

        private void SetFileNameDestination(string text)
        {
            Txt_FilePathDestination.Text = text;
        }

        private void Btn_FileNameDestination_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    SetFileNameDestination(folderBrowserDialog1.SelectedPath);
                }
                catch (SecurityException ex)
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
            }
        }

        private void Btn_FileNameOrigin_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    SetFileNameOrigin(openFileDialog1.FileName);
                }
                catch (SecurityException ex)
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
            }
        }

        private void Btn_AddRegiao_Click(object sender, EventArgs e)
        {
            int countRegion = _regionCustom.Get().Count();
            if (_regionCustom.Get().Where(w => w.Region == Txt_Region.Text).Count() == 0)
            {
                int Codigo = countRegion + 1;
                RegionCustom RegionCustom = new RegionCustom() { Codigo = Codigo, Region = Txt_Region.Text };
                _regionCustom.Set(RegionCustom);
                Combo_Regiao.Items.Add(RegionCustom);
            }
        }

        private void Btn_AddPessoaInRegion_Click(object sender, EventArgs e)
        {
           
                RegionCustom ItemSelecionado = (RegionCustom)Combo_Regiao.Items[Combo_Regiao.SelectedIndex];

                _regionCustom.AddPessoa(ItemSelecionado.Codigo, new Pessoa { name = Txt_Name.Text });
           
        }
    }
}
