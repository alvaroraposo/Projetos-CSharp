using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace ChecklistReportLegislativo
{
    public partial class ThisDocument
    {
        public string checklistXMLPartID = string.Empty;
        private Office.CustomXMLPart checklistXMLPart;
        private XmlDocument xd;
        private XmlNamespaceManager xnm;
        private const string prefix = "xmlns:ns='http://schemas.microsoft.com/vsto/samples'";
        private static string NODEPATH = "ns:checklists/ns:checklist/ns:pa";
        private static string ISSELECTEDPATH = "/ns:isSelected";
        private static string ORGAOPATH = "ns:checklists/ns:checklist/ns:orgao";
        private static string NOMERESPONSAVELPATH = "ns:checklists/ns:checklist/ns:nomeResponsavel";
        private static string CARGORESPONSAVELPATH = "ns:checklists/ns:checklist/ns:cargoResponsavel";
        private static string DATAAVALIACAOPATH = "ns:checklists/ns:checklist/ns:dataAvaliacao";
        private static string XMLPATH = AppDomain.CurrentDomain.BaseDirectory + "Checklist.xml";
        private int countItens = 0;

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            string xmlData = GetXmlFromResource();
            this.Application.ActiveWindow.ActivePane.View.Zoom.Percentage = 100;

            if (xmlData != null)
            {
                countItens = 0;

                for (int i = 0; i < this.Controls.Count; i++)
                {
                    if (this.Controls[i].GetType() != checkBox01.GetType())
                        continue;

                    (this.Controls[i] as CheckBox).CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);
                    countItens++;
                }

                if (xmlData == null || xmlData == "")
                {
                    MessageBox.Show("Arquivo Checklist.xml não encontrado", "Aviso");
                    return;
                }

                AddCustomXmlPart(xmlData);
                BindControlsToCustomXmlPart();
            }
        }

        private string GetXmlFromResource()
        {
            string filename = XMLPATH;

            XmlDocument xd = new XmlDocument();
            FileStream stream1 = new FileStream(filename, FileMode.OpenOrCreate);

            using (System.IO.StreamReader resourceReader = new System.IO.StreamReader(stream1))
            {
                if (resourceReader != null)
                {
                    return resourceReader.ReadToEnd();
                }
            }

            return null;
        }

        private void SaveDocumentIntoXml()
        {
            string filename = XMLPATH;

            FileStream stream = new FileStream(filename, FileMode.Create);

            XmlNode cn = null;
            cn = xd.SelectSingleNode(ORGAOPATH, xnm);
            cn.InnerText = comboBoxContentControl1.Range.Text.Trim('\r', '\a');

            cn = xd.SelectSingleNode(NOMERESPONSAVELPATH, xnm);
            cn.InnerText = this.Tables[1].Rows[3].Cells[2].Range.Text.Trim('\r', '\a');

            cn = xd.SelectSingleNode(CARGORESPONSAVELPATH, xnm);
            cn.InnerText = this.Tables[1].Rows[4].Cells[2].Range.Text.Trim('\r', '\a');

            cn = xd.SelectSingleNode(DATAAVALIACAOPATH, xnm);
            cn.InnerText = datePickerContentControl1.Text.Trim('\r', '\a');

            using (XmlTextWriter wr = new XmlTextWriter(stream, Encoding.UTF8))
            {
                wr.Formatting = Formatting.Indented;
                xd.Save(wr);
            }

            MessageBox.Show("Checklist.xml salvo com sucesso.", "Aviso");
        }

        private void AddCustomXmlPart(string xmlData)
        {
            if (xmlData != null && xmlData != "")
            {
                checklistXMLPart = this.CustomXMLParts.SelectByID(checklistXMLPartID);
                checklistXMLPart = this.CustomXMLParts.Add(xmlData);
                checklistXMLPart.NamespaceManager.AddNamespace("ns", "http://schemas.microsoft.com/vsto/samples");
                checklistXMLPartID = checklistXMLPart.Id;
            }
        }

        private void BindControlsToCustomXmlPart()
        {
            int nodeCount = 0;

            xd = new XmlDocument();
            string str = checklistXMLPart.XML;
            xd.LoadXml(str);

            xnm = new XmlNamespaceManager(xd.NameTable);
            xnm.AddNamespace("ns", "http://schemas.microsoft.com/vsto/samples");

            XmlNode cn = null;
            cn = xd.SelectSingleNode(ORGAOPATH, xnm);
            for (int i = 0; i < comboBoxContentControl1.DropDownListEntries.Count; i++)
            {
                if (cn.InnerText == comboBoxContentControl1.DropDownListEntries[i + 1].Text)
                {
                    comboBoxContentControl1.DropDownListEntries[i + 1].Select();
                    break;
                }
            }

            cn = xd.SelectSingleNode(NOMERESPONSAVELPATH, xnm);
            this.Tables[1].Rows[3].Cells[2].Range.Text = cn.InnerText.Trim('\r', '\a');

            cn = xd.SelectSingleNode(CARGORESPONSAVELPATH, xnm);
            this.Tables[1].Rows[4].Cells[2].Range.Text = cn.InnerText.Trim('\r', '\a');

            cn = xd.SelectSingleNode(DATAAVALIACAOPATH, xnm);
            datePickerContentControl1.Text = cn.InnerText.Trim('\r', '\a');

            for (int i = 0; i < this.Controls.Count; i++)
            {
                if (this.Controls[i].GetType() != checkBox01.GetType())
                    continue;

                CheckBox cb = this.Controls[i] as CheckBox;
                string paNum = cb.Name.Substring(8);

                string nodep = NODEPATH + paNum + ISSELECTEDPATH;
                XmlNode nd = xd.SelectSingleNode(nodep, xnm);
                cb.Checked = Boolean.Parse(nd.InnerText);
            }
        }

        private void checkBox_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb = sender as CheckBox;
            string str = cb.Name.Substring(8);

            string nodep = NODEPATH + str + ISSELECTEDPATH;
            XmlNode nd = xd.SelectSingleNode(nodep, xnm);
            string inner = cb.Checked.ToString().ToLower();
            nd.InnerText = inner;
        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código gerado pelo Designer VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.btnGerarXML.Click += new System.EventHandler(this.btnGerarXML_Click);
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);

        }

        #endregion

        private void btnGerarXML_Click(object sender, EventArgs e)
        {
            SaveDocumentIntoXml();
        }
    }
}
