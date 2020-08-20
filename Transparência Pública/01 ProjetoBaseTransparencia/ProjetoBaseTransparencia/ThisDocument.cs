using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace ProjetoBaseTransparencia
{
    public partial class ThisDocument
    {
        public string checklistXMLPartID = string.Empty;
        private Office.CustomXMLPart checklistXMLPart;
        private XmlNamespaceManager xnm;
        private const string prefix = "xmlns:ns='http://schemas.microsoft.com/vsto/samples'";
        private static string NODEPATH = "ns:checklists/ns:checklist/ns:pa";
        private static float ItemWIDTH = 382.5F;
        private static float StatusWIDTH = 50;
        private static WdColor TABLECOLOR = (WdColor) 75 + 0x100 * 172 + 0x10000 * 198;
        private int countItens = 0;
        Table tableCapitulo = null;
        Table tablePA = null;
        Table tableItem = null;
        Table tableCriterios = null;
        private Table tableStyle;
        private Paragraph capituloStyle;
        Microsoft.Office.Interop.Word.Document docBase;
        private ContentControlListEntries listOrgaos;


        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            bool isWordFileOpen = GetDocbaseTables();
            bool isOrgaosFileOpen = GetDocbaseOrgaos();

            if (isWordFileOpen && isOrgaosFileOpen)
            {
                ParseTablesToChecklist();
                ParseXSD();
                ParseChecklistXML();
                ParseQuestoesLaudoXML();
            }
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
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);

        }

        #endregion


        private void ParseQuestoesLaudoXML()
        {
            string fileXML = AppDomain.CurrentDomain.BaseDirectory + "QuestoesLaudo.xml";
            FileStream fsXML = new FileStream(fileXML, FileMode.Create);
            StreamWriter stream = new StreamWriter(fsXML);

            stream.WriteLine("<?xml version='1.0' encoding='utf-8' ?>\n");
            stream.WriteLine("<analise xmlns='http://schemas.microsoft.com/vsto/samples'>\n");
            stream.WriteLine("\t<laudo>");
            stream.WriteLine("\t\t<orgao></orgao>");
            stream.WriteLine("\t\t<nomeResponsavel></nomeResponsavel>");
            stream.WriteLine("\t\t<cargoResponsavel></cargoResponsavel>");
            stream.WriteLine("\t\t<dataAvaliacao></dataAvaliacao>");

            foreach (Row itemRow in tableItem.Rows)
            {
                if (itemRow.IsFirst)
                    continue;

                string numItem = itemRow.Cells[1].Range.Text.Trim('\r', '\a');

                string pa = "pa";
                pa += numItem;

                stream.WriteLine("\t\t<" + pa + ">");
                stream.WriteLine("\t\t\t<isSelected>false</isSelected>");
                stream.WriteLine("\t\t</" + pa + ">");
            }

            stream.WriteLine("\t</laudo>");
            stream.WriteLine("</analise>");
            stream.Close();
        }

        private void ParseChecklistXML()
        {
            string fileXML = AppDomain.CurrentDomain.BaseDirectory + "Checklist.xml";            

            FileStream fsXML = new FileStream(fileXML, FileMode.Create);
            StreamWriter stream = new StreamWriter(fsXML);
            
            stream.WriteLine("<?xml version='1.0' encoding='utf-8' ?>\n");
            stream.WriteLine("<checklists xmlns='http://schemas.microsoft.com/vsto/samples'>\n");
            stream.WriteLine("\t<checklist>");
            stream.WriteLine("\t\t<orgao></orgao>");
            stream.WriteLine("\t\t<nomeResponsavel></nomeResponsavel>");
            stream.WriteLine("\t\t<cargoResponsavel></cargoResponsavel>");
            stream.WriteLine("\t\t<dataAvaliacao></dataAvaliacao>");

            foreach (Row itemRow in tableItem.Rows)
            {
                if (itemRow.IsFirst)
                    continue;

                string numItem = itemRow.Cells[1].Range.Text.Trim('\r', '\a');

                string pa = "pa";
                pa += numItem;

                if(numItem == "96")
                {
                    string txt = "96";
                }
                if (numItem == "97")
                {
                    string txt = "97";
                }
                stream.WriteLine("\t\t<" + pa + ">");
                stream.WriteLine("\t\t\t<isSelected>false</isSelected>");

                string orientacao = "<orientacao>";
                orientacao += itemRow.Cells[2].Range.Text.Trim('\r', '\a');
                orientacao += "</orientacao>";                

                string situacaoEncontrada = "<situacaoEncontrada>";
                situacaoEncontrada += itemRow.Cells[3].Range.Text.Trim('\r', '\a');
                situacaoEncontrada += "</situacaoEncontrada>";                

                string descricaoAchado = "<descricaoAchado>";
                descricaoAchado += itemRow.Cells[4].Range.Text.Trim('\r', '\a');
                descricaoAchado += "</descricaoAchado>";

                stream.WriteLine("\t\t\t" + orientacao);
                stream.WriteLine("\t\t\t" + situacaoEncontrada);
                stream.WriteLine("\t\t\t" + descricaoAchado);
                stream.WriteLine("\t\t\t<criterios>");

                string tipo = null;
                string obrigatoriedade = null;
                string causa = null;
                string efeito = null;
                string qa = null;
                string peso = null;

                int entrouNoPA = 0;
                foreach (Row paRow in tablePA.Rows)
                {
                    if (paRow.IsFirst)
                        continue;

                    string paItem = paRow.Cells[5].Range.Text.Trim('\r', '\a');
                    string[] criteriosItem = paItem.Split('.');

                    int intNumItem = Int32.Parse(numItem);
                    int paRowIndex = paRow.Index;
                    foreach (string item in criteriosItem)
                    {
                        int intItem = Int32.Parse(item);

                        if (intNumItem != intItem)
                            continue;

                        entrouNoPA++;
                        string strEntrou = numItem + entrouNoPA;
                        
                        List<string> criteriosFundamentacao = new List<string>();
                        foreach (string criterio in criteriosItem)
                        {
                            string strNovoItem = criterio.Trim('\r', '\a');
                            int novoItem = int.Parse(strNovoItem);
                            string colunaFundamentacao = tableItem.Rows[novoItem + 1].Cells[6].Range.Text.Trim('\r', '\a');
                            string[] fundamentacaoArray = colunaFundamentacao.Split('.');

                            bool isFundamentacaoNaLista = false;
                            foreach (string array in fundamentacaoArray)
                            {
                                foreach (string lista in criteriosFundamentacao)
                                {
                                    if (array == lista)
                                    {
                                        isFundamentacaoNaLista = true;
                                        break;
                                    }
                                }

                                if (!isFundamentacaoNaLista)
                                    criteriosFundamentacao.Add(array);
                            }
                        }

                        foreach (string criterio in criteriosFundamentacao)
                        {
                            int intCriterio = Int32.Parse(criterio);
                            Row rowCriterio = tableCriterios.Rows[intCriterio + 1];

                            string artigo = rowCriterio.Cells[2].Range.Text.Trim('\r', '\a');
                            string descricao = rowCriterio.Cells[3].Range.Text.Trim('\r', '\a');
                            stream.WriteLine("\t\t\t\t<criterio>");
                            stream.WriteLine("\t\t\t\t\t<artigo>");
                            stream.WriteLine("\t\t\t\t\t" + artigo);
                            stream.WriteLine("\t\t\t\t\t</artigo>");
                            stream.WriteLine("\t\t\t\t\t<descricao>");
                            stream.WriteLine("\t\t\t\t\t" + descricao);
                            stream.WriteLine("\t\t\t\t\t</descricao>");
                            stream.WriteLine("\t\t\t\t</criterio>");
                        }
                    }

                    if (entrouNoPA == 0)
                        continue;

                    tipo = "\t\t\t<tipo>";
                    tipo += paRow.Cells[4].Range.Text.Trim('\r', '\a');
                    tipo += "</tipo>";

                    obrigatoriedade = "\t\t\t<obrigatoriedade>";
                    obrigatoriedade += itemRow.Cells[5].Range.Text.Trim('\r', '\a');
                    obrigatoriedade += "</obrigatoriedade>";

                    causa = "\t\t\t<causa>";
                    causa += paRow.Cells[6].Range.Text.Trim('\r', '\a');
                    causa += "</causa>";

                    efeito = "\t\t\t<efeito>";
                    efeito += paRow.Cells[7].Range.Text.Trim('\r', '\a');
                    efeito += "</efeito>";

                    qa = "\t\t\t<qa>";
                    qa += paRow.Cells[8].Range.Text.Trim('\r', '\a');
                    qa += "</qa>";

                    peso = "\t\t\t<peso>";
                    peso += itemRow.Cells[7].Range.Text.Trim('\r', '\a');
                    peso += "</peso>";

                    break;
                }
                stream.WriteLine("\t\t\t</criterios>");

                stream.WriteLine(tipo);
                stream.WriteLine(obrigatoriedade);
                stream.WriteLine(causa);
                stream.WriteLine(efeito);
                stream.WriteLine(qa);
                stream.WriteLine(peso);
                stream.WriteLine("\t\t</" + pa + ">");
            }

            stream.WriteLine("\t</checklist>");
            stream.WriteLine("</checklists>");
            stream.Close();
        }

        private void ParseXSD()
        {
            string fileBegin = AppDomain.CurrentDomain.BaseDirectory + "xsdbegin.txt";
            string fileEnd = AppDomain.CurrentDomain.BaseDirectory + "xsdend.txt";
            string fileXSD = AppDomain.CurrentDomain.BaseDirectory + "Checklist.xsd";
            string middle = "";

            FileStream fsBegin = new FileStream(fileBegin, FileMode.Open);
            StreamReader strBegin = new StreamReader(fsBegin);

            FileStream fsEnd = new FileStream(fileEnd, FileMode.Open);
            StreamReader strEnd = new StreamReader(fsEnd);

            FileStream fsXSD = new FileStream(fileXSD, FileMode.Create);
            StreamWriter stream = new StreamWriter(fsXSD);

            string write = strBegin.ReadLine();
            while (write != null)
            {
                stream.WriteLine(write);
                write = strBegin.ReadLine();
            }

            foreach (Row itemRow in tableItem.Rows)
            {
                if (itemRow.IsFirst)
                    continue;

                middle = "\t<xs:element name='pa";
                middle += itemRow.Cells[1].Range.Text.Trim('\r', '\a') + "' type='PontoAuditoriaType'" + "/>";

                stream.WriteLine(middle);                
            }          

            write = strEnd.ReadLine();
            while (write != null)
            {
                stream.WriteLine(write);
                write = strEnd.ReadLine();
            }

            stream.Close();
        }

        private void ParseTablesToChecklist()
        {
            foreach (Row rowCapitulo in tableCapitulo.Rows)
            {
                if (rowCapitulo.IsFirst)
                    continue;

                ParseTituloCapitulo(rowCapitulo);
                ParseTable(rowCapitulo);
                
                if (!rowCapitulo.IsLast)
                    Words.Last.InsertBreak(WdBreakType.wdPageBreak);
            }
        }

        private void ParseTable(Row rowCapitulo)
        {
            foreach (Row rowPA in tablePA.Rows)
            {
                if (rowPA.IsFirst)
                    continue;

                if (rowCapitulo.Cells[1].Range.Text != rowPA.Cells[2].Range.Text)
                    continue;

                object miss = System.Reflection.Missing.Value;

                Paragraph par = Content.Paragraphs.Add(miss);
                par.Range.InsertParagraphAfter();

//Separa Itens Auditoria e Fundamentação
                string[] criterios = rowPA.Cells[5].Range.Text.Split('.');

                if (criterios == null || criterios.Length <= 0)
                    continue;

               List<string> fundamentacao = new List<string>();

                foreach(string criterio in criterios)
                {
                    string strNovoItem = criterio.Trim('\r', '\a');
                    int novoItem = int.Parse(strNovoItem);
                    string colunaFundamentacao = tableItem.Rows[novoItem + 1].Cells[6].Range.Text.Trim('\r', '\a');
                    string[] fundamentacaoArray = colunaFundamentacao.Split('.');

                    bool isFundamentacaoNaLista = false;
                    foreach (string array in fundamentacaoArray)
                    {
                        foreach (string lista in fundamentacao)
                        {
                            if (array == lista)
                            {
                                isFundamentacaoNaLista = true;
                                break;
                            }                            
                        }

                        if (!isFundamentacaoNaLista)
                            fundamentacao.Add(array);
                    }
                }

                if (fundamentacao == null || fundamentacao.Count <= 0)
                    continue;
//Título da Tabela
                Table t = par.Application.ActiveDocument.Tables.Add(par.Range, criterios.Length + fundamentacao.Count + 3, 2);
                t.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
                t.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                t.Borders.OutsideColor = TABLECOLOR;
                t.Borders.InsideColor = TABLECOLOR;                
                t.Rows[1].Cells[1].Range.Text = rowPA.Cells[3].Range.Text.Trim('\r', '\a');
                t.Rows[1].Cells.Merge();
                t.Rows[1].Cells[1].Range.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);

                t.Rows[1].Cells[1].Range.Shading.BackgroundPatternColor = TABLECOLOR;
                t.Rows[1].Cells[1].Range.Font.Color = WdColor.wdColorWhite;

//Orientação x Dados Obrigatórios e Status
                t.Rows[2].Cells[1].Range.Text = rowPA.Cells[4].Range.Text.Trim('\r', '\a');
                t.Rows[2].Cells[1].Width = ItemWIDTH;
                t.Rows[2].Cells[1].Range.Font.Bold = -1;
                t.Rows[2].Cells[1].Range.Font.Italic = -1;
                t.Rows[2].Cells[2].Width = StatusWIDTH;                
                t.Rows[2].Cells[2].Range.Text = "Achado";
                t.Rows[2].Cells[2].Range.Font.Bold = -1;
                t.Rows[2].Cells[2].Range.Font.Italic = -1;
                t.Rows[2].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
//Itens de Auditoria
                int novoRow;             
                for(novoRow = 0; novoRow < criterios.Length; novoRow++)
                {
                    string strNovoItem = criterios[novoRow].Trim('\r','\a');
                    int novoItem = int.Parse(strNovoItem);

                    t.Rows[novoRow + 3].Cells[1].Range.Text = tableItem.Rows[novoItem + 1].Cells[2].Range.Text.Trim('\r', '\a');
                    
//                    object obj = (object)t.Rows[novoRow + 3].Cells[2].Range;
//                    Word.ContentControl cb = t.Rows[novoRow + 3].Cells[2].Range.ContentControls.Add(WdContentControlType.wdContentControlCheckBox, obj);
                    countItens++;

                    t.Rows[novoRow + 3].Cells[1].Width = ItemWIDTH;
                    t.Rows[novoRow + 3].Cells[2].Width = StatusWIDTH;
                    t.Rows[novoRow + 3].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
//Referência
                novoRow = novoRow + 3;
                t.Rows[novoRow].Cells[1].Range.Text = "Referência";
                t.Rows[novoRow].Cells[1].Range.Font.Bold = -1;
                t.Rows[novoRow].Cells[1].Range.Font.Italic = -1;
                t.Rows[novoRow].Cells.Merge();
//Fundamentação
                int fundRow;
                for (fundRow = 0; fundRow < fundamentacao.Count; fundRow++)
                {
                    string strNovoItem = fundamentacao[fundRow].Trim('\r', '\a');
                    int novoItem = int.Parse(strNovoItem);

                    t.Rows[novoRow + fundRow + 1].Cells[1].Range.Text = tableCriterios.Rows[novoItem + 1].Cells[2].Range.Text.Trim('\r', '\a');
                    t.Rows[novoRow + fundRow + 1].Cells.Merge();
                }
//Cumprimento
/*                novoRow = novoRow + fundRow + 1;
                t.Rows[novoRow].Cells[1].Range.Text = "Cumprimento";
                t.Rows[novoRow].Cells[1].Range.Font.Bold = -1;
                t.Rows[novoRow].Cells[1].Range.Font.Italic = -1;
                t.Rows[novoRow].Cells.Merge();

                novoRow++;                
                string strObrigatorioPA = rowPA.Cells[7].Range.Text.Trim('\r', '\a');
                if (String.Compare(strObrigatorioPA, "Sim", true) == 0)
                {
                    t.Rows[novoRow].Cells[1].Range.Text = "Obrigatório";
                }
                else
                {
                    t.Rows[novoRow].Cells[1].Range.Text = "Recomendável";
                }

                t.Rows[novoRow].Cells.Merge();*/
            }
        }

        private void ParseTituloCapitulo(Row r)
        {
            object start = 0;
            object miss = System.Reflection.Missing.Value;

            Paragraph par = Content.Paragraphs.Add(miss);            
            Content.Paragraphs.Last.set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
            Content.Paragraphs.Last.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par.Range.set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
                par.Range.Font.Color = WdColor.wdColorBlack;
            par.Range.Text = r.Cells[2].Range.Text;            
            Content.Paragraphs.Last.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
            par.Range.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
            par.Range.InsertParagraphAfter();
        }

        private bool GetDocbaseOrgaos()
        {
            string filename = AppDomain.CurrentDomain.BaseDirectory + "orgaos_auditados_TCEAM.txt";

            FileStream stream1 = new FileStream(filename, FileMode.Open);
            StreamReader strReader = new StreamReader(stream1);
            
            if (strReader == null)
                return false;

            string linha = strReader.ReadLine();
            int counter = 0;
            while (linha != null)
            {
                comboBoxContentControl1.DropDownListEntries.Add(linha, counter.ToString());
                linha = strReader.ReadLine();
                counter++;
            }

            listOrgaos = comboBoxContentControl1.DropDownListEntries;
            
            return true;
        }

        private bool GetDocbaseTables()
        {
            string filename = AppDomain.CurrentDomain.BaseDirectory + "\\BaseTransparenciaTCEAM.docx";

            FileStream stream1 = new FileStream(filename, FileMode.Open);
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            object miss = System.Reflection.Missing.Value;
            object readOnly = true;
            Microsoft.Office.Interop.Word.Document docBase = app.Documents.Open(filename, miss, readOnly, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss, miss);

            if (docBase == null)
                return false;

            tableCapitulo = docBase.Tables[1];
            tablePA = docBase.Tables[2];
            tableItem = docBase.Tables[3];
            tableCriterios = docBase.Tables[4];

            return true;
        }
    }
}
