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

namespace RelatorioTransparencia
{
    public partial class ThisDocument
    {
        private static WdColor COLORQ1 = (WdColor)221 + 0x100 * 217 + 0x10000 * 195;
        private static WdColor COLORQ2 = (WdColor)214 + 0x100 * 227 + 0x10000 * 188;
        private static WdColor COLORQ3 = (WdColor)204 + 0x100 * 192 + 0x10000 * 217;
        private static WdColor COLORQ4 = (WdColor)253 + 0x100 * 233 + 0x10000 * 217;
        private static WdColor COLORQ5 = (WdColor)198 + 0x100 * 217 + 0x10000 * 241;
        private static int PAR_OCORRENCIAS_VERIFICADAS = 62;
        private static int PAR_TABLE_ACHADOS_Q1 = 131;
        private static int PAR_NOVA_TABLE_ACHADOS_Q1 = 153;
        private static int PAR_TABLE_ACHADOS_Q2 = 159;
        private static int PAR_NOVA_TABLE_ACHADOS_Q2 = 181;
        private static int PAR_TABLE_ACHADOS_Q3 = 187;
        private static int PAR_NOVA_TABLE_ACHADOS_Q3 = 209;
        private static int PAR_TABLE_ACHADOS_Q4 = 215;
        private static int PAR_NOVA_TABLE_ACHADOS_Q4 = 237;
        private static int PAR_TABLE_ACHADOS_Q5 = 243;
        private static int PAR_NOVA_TABLE_ACHADOS_Q5 = 265;
        private static int PAR_CONCLUSAO_IRREGULARIDADE = 271;
        private static int PAR_CONCLUSAO_IMPROPRIEDADE = 273;
        private static int PAR_MATRIZ_RESPONSABILIZACAO = 444;
        private static string XMLPATH = AppContext.BaseDirectory + "\\Checklist.xml";
        public string checklistXMLPartID = string.Empty;
        private Office.CustomXMLPart checklistXMLPart;
        private XmlDocument xd;
        private XmlNamespaceManager xnm;
        private static string NODEPATH = "ns:checklists/ns:checklist/ns:pa";
        private static string ISSELECTEDPATH = "/ns:isSelected";
        private static string DESCRICAOACHADOPATH = "/ns:descricaoAchado";
        private static string QUESTAOAUDITORIAPATH = "/ns:qa";
        private static string SITUACAOENCONTRADAPATH = "/ns:situacaoEncontrada";
        private static string CAUSAPATH = "/ns:causa";
        private static string EFEITOPATH = "/ns:efeito";
        private static string OBRIGATORIEDADEPATH = "/ns:obrigatoriedade";
        private static string CRITERIOSPATH = "/ns:criterios";
        private static string ORGAOPATH = "ns:checklists/ns:checklist/ns:orgao";        
        private static string NOMERESPONSAVELPATH = "ns:checklists/ns:checklist/ns:nomeResponsavel";
        private static string CARGORESPONSAVELPATH = "ns:checklists/ns:checklist/ns:cargoResponsavel";
        private static string DATAAVALIACAOPATH = "ns:checklists/ns:checklist/ns:dataAvaliacao";        
        private static string NUMACHADO = "$NUM_ACHADO$";
        private static string DESCACHADO = "$DESC_ACHADO$";
        private static string STRORGAO = "$ORGAO$";
        private static string STRPERIODO = "$PERIODO$";
        private static string STRMESANO = "$MES_ANO$";
        private static string STREXERCICIO = "$EXERCICIO$";
        private static string STRSITUACAOENCONTRADA = "$SITUACAO_ENCONTRADA$";
        private static string STRARTIGOS = "$ARTIGOS$";
        private static string STRCAUSA = "$CAUSA$";
        private static string STREFEITO = "$EFEITO$";
        private static string STRDATA = "$DATA$";
        private static string STRCONDUTARESPONSAVEL = "$CONDUTA_RESPONSAVEL$";
        private static string STRNEXOCAUSALIDADE = "$NEXO_CAUSALIDADE$";
        private static string STRNOMERESPONSAVEL = "$NOME_RESPONSAVEL$";
        private static string STRCARGORESPONSAVEL = "$CARGO_RESPONSAVEL$";
        private string orgao;
        private string nomeResponsavel;
        private string cargoResponsavel;
        private string dataAvaliacao;
        private List<int> listaIntAchados;
        private List<string> listaStringAchados;
        private List<string> listaSituacaoEncontrada;
        private List<string> listaCausa;
        private List<string> listaEfeito;
        private List<string> listaArtigos;
        private List<int> listaQuestoesAuditoriaAchados;
        private List<int> listaOrdemAchados;
        private List<bool> listaObrigatoriedadeAchados;
        private Paragraph paragrafoOcorrenciasVerificadas;
        private Paragraph paragrafoTableAchadosQ1;
        private Paragraph paragrafoNovasTableAchadosQ1;
        private Table tableAchadoQ1;
        private Paragraph paragrafoTableAchadosQ2;
        private Paragraph paragrafoNovasTableAchadosQ2;
        private Table tableAchadoQ2;
        private Paragraph paragrafoTableAchadosQ3;
        private Paragraph paragrafoNovasTableAchadosQ3;
        private Table tableAchadoQ3;
        private Paragraph paragrafoTableAchadosQ4;
        private Paragraph paragrafoNovasTableAchadosQ4;
        private Table tableAchadoQ4;
        private Paragraph paragrafoTableAchadosQ5;
        private Paragraph paragrafoNovasTableAchadosQ5;
        private Table tableAchadoQ5;
        private Paragraph paragrafoConclusaoIrregularidade;
        private Paragraph paragrafoConclusaoImpropriedade;
        private Paragraph paragrafoMatrizResponsabilizacao;
        private Table tableMatrizResponsabilizacao;

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            Clipboard.Clear();

            string xmlData = GetXmlFromResource();

            if (xmlData == null || xmlData == "")
                return;

            AddCustomXmlPart(xmlData);
            readXMLFile();
            readAchados();

            paragrafoConclusaoIrregularidade = Content.Paragraphs[PAR_CONCLUSAO_IRREGULARIDADE];
            paragrafoConclusaoImpropriedade = Content.Paragraphs[PAR_CONCLUSAO_IMPROPRIEDADE];
            paragrafoOcorrenciasVerificadas = Content.Paragraphs[PAR_OCORRENCIAS_VERIFICADAS];
            paragrafoMatrizResponsabilizacao = Content.Paragraphs[PAR_MATRIZ_RESPONSABILIZACAO];

            string t1 = paragrafoConclusaoIrregularidade.Previous().Range.Text;
            string t2 = paragrafoConclusaoIrregularidade.Range.Text;
            string t3 = paragrafoConclusaoIrregularidade.Next().Range.Text;     

            tableMatrizResponsabilizacao = paragrafoMatrizResponsabilizacao.Range.Tables[1];
            string teste = tableMatrizResponsabilizacao.Rows[7].Cells[1].Range.Text;           

            paragrafoTableAchadosQ1 = Content.Paragraphs[PAR_TABLE_ACHADOS_Q1];
            tableAchadoQ1 = paragrafoTableAchadosQ1.Range.Tables[1];
            paragrafoNovasTableAchadosQ1 = Content.Paragraphs[PAR_NOVA_TABLE_ACHADOS_Q1];

            paragrafoTableAchadosQ2 = Content.Paragraphs[PAR_TABLE_ACHADOS_Q2];
            tableAchadoQ2 = paragrafoTableAchadosQ2.Range.Tables[1];
            paragrafoNovasTableAchadosQ2 = Content.Paragraphs[PAR_NOVA_TABLE_ACHADOS_Q2];

            paragrafoTableAchadosQ3 = Content.Paragraphs[PAR_TABLE_ACHADOS_Q3];
            tableAchadoQ3 = paragrafoTableAchadosQ3.Range.Tables[1];
            paragrafoNovasTableAchadosQ3 = Content.Paragraphs[PAR_NOVA_TABLE_ACHADOS_Q3];

            paragrafoTableAchadosQ4 = Content.Paragraphs[PAR_TABLE_ACHADOS_Q4];
            tableAchadoQ4 = paragrafoTableAchadosQ4.Range.Tables[1];
            paragrafoNovasTableAchadosQ4 = Content.Paragraphs[PAR_NOVA_TABLE_ACHADOS_Q4];

            paragrafoTableAchadosQ5 = Content.Paragraphs[PAR_TABLE_ACHADOS_Q5];
            tableAchadoQ5 = paragrafoTableAchadosQ5.Range.Tables[1];
            paragrafoNovasTableAchadosQ5 = Content.Paragraphs[PAR_NOVA_TABLE_ACHADOS_Q5];

            replaceAchadosMatrizResponsabilizacao(tableMatrizResponsabilizacao);
            replaceAchadosConclusaoImpropriedade();
            replaceAchadosConclusaoIrregularidade();

            DateTime today = DateTime.Today;
            string strDate = today.ToLongDateString().Split(',')[1];
            SearchReplace(STRDATA, strDate.Substring(1));

            int achado = 0;
            achado = replaceTablesAchados(paragrafoNovasTableAchadosQ1, 1, achado);
            achado = replaceTablesAchados(paragrafoNovasTableAchadosQ2, 2, achado);
            achado = replaceTablesAchados(paragrafoNovasTableAchadosQ3, 3, achado);
            achado = replaceTablesAchados(paragrafoNovasTableAchadosQ4, 4, achado);
            achado = replaceTablesAchados(paragrafoNovasTableAchadosQ5, 5, achado);

            replaceAchadosResumo();

            SearchReplace(STRORGAO, orgao);
            string[] dateArray = dataAvaliacao.Split(' ');
            if (dateArray == null || dateArray.Length < 5)
                return;

            string mesAno = dateArray[2] + "/" + dateArray[4];            
            SearchReplace(STRMESANO, mesAno);
            SearchReplace(STRPERIODO, dataAvaliacao);
            SearchReplace(STRNOMERESPONSAVEL, nomeResponsavel);
            SearchReplace(STRCARGORESPONSAVEL, cargoResponsavel);
        }


        private void replaceAchadosMatrizResponsabilizacao(Table matriz)
        {
            List<string> listaBusca = new List<string>();

            for(int i = 0; i < matriz.Rows[7].Cells.Count; i++)
            {
                listaBusca.Add(matriz.Rows[7].Cells[i + 1].Range.Text);
            }

            Range formatado = matriz.Rows[7].Range.FormattedText;
            

            for (int i = 0; i < listaOrdemAchados.Count; i++)
            {
                string strAchado = "Achado ";

                if (i + 1 < 10)
                    strAchado += "0" + (i + 1) + ".";
                else
                    strAchado += (i + 1) + ".";

                string substituir = strAchado + " " + listaStringAchados[listaOrdemAchados[i]];

                if (dataAvaliacao != null)
                {
                    string[] splitDataAvaliacao = dataAvaliacao.Split(' ');
                    if (splitDataAvaliacao != null && splitDataAvaliacao.Length >= 4)
                    {
                        string exercicio = dataAvaliacao.Split(' ')[4];
                        SearchReplace(STREXERCICIO, exercicio, matriz.Rows[matriz.Rows.Count].Range);
                    }
                }

                SearchReplace(DESCACHADO, substituir, matriz.Rows[matriz.Rows.Count].Range);
                SearchReplace(STRNOMERESPONSAVEL, nomeResponsavel, matriz.Rows[matriz.Rows.Count].Range);
                SearchReplace(STRCARGORESPONSAVEL, cargoResponsavel, matriz.Rows[matriz.Rows.Count].Range);
                

                int paOrdem = listaIntAchados[listaOrdemAchados[i]];
                string pa = "";

                if (paOrdem < 10)
                    pa = "0" + paOrdem;
                else
                    pa = paOrdem.ToString();

                string consultaCriterio = NODEPATH + pa + CRITERIOSPATH + "/ns:criterio";
                XmlNodeList nodeList = xd.SelectNodes(consultaCriterio, xnm);

                string artigos = "";
                foreach (XmlNode node in nodeList)
                {
                    string art = node.SelectSingleNode("ns:artigo", xnm).InnerText.Trim('\n', '\r', '\t');
                    if (node != nodeList[0])
                        art = "; " + art;

                    artigos += art;
                }

                string consultaEfeito = NODEPATH + pa + EFEITOPATH;
                string consultaDescricao = NODEPATH + pa + DESCRICAOACHADOPATH;
                string consultaCausa = NODEPATH + pa + CAUSAPATH;
                XmlNode nodeEfeito = xd.SelectSingleNode(consultaEfeito, xnm);
                XmlNode nodeDescricao = xd.SelectSingleNode(consultaDescricao, xnm);
                XmlNode nodeCausa = xd.SelectSingleNode(consultaCausa, xnm);

                string efeito = nodeEfeito.InnerText.Trim('\r', '\a');
                string descricao = nodeDescricao.InnerText.Trim('\r', '\a');
                string nexo = descricao.Trim('.') + " " + efeito.ToLower().Trim('.') + "," + " e afronta: " + artigos;
                string causa = nodeCausa.InnerText.Trim('\r', '\a');

                SearchReplace(STRCONDUTARESPONSAVEL, causa, matriz.Rows[matriz.Rows.Count].Range);
                SearchReplace(STRNEXOCAUSALIDADE, nexo, matriz.Rows[matriz.Rows.Count].Range);

                if (i + 1 < listaOrdemAchados.Count)
                {
                    int rowCountAntes = matriz.Rows.Count;
                    Row r = matriz.Rows.Add();

                    for(int cell = 0; cell < r.Cells.Count; cell++)
                    {
                        r.Cells[cell + 1].Range.Text = listaBusca[cell];
                    }

                    int rowCountDepois = matriz.Rows.Count;

                    if(rowCountDepois > rowCountAntes + 1)
                    {
                        matriz.Rows[matriz.Rows.Count].Delete();
                    }
                }
            }
        }

        private void replaceAchadosConclusao(bool isObrigatorio, Paragraph paragrafo)
        {
            Paragraph par = null;
            int achados = 0;

            for (int i = 0; i < listaOrdemAchados.Count; i++)
            {
                if (i == 0)
                    par = paragrafo;

                if (listaObrigatoriedadeAchados[listaOrdemAchados[i]] == !isObrigatorio)
                    continue;

                achados++;

                if (i + 1 < listaOrdemAchados.Count)
                {
                    for(int x = i + 1; x < listaOrdemAchados.Count; x++)
                    {
                        bool b = listaObrigatoriedadeAchados[listaOrdemAchados[x]];
                        if(b == isObrigatorio)
                        {
                            par.Range.InsertParagraphAfter();

                            Paragraph next = par.Next();

                            next.Range.FormattedText = par.Range.FormattedText;
                            break;
                        }
                    }                    
                }

                string strAchado = "Achado ";
                string pa = "";

                int paOrdem = listaIntAchados[listaOrdemAchados[i]];
                if (paOrdem < 10)
                    pa = "0" + paOrdem;
                else
                    pa = paOrdem.ToString();

                strAchado += pa + ".";
                SearchReplace(DESCACHADO, listaStringAchados[listaOrdemAchados[i]].Trim('.'), par.Range);

                string consulta = NODEPATH + pa + CRITERIOSPATH + "/ns:criterio";
                XmlNodeList nodeList = xd.SelectNodes(consulta, xnm);

                string artigos = "";
                foreach (XmlNode node in nodeList)
                {
                    string art = node.SelectSingleNode("ns:artigo", xnm).InnerText.Trim('\n', '\r', '\t');
                    if (node != nodeList[0])
                        art = "; " + art;

                    artigos += art;
                }

                SearchReplace(STRARTIGOS, artigos, par.Range);

                par = par.Next();
            }

            if (achados == 0 && par != null)
                par.Range.Text = "";
        }

        private void replaceAchadosConclusaoIrregularidade()
        {
            replaceAchadosConclusao(true, paragrafoConclusaoIrregularidade);
        }

        private void replaceAchadosConclusaoImpropriedade()
        {
            replaceAchadosConclusao(false, paragrafoConclusaoImpropriedade);
        }

        private void replaceAchadosResumo()
        {
            Paragraph par = null;
            par = paragrafoOcorrenciasVerificadas;

            for (int i = 0; i < listaOrdemAchados.Count; i++)
            {
                if (i + 1 < listaOrdemAchados.Count)
                {
                    par.Range.InsertParagraphAfter();
                    Paragraph next = par.Next();
                    next.Range.FormattedText = par.Range.FormattedText;
                }

                string strAchado = "Achado ";

                if (i + 1 < 10)
                    strAchado += "0" + (i + 1) + ".";
                else
                    strAchado += (i + 1) + ".";

                SearchReplace(NUMACHADO, strAchado, par.Range);
                SearchReplace(DESCACHADO, listaStringAchados[listaOrdemAchados[i]], par.Range);
                par = par.Next();
            }
        }

        private int replaceTablesAchados(Paragraph paragrafoNovaTable, int questaoAuditoria, int achado)
        {
            Paragraph par = paragrafoNovaTable;                        
            Table t = null;

            switch(questaoAuditoria)
            {
                case 1: t = tableAchadoQ1; break;
                case 2: t = tableAchadoQ2; break;
                case 3: t = tableAchadoQ3; break;
                case 4: t = tableAchadoQ4; break;
                case 5: t = tableAchadoQ5; break;
            }

            Range formatado = t.Range.FormattedText.Duplicate;
            string txtPesquisa = formatado.Text;

            int antes = achado;
            int count = listaOrdemAchados.Count;

            for (int i = 0; i < listaOrdemAchados.Count; i++)
            {
                Word.Table tableCopy = null;

                int questao = listaQuestoesAuditoriaAchados[listaOrdemAchados[i]];

                if (questao != questaoAuditoria)
                    continue;

                string strAchado = "Achado ";                

                if (achado + 1 < 10)
                    strAchado += "0" + (achado + 1) + ".";
                else
                    strAchado += (achado + 1) + ".";

                if (t == null)
                    break;                

                SearchReplace(NUMACHADO, strAchado, t.Range);
                SearchReplace(DESCACHADO, listaStringAchados[listaOrdemAchados[i]], t.Range);
                SearchReplace(STRSITUACAOENCONTRADA, listaSituacaoEncontrada[listaOrdemAchados[i]], t.Range);
                SearchReplace(STRCAUSA, listaCausa[listaOrdemAchados[i]], t.Range);
                SearchReplace(STREFEITO, listaEfeito[listaOrdemAchados[i]], t.Range);
                SearchReplace(STRARTIGOS, listaArtigos[listaOrdemAchados[i]], t.Range);

                achado++;

                if (i > 0)
                {
                    par.Range.InsertParagraphBefore();
                    par = par.Previous();
                }

                if ((i + 1 >= listaOrdemAchados.Count))
                    break;

                if (listaQuestoesAuditoriaAchados[listaOrdemAchados[i + 1]] != listaQuestoesAuditoriaAchados[listaOrdemAchados[i]])
                    break;

                tableCopy = addTableAchado(par.Range, questaoAuditoria);
                t = tableCopy;

                if (par.Next().Range.Tables.Count == 0)
                {
                    par = par.Next();
                }
                else
                {
                    par.Range.InsertParagraphBefore();
                    par = par.Previous();
                    par.Range.InsertParagraphAfter();
                    par = par.Next();
                }
            }

            if (antes == achado)
                t.Delete();

            return achado;
        }       

        private Table addTableAchado(Range range, int questaoAuditoria)
        {
            Table newTable = range.Tables.Add(range, 8, 2, ref missing, ref missing);
            newTable.Rows[1].Cells[1].Range.Text = NUMACHADO;
            switch(questaoAuditoria)
            {
                case 1: newTable.Rows[1].Cells[1].Shading.BackgroundPatternColor = COLORQ1; break;
                case 2: newTable.Rows[1].Cells[1].Shading.BackgroundPatternColor = COLORQ2; break;
                case 3: newTable.Rows[1].Cells[1].Shading.BackgroundPatternColor = COLORQ3; break;
                case 4: newTable.Rows[1].Cells[1].Shading.BackgroundPatternColor = COLORQ4; break;
                case 5: newTable.Rows[1].Cells[1].Shading.BackgroundPatternColor = COLORQ5; break;
            }

            newTable.Rows[1].Cells[1].SetWidth(137, WdRulerStyle.wdAdjustProportional);
            newTable.Rows[1].Cells[2].Range.Text = DESCACHADO;
            newTable.Rows[1].Cells[2].Range.Font.Bold = 0;
            newTable.Rows[2].Cells[1].Range.Text = "Situação encontrada(NAG 4111.3.2) " + "\r\a" + STRSITUACAOENCONTRADA;
            newTable.Rows[2].Cells.Merge();
            newTable.Rows[2].Cells[1].Range.Paragraphs[2].Range.Font.Bold = 0;

            newTable.Rows[3].Cells[1].Range.Text = "Evidência (NAG 1113)";
            newTable.Rows[3].Cells.Merge();
            newTable.Rows[4].Cells[1].Range.Text = "Critério Legal (NAG 4111.3.1) " + "\r\a" + STRARTIGOS;
            newTable.Rows[4].Cells.Merge();
            newTable.Rows[4].Cells[1].Range.Paragraphs[2].Range.Font.Bold = 0;
            newTable.Rows[5].Cells[1].Range.Text = "Causa (NAG 4111.3.3) " + "\r\a" + STRCAUSA;
            newTable.Rows[5].Cells.Merge();
            newTable.Rows[5].Cells[1].Range.Paragraphs[2].Range.Font.Bold = 0;
            newTable.Rows[6].Cells[1].Range.Text = "Efeito (NAG 4111.3.4) " + "\r\a" + STREFEITO;
            newTable.Rows[6].Cells.Merge();
            newTable.Rows[6].Cells[1].Range.Paragraphs[2].Range.Font.Bold = 0;
            newTable.Rows[7].Cells[1].Range.Text = "Opinião do Auditado (NAG 4111.3.5)";
            newTable.Rows[7].Cells.Merge();
            newTable.Rows[8].Cells[1].Range.Text = "Conclusão (NAG 4111.3.6)";
            newTable.Rows[8].Cells.Merge();
            newTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            newTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;

            while(newTable.Rows.Count > 8)
            {
                int count = newTable.Rows.Count;
                newTable.Rows[count].Delete();
            }

            return newTable;
        }

        private void readAchados()
        {
            int itemCount = 0;

            listaIntAchados = new List<int>();
            listaStringAchados = new List<string>();
            listaSituacaoEncontrada = new List<string>();
            listaArtigos = new List<string>();
            listaCausa = new List<string>();
            listaEfeito = new List<string>();
            listaQuestoesAuditoriaAchados = new List<int>();
            listaOrdemAchados = new List<int>();
            listaObrigatoriedadeAchados = new List<bool>();

            string path = NODEPATH + "01";
            System.Xml.XmlNode node = xd.SelectSingleNode(path, xnm);
            string strItem = null;
            string strNextItem = null;

            while (node != null)
            {
                if (itemCount + 1 < 10)
                {
                    strItem = "0" + (itemCount + 1);
                }
                else
                {
                    strItem = "" + (itemCount + 1);
                }

                XmlNode cn = xd.SelectSingleNode(NODEPATH + strItem + ISSELECTEDPATH, xnm);

                bool itemStatus = false;

                if (cn != null)
                    itemStatus = Boolean.Parse(cn.InnerText);

                if (itemCount + 2 < 10)
                {
                    strNextItem = "0" + (itemCount + 2);
                }
                else
                {
                    strNextItem = "" + (itemCount + 2);
                }

                node = xd.SelectSingleNode(NODEPATH + strNextItem, xnm);

                if (!itemStatus)
                {
                    itemCount++;
                    continue;
                }

                XmlNode descricao = xd.SelectSingleNode(NODEPATH + strItem + DESCRICAOACHADOPATH, xnm);
                XmlNode questao = xd.SelectSingleNode(NODEPATH + strItem + QUESTAOAUDITORIAPATH, xnm);
                XmlNode situacao = xd.SelectSingleNode(NODEPATH + strItem + SITUACAOENCONTRADAPATH, xnm);
                XmlNode causa = xd.SelectSingleNode(NODEPATH + strItem + CAUSAPATH, xnm);
                XmlNode efeito = xd.SelectSingleNode(NODEPATH + strItem + EFEITOPATH, xnm);
                XmlNode criterios = xd.SelectSingleNode(NODEPATH + strItem + CRITERIOSPATH, xnm);
                XmlNode obrigatoriedade = xd.SelectSingleNode(NODEPATH + strItem + OBRIGATORIEDADEPATH, xnm);

                listaIntAchados.Add(Int32.Parse(strItem));
                listaStringAchados.Add(descricao.InnerText.Trim('\r','\a'));
                listaQuestoesAuditoriaAchados.Add(Int32.Parse(questao.InnerText.Trim('\r', '\a')));
                listaSituacaoEncontrada.Add(situacao.InnerText.Trim('\r', '\a'));
                listaCausa.Add(causa.InnerText.Trim('\r', '\a'));
                listaEfeito.Add(efeito.InnerText.Trim('\r', '\a'));
                listaArtigos.Add(concatenaArtigos(strItem));

                string strObrigatoriedade = obrigatoriedade.InnerText.Trim('\r', '\a');
                bool boolObrigatoriedade = true;

                if (strObrigatoriedade != "Recomendada")
                    boolObrigatoriedade = true;
                else
                    boolObrigatoriedade = false;

                listaObrigatoriedadeAchados.Add(boolObrigatoriedade);

                itemCount++;
            }

            int ordem = 1;

            while(ordem <= 5)
            {
                for(int i = 0; i < listaQuestoesAuditoriaAchados.Count; i++)
                {
                    if(listaQuestoesAuditoriaAchados[i] == ordem)
                        listaOrdemAchados.Add(i);
                }
                ordem++;
            }
        }

        private string concatenaArtigos(string strItem)
        {
            string artigos = "";
            string consulta = NODEPATH + strItem + CRITERIOSPATH + "/ns:criterio";
            XmlNodeList nodeList = xd.SelectNodes(consulta, xnm);
            foreach(XmlNode node in nodeList)
            {
                XmlNode nodeArtigo = node.SelectSingleNode("ns:artigo", xnm);
                string artigo = nodeArtigo.InnerText.Trim('\r', '\n', '\t');

                if (node != nodeList.Item(0))
                    artigo = "\r" + artigo;

                XmlNode nodeDescricao = node.SelectSingleNode("ns:descricao", xnm);
                string descricao = "\r\"" + nodeDescricao.InnerText.Trim('\r', '\n', '\t') + "\"\r";

                artigos += artigo + descricao;
            }
            
            return artigos;
        }

        private void readXMLFile()
        {
            xd = new XmlDocument();
            string str = checklistXMLPart.XML;
            xd.LoadXml(str);

            xnm = new XmlNamespaceManager(xd.NameTable);
            xnm.AddNamespace("ns", "http://schemas.microsoft.com/vsto/samples");

            XmlNode nodeOrgao = xd.SelectSingleNode(ORGAOPATH, xnm);
            XmlNode nodeNomeResponsavel = xd.SelectSingleNode(NOMERESPONSAVELPATH, xnm);
            XmlNode nodeCargoResponsavel = xd.SelectSingleNode(CARGORESPONSAVELPATH, xnm);
            XmlNode nodeDataAvaliacao = xd.SelectSingleNode(DATAAVALIACAOPATH, xnm);

            orgao = nodeOrgao.InnerText;
            nomeResponsavel = nodeNomeResponsavel.InnerText;
            cargoResponsavel = nodeCargoResponsavel.InnerText;
            dataAvaliacao = nodeDataAvaliacao.InnerText;
        }

        private void SearchReplace(string original, string substituto)
        {
            if (substituto == null || substituto == "")
                return;

            Word.Find findObject = Application.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = original;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = substituto;

            object replaceAll = Word.WdReplace.wdReplaceAll;

            bool found = findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceAll, ref missing, ref missing, ref missing, ref missing);

            bool teste = found;
        }

        private void SearchReplace(string original, string substituto, Range range)
        {
            Word.Find findObject = range.Find;
            findObject.ClearFormatting();
            findObject.Text = original;
            findObject.Replacement.ClearFormatting();

            object replaceAll = Word.WdReplace.wdReplaceNone;

            bool found = findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceAll, ref missing, ref missing, ref missing, ref missing);

            if(found)
                range.Text = substituto;
        }

        private string GetXmlFromResource()
        {
            string filename = XMLPATH;

            FileStream stream1 = null;
            try
            {
                stream1 = new FileStream(filename, FileMode.Open);
            }
            catch (FileNotFoundException fnfe)
            {
                MessageBox.Show("Arquivo Checklist.xml não encontrado", "Aviso");
            }

            if (stream1 == null)
                return null;

            using (System.IO.StreamReader resourceReader = new System.IO.StreamReader(stream1))
            {
                if (resourceReader != null)
                {
                    return resourceReader.ReadToEnd();
                }
            }

            return null;
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
            this.Startup += new System.EventHandler(ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion
    }
}
