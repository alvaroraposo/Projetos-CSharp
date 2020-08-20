using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace AnaliseDadosXML
{
    public partial class Plan2
    {
        private static string ORGAOPATH = "ns:checklists/ns:checklist/ns:orgao";
        private static string NOMERESPONSAVELPATH = "ns:checklists/ns:checklist/ns:nomeResponsavel";
        private static string CARGORESPONSAVELPATH = "ns:checklists/ns:checklist/ns:cargoResponsavel";
        private static string DATAAVALIACAOPATH = "ns:checklists/ns:checklist/ns:dataAvaliacao";
        private static string NODEPATH = "ns:checklists/ns:checklist/ns:pa";
        private static string ORIENTACAOPATH = "/ns:orientacao";
        private static string ISSELECTEDPATH = "/ns:isSelected";
        private static string PESOPATH = "/ns:peso";
        private static string XMLPATH = AppDomain.CurrentDomain.BaseDirectory + "Checklist.xml";
        private static string ANO = "2019";
        private static string SEMESTRE = "01 Semestre";
        public static int INICIOITEMAUDITORIAROW = 3;
        public static int INICIOITEMAUDITORIACOL = 2;
        public static int INICIOORGAOROW = 2;
        public static int INICIOORGAOCOL = 3;
        public string checklistXMLPartID = string.Empty;
        private Office.CustomXMLPart checklistXMLPart;
        private XmlDocument xd;
        private XmlNamespaceManager xnm;
        private string orgao;
        private string nomeResponsavel;
        private string cargoResponsavel;
        private string dataAvaliacao;
        public int linhaTotal = 0;
        public int fimColunaGrafico = 0;

        private void Plan2_Startup(object sender, System.EventArgs e)
        {
            
        }

        public void carregarPlan2()
        {
            string path = "02-Setorial\\DIATI\\AUDITORIAS\\" + ANO + "\\02. Portais do Poder Legislativo\\*.*";
//            string path = "02-Setorial\\DIATI\\AUDITORIAS\\2019\\03. Outros\\*.*";
            string pasta = SEMESTRE;            
//            string path = "03-Gerencial\\Portais\\*.*";
            int orgaoCol = AnaliseDadosXML.Plan2.INICIOORGAOCOL;

            System.IO.DriveInfo di = new System.IO.DriveInfo("Z:");
            if (!di.IsReady)
            {
                Console.WriteLine("The drive {0} could not be read", di.Name);
                return;
            }

            DirectoryInfo root = di.RootDirectory;
            DirectoryInfo[] orgaos = root.GetDirectories(path);

            foreach (DirectoryInfo orgao in orgaos)
            {
                DirectoryInfo[] mes = orgao.GetDirectories(pasta);

                if (mes == null || mes.Length == 0)
                    continue;

                FileInfo[] fileInfo = mes[0].GetFiles("Checklist.xml");

                foreach (FileInfo f in fileInfo)
                {
                    string xmlData = Globals.Plan2.GetXmlFromResource(f);

                    if (xmlData == null || xmlData == "")
                        return;

                    Excel.Workbook WB = Globals.Plan2.Application.ActiveWorkbook;
                    AddCustomXmlPart(xmlData, WB);
                    readXMLFile();

                    if (orgaoCol == INICIOORGAOCOL)
                        readPontosAuditoria();

                    readXML(orgaoCol);
                    orgaoCol++;
                }
            }

            fimColunaGrafico = orgaoCol - 1;
        }

        public void readPontosAuditoria()
        {
            int item = 0;
            string strItem = "";
            XmlNode nodeOrientacao = null;

            do
            {
                if (item + 1 >= 10)
                {
                    strItem = "" + (item + 1);
                }
                else
                {
                    strItem = "0" + (item + 1);
                }


                string query = NODEPATH + strItem + ORIENTACAOPATH;
                nodeOrientacao = xd.SelectSingleNode(query, xnm);
                if (nodeOrientacao == null)
                    break;

                Cells[INICIOITEMAUDITORIACOL][INICIOITEMAUDITORIAROW + item] = nodeOrientacao.InnerText.Trim('\r', '\a');
                item++;
            }
            while (nodeOrientacao != null);

            linhaTotal = INICIOITEMAUDITORIAROW + item + 1;
            Cells[INICIOITEMAUDITORIACOL][linhaTotal] = "TOTAL:";            
            Cells.Columns.AutoFit();
            Rows.AutoFit();
        }

        public void readXMLFile()
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


        public string GetXmlFromResource(FileInfo fileInfo)
        {

            FileStream stream1 = null;
            try
            {
                stream1 = fileInfo.OpenRead();
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

        // Define GetShortPathName API function.
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern uint GetShortPathName(string lpszLongPath, char[] lpszShortPath, int cchBuffer);

        public void readXML(int orgaoCol)
        {
            readDados(orgaoCol);

            string mostra = null;
            if (orgao == null || orgao == "")
                return;

            string[] nomeOrgao = orgao.Split(' ');

            if (nomeOrgao == null || nomeOrgao.Length < 4)
                return;

            if(nomeOrgao[0] == "Prefeitura" || nomeOrgao[0] == "prefeitura")
            {
                mostra = "PM";

                for(int i = 0; i < nomeOrgao.Length; i++)
                {
                    if (i <= 2)
                        continue;

                    mostra += " " + nomeOrgao[i];
                }
            }
            else if (nomeOrgao[0] == "Câmara" || nomeOrgao[0] == "camara")
            {
                mostra = "CM";

                for (int i = 0; i < nomeOrgao.Length; i++)
                {
                    if (i <= 2)
                        continue;

                    mostra += " " + nomeOrgao[i];
                }
            }
            else if(nomeOrgao[0] == "Governo" || nomeOrgao[0] == "governo")
            {
                mostra = "Governo";
            }
            else if(nomeOrgao[0] == "Assembleia" || nomeOrgao[0] == "assembleia")
            {
                mostra = "Assembleia";
            }


            Cells[orgaoCol][INICIOORGAOROW] = mostra;
            Cells.Columns.AutoFit();
        }

        private void readDados(int orgaoCol)
        {
            int item = 0;
            string strItem = "";
            XmlNode nodeIsSelected = null;

            do
            {
                if (item + 1 >= 10)
                {
                    strItem = "" + (item + 1);
                }
                else
                {
                    strItem = "0" + (item + 1);
                }

                string query = NODEPATH + strItem + ISSELECTEDPATH;
                nodeIsSelected = xd.SelectSingleNode(query, xnm);

                if (nodeIsSelected == null)
                    break;

                bool isSelected = Boolean.Parse(nodeIsSelected.InnerText.Trim('\r', '\a'));                

                if (isSelected)
                {
                    Cells[orgaoCol][INICIOITEMAUDITORIAROW + item] = 0;
                    Cells[orgaoCol][INICIOITEMAUDITORIAROW + item].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }
                else
                {
                    query = NODEPATH + strItem + PESOPATH;
                    XmlNode peso = xd.SelectSingleNode(query, xnm);
                    if (peso == null)
                        return;

                    float floatPeso = float.Parse(peso.InnerText.Trim('\r', '\a'));
                    Cells[orgaoCol][INICIOITEMAUDITORIAROW + item] = floatPeso / 100;
                    Cells[orgaoCol][INICIOITEMAUDITORIAROW + item].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Aqua);
                }

                item++;
            }
            while (nodeIsSelected != null);
            
            string coluna = numColumntoLetter(orgaoCol - 1);
            string cellInicio = coluna + INICIOITEMAUDITORIAROW;
            string cellFim = coluna + (INICIOITEMAUDITORIAROW + item - 1);
            string range = cellInicio + ":" + cellFim;
            string soma = "=SUM(" + cellInicio + ":" + cellFim + ")";
            
            Cells[orgaoCol][INICIOITEMAUDITORIAROW + item + 1] = soma;
        }

        public string numColumntoLetter(int intCol)
        {
            int intFirstLetter = ((intCol) / 676) + 64;
            int intSecondLetter = ((intCol % 676) / 26) + 64;
            int intThirdLetter = (intCol % 26) + 65;

            char FirstLetter = (intFirstLetter > 64) ? (char)intFirstLetter : ' ';
            char SecondLetter = (intSecondLetter > 64) ? (char)intSecondLetter : ' ';
            char ThirdLetter = (char)intThirdLetter;

            return string.Concat(FirstLetter, SecondLetter, ThirdLetter).Trim();
        }

        public void AddCustomXmlPart(string xmlData, Excel.Workbook document)
        {
            if (xmlData != null && xmlData != "")
            {                
                checklistXMLPart = document.CustomXMLParts.SelectByID(checklistXMLPartID);
                checklistXMLPart = document.CustomXMLParts.Add(xmlData);
                checklistXMLPart.NamespaceManager.AddNamespace("ns", "http://schemas.microsoft.com/vsto/samples");
                checklistXMLPartID = checklistXMLPart.Id;
            }
        }

        private void Plan2_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código gerado pelo Designer VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Plan2_Startup);
            this.Shutdown += new System.EventHandler(Plan2_Shutdown);
        }

        #endregion

    }
}
