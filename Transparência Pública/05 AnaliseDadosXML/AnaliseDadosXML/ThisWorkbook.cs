using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace AnaliseDadosXML
{
    public partial class ThisWorkbook
    {
        private static string ANO = "2019";
        private static string MES = "05 - Maio";
        private static int INICIOORGAOROW = 3;
        private static int ORGAOCOL = 1;
        private static int PERCENTUALCOL = 2;

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            carregarPlan2();            
            carregarPlan1();
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void carregarPlan1()
        {
            int inicioPlan2Col = AnaliseDadosXML.Plan2.INICIOORGAOCOL;
            int inicioPlan2OrgaoRow = AnaliseDadosXML.Plan2.INICIOORGAOROW;
            int inicioPlan2TotalRow = Globals.Plan2.linhaTotal;
            int inicioPlan1Row = INICIOORGAOROW;
            int fimPlan2Col = Globals.Plan2.fimColunaGrafico;
            object misValue = System.Reflection.Missing.Value;

            for (int i = inicioPlan2Col; i <= fimPlan2Col; i++)
            {
                var cell = (Excel.Range)Globals.Plan2.Cells[i][inicioPlan2OrgaoRow];
                Globals.Plan1.Cells[ORGAOCOL][inicioPlan1Row] = cell.Value;

                var cell2 = (Excel.Range)Globals.Plan2.Cells[i][inicioPlan2TotalRow];
                Globals.Plan1.Cells[ORGAOCOL + 1][inicioPlan1Row] = cell2.Value;

                if(cell2.Value * 100 < 20)
                {
                    Globals.Plan1.Cells[ORGAOCOL][inicioPlan1Row].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Maroon);
                    Globals.Plan1.Cells[ORGAOCOL + 1][inicioPlan1Row].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Maroon);
                }
                else if(cell2.Value * 100 < 40)
                {
                    Globals.Plan1.Cells[ORGAOCOL][inicioPlan1Row].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.OrangeRed);
                    Globals.Plan1.Cells[ORGAOCOL + 1][inicioPlan1Row].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.OrangeRed);
                }
                else if(cell2.Value * 100 < 50)
                {
                    Globals.Plan1.Cells[ORGAOCOL][inicioPlan1Row].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                    Globals.Plan1.Cells[ORGAOCOL + 1][inicioPlan1Row].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                }
                else if (cell2.Value * 100 < 70)
                {
                    Globals.Plan1.Cells[ORGAOCOL][inicioPlan1Row].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    Globals.Plan1.Cells[ORGAOCOL + 1][inicioPlan1Row].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                }
                else
                {
                    Globals.Plan1.Cells[ORGAOCOL][inicioPlan1Row].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    Globals.Plan1.Cells[ORGAOCOL + 1][inicioPlan1Row].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                }

                inicioPlan1Row++;
            }

            Globals.Plan1.Cells.Columns.AutoFit();

            string cellInicioColuna = Globals.Plan2.numColumntoLetter(ORGAOCOL - 1);
            string cellInicioRow = (inicioPlan2OrgaoRow + 1).ToString();
            string cellInicio = cellInicioColuna + cellInicioRow;

            string cellFimColuna = Globals.Plan2.numColumntoLetter(ORGAOCOL);
            string cellFimRow = (inicioPlan1Row - 1).ToString();
            string cellFim = cellFimColuna + cellFimRow;

            Excel.Range chartRange = Globals.Plan1.Range[cellInicio, cellFim];
            chartRange.Sort(chartRange.Columns[2], Excel.XlSortOrder.xlDescending);

            Globals.Plan1.Chart_1.SetSourceData(chartRange, misValue);
            Excel.Series series = (Excel.Series) Globals.Plan1.Chart_1.SeriesCollection(1);            
            series.Name = "Portais da Transparência do Poder Legislativo Amazonense";            

            inicioPlan1Row = INICIOORGAOROW;
            int point = 1;
            for (int i = inicioPlan2Col; i <= fimPlan2Col; i++)
            {
                Excel.Point p = series.Points(point);
                p.Format.Fill.ForeColor.RGB = (int) Globals.Plan1.Cells[ORGAOCOL][inicioPlan1Row].Interior.Color;
                inicioPlan1Row++;
                point++;
            }
        }
        private void carregarPlan2()
        {
            Globals.Plan2.carregarPlan2();
        }


        #region Código gerado pelo Designer VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
