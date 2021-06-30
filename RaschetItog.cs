using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;

namespace EcoIntensive
{
    public partial class RaschetItog : Form
    {
        public RaschetItog()
        {
            InitializeComponent();
        }

        private double[,] massivvichislenii = new double[10000, 10000];

        private void VichislenieFormuliEcoIntensiv(double e, double l, double r, double de, double n, int i)
        {
            for (int j = 0; j < n; j++)
            {
                double ZnachenieChislitel = 0;
                double ZnachenieZnamenatel = 0;
                for (int k = 1; k <= j + 1; k++)
                {
                    ZnachenieChislitel = ZnachenieChislitel + (1 - Math.Pow((1 - de), k));
                    ZnachenieZnamenatel = ZnachenieZnamenatel + (1 / (Math.Pow((1 + r), k)));
                }
                massivvichislenii[i, j] = (e / (de * l)) * (ZnachenieChislitel / ZnachenieZnamenatel);
            }

        }

        int tablestrok = 0;
        private void VichislenieFormuli(double p, double c, double de, double d, double e, double dmax, double dshag)
        {
            double znachenied = d;
            int j = 0;
            double tochkaconflictastart = (de * (p - c)) / e;
            double tochkaconflicta = (p - c) / e;
            if (tochkaconflictastart > d)
            {
                znachenied = tochkaconflictastart + 0.0001;
            }
            else
            {
                znachenied = d;
            };



            j = 0;
            while (znachenied < dmax && znachenied < tochkaconflicta)
            {
                massivvichislenii[0, j] = znachenied;
                massivvichislenii[1, j] = (Math.Log(1 - (de * (p - c)) / (znachenied * e)) / Math.Log(1 - de)) - 1;
                znachenied = znachenied + dshag;
                j++;
            }
            tablestrok = j;
            if (tochkaconflicta < dmax)
            {
                massivvichislenii[0, j] = tochkaconflicta;
                massivvichislenii[1, j] = (Math.Log(1 - (de * (p - c)) / (tochkaconflicta * e)) / Math.Log(1 - de)) - 1;
            }
            else
            {
                massivvichislenii[0, j] = dmax;
                massivvichislenii[1, j] = (Math.Log(1 - (de * (p - c)) / (dmax * e)) / Math.Log(1 - de)) - 1;
            };

        }

        private void RaschetItog_Load(object sender, EventArgs e)
        {
            Random rnd = new Random();
            int i = 0;

            if (Program.VidRascheta == false) 
            {
                chart1.Legends.Add("Legends0");
                chart1.Legends["Legends0"].Docking = Docking.Left;
                chart1.Legends["Legends0"].IsDockedInsideChartArea = false;
                chart1.Legends["Legends0"].TableStyle = LegendTableStyle.Wide;
                chart1.Legends["Legends0"].Alignment = StringAlignment.Center;
                chart1.Legends["Legends0"].LegendStyle = LegendStyle.Column;
                label12.Text = "Итоговый расчет эко-интенсивности";
                int rowMax = 0;
                bool pervaiacollectcia = true;
                for (int it = 0; it < Program.TecushiyEcoIntens[2]; it++)
                {
                    if (rowMax < Program.SpisokEcoIntens[it, 4]) rowMax = (int)Program.SpisokEcoIntens[it, 4];
                };

                for (int it = 0; it < Program.TecushiyEcoIntens[1]; it++)
                {
                    
                    int r, g, b;
                    if (Program.SpisokEcoIntens[it, 1] != 0 && Program.SpisokEcoIntens[it, 3] != 0)
                    {


                        if (Program.SpisokEcoIntens[it, 5] == 0 || Program.SpisokEcoIntens[it, 6] == 0)
                        {
                            VichislenieFormuliEcoIntensiv(Program.SpisokEcoIntens[it, 0], Program.SpisokEcoIntens[it, 1], Program.SpisokEcoIntens[it, 2], Program.SpisokEcoIntens[it, 3], Program.SpisokEcoIntens[it, 4], i);

                            TableItog.Columns.Add("n" + (i + 1).ToString(), "n" + (i + 1).ToString());
                            TableItog.Columns.Add("E" + (i + 1).ToString(), "E" + (i + 1).ToString());

                            if (pervaiacollectcia == true)
                            {
                                for (int numi = 0; numi < rowMax; numi++)
                                {
                                    TableItog.Rows.Add();
                                };
                                pervaiacollectcia = false;
                            };

                            r = rnd.Next(0, 255);
                            g = rnd.Next(0, 255);
                            b = rnd.Next(0, 255);
                            chart1.Series.Add("Series" + (i).ToString());
                            chart1.Series["Series" + i.ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                            chart1.Series["Series" + i.ToString()].Color = Color.FromArgb(255, r, g, b);


                            chart1.Series["Series" + i.ToString()].Legend = "Legends0";
                            chart1.Series["Series" + i.ToString()].BorderWidth = 2;
                            chart1.Series["Series" + i.ToString()].LegendText = "E" + (i + 1).ToString() + " при δ=" + Program.SpisokEcoIntens[it, 3].ToString() + " e=" + Program.SpisokEcoIntens[it, 0].ToString() + " r=" + Program.SpisokEcoIntens[it, 2].ToString() + " L=" + Program.SpisokEcoIntens[it, 1].ToString() + " N=" + Program.SpisokEcoIntens[it, 4].ToString();

                            //chart1.Legends["Legends" + i.ToString()].Alignment = StringAlignment.Center;
                            //chart1.Legends["Legends" + i.ToString()].Docking = Docking.Bottom;



                            for (int j = 0; j < Program.SpisokEcoIntens[it, 4]; j++)
                            {
                                TableItog.Rows[j].Cells[2 * i].Value = (j + 1).ToString();
                                TableItog.Rows[j].Cells[2 * i + 1].Value = String.Format("{0:#0.0##}", massivvichislenii[i, j]); ;

                                chart1.Series["Series" + i.ToString()].Points.AddXY(j + 1, Math.Round(massivvichislenii[i, j], 3));
                            }
                        }
                        else
                        {
                            double techDe = Program.SpisokEcoIntens[it, 3];
                            while (techDe < Program.SpisokEcoIntens[it, 5])
                            {
                                VichislenieFormuliEcoIntensiv(Program.SpisokEcoIntens[it, 0], Program.SpisokEcoIntens[it, 1], Program.SpisokEcoIntens[it, 2], techDe, Program.SpisokEcoIntens[it, 4], i);

                                TableItog.Columns.Add("n" + (i + 1).ToString(), "n" + (i + 1).ToString());
                                TableItog.Columns.Add("E" + (i + 1).ToString(), "E" + (i + 1).ToString());
                                if (pervaiacollectcia == true)
                                {
                                    for (int numi = 0; numi < rowMax; numi++)
                                    {
                                        TableItog.Rows.Add();
                                    };
                                    pervaiacollectcia = false;
                                };

                                r = rnd.Next(0, 255);
                                g = rnd.Next(0, 255);
                                b = rnd.Next(0, 255);
                                chart1.Series.Add("Series" + (i).ToString());
                                chart1.Series["Series" + i.ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                                chart1.Series["Series" + i.ToString()].Color = Color.FromArgb(255, r, g, b);

                                chart1.Series["Series" + i.ToString()].Legend = "Legends0";
                                chart1.Series["Series" + i.ToString()].BorderWidth = 2;
                                chart1.Series["Series" + i.ToString()].LegendText = "E" + (i + 1).ToString() + " при δ=" + techDe + " e=" + Program.SpisokEcoIntens[it, 0].ToString() + " r=" + Program.SpisokEcoIntens[it, 2].ToString() + " L=" + Program.SpisokEcoIntens[it, 1].ToString() + " N=" + Program.SpisokEcoIntens[it, 4].ToString();
                                if (TableItog.Rows.Count < Program.SpisokEcoIntens[it, 4])
                                {
                                    while (Program.SpisokEcoIntens[it, 4] - TableItog.Rows.Count > 0)
                                    {
                                        TableItog.Rows.Add();
                                    };
                                }

                                for (int j = 0; j < Program.SpisokEcoIntens[it, 4]; j++)
                                {

                                    TableItog.Rows[j].Cells[2 * i].Value = (j + 1).ToString();
                                    TableItog.Rows[j].Cells[2 * i + 1].Value = String.Format("{0:#0.0##}", massivvichislenii[i, j]); ;

                                    chart1.Series["Series" + i.ToString()].Points.AddXY(j + 1, Math.Round(massivvichislenii[i, j], 3));
                                }
                                techDe = techDe + Program.SpisokEcoIntens[it, 6];
                                i++;
                            }
                            VichislenieFormuliEcoIntensiv(Program.SpisokEcoIntens[it, 0], Program.SpisokEcoIntens[it, 1], Program.SpisokEcoIntens[it, 2], techDe, Program.SpisokEcoIntens[it, 4], i);

                            TableItog.Columns.Add("n" + (i + 1).ToString(), "n" + (i + 1).ToString());
                            TableItog.Columns.Add("E" + (i + 1).ToString(), "E" + (i + 1).ToString());
                            if (pervaiacollectcia == true)
                            {
                                for (int numi = 0; numi < rowMax; numi++)
                                {
                                    TableItog.Rows.Add();
                                };
                                pervaiacollectcia = false;
                            };
                            r = rnd.Next(0, 255);
                            g = rnd.Next(0, 255);
                            b = rnd.Next(0, 255);
                            chart1.Series.Add("Series" + (i).ToString());
                            chart1.Series["Series" + i.ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                            chart1.Series["Series" + i.ToString()].Color = Color.FromArgb(255, r, g, b);
                            chart1.Series["Series" + i.ToString()].Legend = "Legends0";
                            chart1.Series["Series" + i.ToString()].BorderWidth = 2;
                            chart1.Series["Series" + i.ToString()].LegendText = "E" + (i + 1).ToString() + " при δ=" + techDe + " e=" + Program.SpisokEcoIntens[it, 0].ToString() + " r=" + Program.SpisokEcoIntens[it, 2].ToString() + " L=" + Program.SpisokEcoIntens[it, 1].ToString() + " N=" + Program.SpisokEcoIntens[it, 4].ToString();

                            if (TableItog.Rows.Count < Program.SpisokEcoIntens[it, 4])
                            {
                                while (Program.SpisokEcoIntens[it, 4] - TableItog.Rows.Count > 0)
                                {
                                    TableItog.Rows.Add();
                                };
                            }

                            for (int j = 0; j < Program.SpisokEcoIntens[it, 4]; j++)
                            {

                                TableItog.Rows[j].Cells[2 * i].Value = (j + 1).ToString();
                                TableItog.Rows[j].Cells[2 * i + 1].Value = String.Format("{0:#0.0##}", massivvichislenii[i, j]); ;

                                chart1.Series["Series" + i.ToString()].Points.AddXY(j + 1, Math.Round(massivvichislenii[i, j], 3));
                            }
                            techDe = techDe + Program.SpisokEcoIntens[it, 6];

                        }
                        i++;
                    }
                }
            }
            else
            {
                int r, g, b;

                label12.Text = "Итоговый расчет конфликта интересов";
                for (int it = 0; it < Program.TecushiyConflictInteres[1]; it++)
                {
                    double p = 0, c = 0, de = 0, d = 0, ee = 0, dmax = 0, dshag = 0;
                    p = Program.SpisokConflictInteres[it, 0];
                    c = Program.SpisokConflictInteres[it, 1];
                    de = Program.SpisokConflictInteres[it, 2];
                    d = Program.SpisokConflictInteres[it, 3];
                    ee = Program.SpisokConflictInteres[it, 4];
                    dmax = Program.SpisokConflictInteres[it, 5];
                    dshag = Program.SpisokConflictInteres[it, 6];
                    double tochkaconflicta = 0;
                    double tochkaconflictastart = 0;
                    tochkaconflicta = (p - c) / ee;
                    tochkaconflictastart = (de * (p - c)) / ee;
                    if (tochkaconflictastart < tochkaconflicta && dshag != 0 && (((tochkaconflictastart <= dmax) && (tochkaconflictastart >= d)) || ((tochkaconflicta <= dmax) && (tochkaconflicta >= d))))
                    {
                        VichislenieFormuli(Program.SpisokConflictInteres[it, 0], Program.SpisokConflictInteres[it, 1], Program.SpisokConflictInteres[it, 2], Program.SpisokConflictInteres[it, 3], Program.SpisokConflictInteres[it, 4], Program.SpisokConflictInteres[it, 5], Program.SpisokConflictInteres[it, 6]);

                        TableItog.Columns.Add("d" + (i + 1).ToString(), "d" + (i + 1).ToString());
                        TableItog.Columns.Add("be" + (i + 1).ToString(), "β" + (i + 1).ToString() + "(d" + (i + 1).ToString() + ")");

                        r = rnd.Next(0, 255);
                        g = rnd.Next(0, 255);
                        b = rnd.Next(0, 255);

                        chart1.Series.Add("Series" + (i).ToString());
                        chart1.Series["Series" + i.ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                        chart1.Series["Series" + i.ToString()].Color = Color.FromArgb(255, r, g, b);

                        Axis ax = new Axis();
                        ax.Title = "d";
                        chart1.ChartAreas[0].AxisX = ax;
                        Axis ay = new Axis();
                        ay.Title = "β";
                        chart1.ChartAreas[0].AxisY = ay;

                        chart1.Legends.Add("Legends" + (i).ToString());
                        chart1.Legends["Legends0"].Docking = Docking.Left;
                        chart1.Legends["Legends0"].IsDockedInsideChartArea = false;
                        chart1.Legends["Legends0"].TableStyle = LegendTableStyle.Wide;
                        chart1.Legends["Legends0"].Alignment = StringAlignment.Center;
                        chart1.Legends["Legends0"].LegendStyle = LegendStyle.Column;
                        chart1.Series["Series" + i.ToString()].Legend = "Legends0";
                        chart1.Series["Series" + i.ToString()].BorderWidth = 2;
                        chart1.Series["Series" + i.ToString()].LegendText = "β при p=" + Program.SpisokConflictInteres[it, 0].ToString() + " c=" + Program.SpisokConflictInteres[it, 1].ToString() + " δ=" + Program.SpisokConflictInteres[it, 2].ToString() + " e=" + Program.SpisokConflictInteres[it, 4].ToString() + " d=" + Program.SpisokConflictInteres[it, 3].ToString() + ";" + (Program.SpisokConflictInteres[it, 3] + Program.SpisokConflictInteres[it, 6]).ToString() + "..." + Program.SpisokConflictInteres[it, 5].ToString();

                        //chart1.Legends["Legends" + i.ToString()].Alignment = StringAlignment.Center;
                        //chart1.Legends["Legends" + i.ToString()].Docking = Docking.Bottom;

                        if (TableItog.Rows.Count < tablestrok)
                        {
                            while (tablestrok - TableItog.Rows.Count > 0)
                            {
                                TableItog.Rows.Add();
                            };
                            TableItog.Rows.Add();
                        }

                        for (int j = 0; j < tablestrok; j++)
                        {
                            TableItog.Rows[j].Cells[2 * i].Value = Math.Round(massivvichislenii[0, j], 3);
                            TableItog.Rows[j].Cells[2 * i + 1].Value = Math.Round(massivvichislenii[1, j], 3);

                            chart1.Series["Series" + i.ToString()].Points.AddXY(Math.Round(massivvichislenii[0, j], 3), Math.Round(massivvichislenii[1, j], 3));
                        }
                        TableItog.Rows[tablestrok].Cells[2 * i].Value = Math.Round(massivvichislenii[0, tablestrok], 3);
                        TableItog.Rows[tablestrok].Cells[2 * i + 1].Value = Math.Round(massivvichislenii[1, tablestrok], 3);

                        chart1.Series["Series" + i.ToString()].Points.AddXY(Math.Round(massivvichislenii[0, tablestrok], 3), Math.Round(massivvichislenii[1, tablestrok], 3));

                        i++;
                    };
                }
            };
        }

        private void ExportExcel_Click(object sender, EventArgs e)
        {
            int iitog = 0, iraschet = 0;
            Excel.Application ex = new Excel.Application();
            Excel.Workbook workBook;
            Excel.Worksheet sheet;
            Excel.SeriesCollection seriesCollection;
            Excel.Series series;
            Excel.Range rng1;
            Excel.Worksheet sheetItog;
            Excel.SeriesCollection seriesCollectionItog;
            Excel.Series seriesItog;
            Excel.Range rng1Itog;
            double tochkaconflicta = 0;
            double tochkaconflictastart = 0;

            if (Program.VidRascheta == false)
            {
                int vsegoRaschetov = Program.TecushiyEcoIntens[1];
                
                workBook = ex.Workbooks.Add();
                ex.Iteration = true;


                sheetItog = (Excel.Worksheet)ex.Worksheets.get_Item(1);
                sheetItog.Name = "Результат";
                sheetItog.Cells[1, 1] = "Результирующая диаграмма";
                sheetItog.Cells[3, 1] = "№";
                sheetItog.Cells[3, 2] = "e";
                sheetItog.Cells[3, 3] = "r";
                sheetItog.Cells[3, 4] = "L";
                sheetItog.Cells[3, 5] = "N";
                sheetItog.Cells[3, 6] = "δ min";
                sheetItog.Cells[3, 7] = "δ max";
                sheetItog.Cells[3, 8] = "δ шаг";


                Excel.ChartObjects chartObjsItog = (Excel.ChartObjects)sheetItog.ChartObjects();
                Excel.ChartObject chartObjItog = chartObjsItog.Add(70, 300, 800, 400);
                Excel.Chart xlChartItog = chartObjItog.Chart;
                xlChartItog.HasTitle = true;
                xlChartItog.HasTitle = true;
                xlChartItog.ChartTitle.Text = "Расчет эко-интенсивности";
                xlChartItog.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).HasTitle = true;
                xlChartItog.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = "n";
                xlChartItog.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasTitle = true;
                xlChartItog.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = "E(n)";
                xlChartItog.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;
                xlChartItog.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasMinorGridlines = false;
                xlChartItog.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasMajorGridlines = false;
                xlChartItog.ChartType = Excel.XlChartType.xlXYScatterLinesNoMarkers;

                seriesCollectionItog = xlChartItog.SeriesCollection();
                int nomerlista = 0;

                for (int nomerrascheta = 0; nomerrascheta < vsegoRaschetov; nomerrascheta++)
                {

                    if (Program.SpisokEcoIntens[nomerrascheta, 3] != 0 && Program.SpisokEcoIntens[nomerrascheta, 1] != 0)
                    {

                        double ee = 0, r = 0, l = 0, n = 0, de = 0, demax = 0, deshag = 0;
                        ee = Program.SpisokEcoIntens[nomerrascheta, 0];
                        r = Program.SpisokEcoIntens[nomerrascheta, 2];
                        l = Program.SpisokEcoIntens[nomerrascheta, 1];
                        n = Program.SpisokEcoIntens[nomerrascheta, 4];
                        de = Program.SpisokEcoIntens[nomerrascheta, 3];
                        demax = Program.SpisokEcoIntens[nomerrascheta, 5];
                        deshag = Program.SpisokEcoIntens[nomerrascheta, 6];



                        //ex.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;
                        //ex.ScreenUpdating = true;
                        //ex.DisplayAlerts = true;
                        //ex.UserControl = true;
                        //ex.EnableEvents = true;
                        //ex.UserControl = true;
                        var xlSheets = ex.Sheets as Excel.Sheets;


                        sheetItog.Cells[4 + iitog, 1] = (iitog + 1).ToString();
                        sheetItog.Cells[4 + iitog, 2] = ee.ToString();
                        sheetItog.Cells[4 + iitog, 3] = r.ToString();
                        sheetItog.Cells[4 + iitog, 4] = l.ToString();
                        sheetItog.Cells[4 + iitog, 5] = n.ToString();


                        Excel.Range rng;
                        if (demax == 0 || deshag == 0)
                        {


                            sheetItog.Cells[4 + iitog, 6] = de.ToString();
                            sheetItog.Cells[3, 11 + (iraschet * 4)] = "n" + (iraschet + 1).ToString();
                            sheetItog.Cells[3, 12 + (iraschet * 4)] = "E" + (iraschet + 1).ToString() + " δ=" + de + " e=" + ee + " r=" + r + " L=" + l + " N=" + n;


                            for (int i = 0; i < n; i++)
                            {

                                sheetItog.Cells[i + 4, 11 + (iraschet * 4)] = i + 1;
                                ((Excel.Range)sheetItog.Cells[i + 4, 13 + (iraschet * 4)]).FormulaR1C1 = "=1-(1-R" + (4 + iitog).ToString() + "C6)^RC[-2]";
                                ((Excel.Range)sheetItog.Cells[i + 4, 13 + (iraschet * 4)]).Font.Color = Color.White;
                                ((Excel.Range)sheetItog.Cells[i + 4, 14 + (iraschet * 4)]).FormulaR1C1 = "=1/((1+R" + (4 + iitog).ToString() + "C3)^RC[-3])";
                                ((Excel.Range)sheetItog.Cells[i + 4, 14 + (iraschet * 4)]).Font.Color = Color.White;
                                ((Excel.Range)sheetItog.Cells[i + 4, 12 + (iraschet * 4)]).FormulaR1C1 = "= (R" + (4 + iitog).ToString() + "C2 / (R" + (4 + iitog).ToString() + "C6 * R" + (4 + iitog).ToString() + "C4)) * (SUM(R[-" + i.ToString() + "]C[1]:RC[1]) / SUM(R[-" + i.ToString() + "]C[2]:RC[2]))";
                                sheetItog.Cells[i + 4, 12 + (iraschet * 4)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                            }

                            



                            seriesItog = seriesCollectionItog.NewSeries();
                            var startCellitog = sheetItog.Cells[4, 11 + (iraschet * 4)];
                            var endCellitog = sheetItog.Cells[n + 3, 11 + (iraschet * 4)];
                            rng1Itog = sheetItog.Range[startCellitog, endCellitog];
                            seriesItog.XValues = rng1Itog;
                            startCellitog = sheetItog.Cells[4, 12 + (iraschet * 4)];
                            endCellitog = sheetItog.Cells[n + 3, 12 + (iraschet * 4)];
                            rng1Itog = sheetItog.Range[startCellitog, endCellitog];
                            seriesItog.Values = rng1Itog;
                            seriesItog.Name = "E" + (iraschet + 1).ToString() + " при δ=" + de + " e=" + ee + " r=" + r + " L=" + l + " N=" + n;

                            iitog++;
                            iraschet++;
                        }
                        else
                        {
                            
                            sheetItog.Cells[4 + iitog, 6] = de.ToString();
                            sheetItog.Cells[4 + iitog, 7] = demax.ToString();
                            sheetItog.Cells[4 + iitog, 8] = deshag.ToString();
                            sheetItog.Cells[3, 9] = "n";
                            sheetItog.Cells[3, 10] = "E(n)";
                            ((Excel.Range)sheetItog.Cells[3, 9]).Font.Color = Color.White;
                            ((Excel.Range)sheetItog.Cells[3, 10]).Font.Color = Color.White;


                            int j = 0;
                            double techDe = de;

                           

                            
                            while (techDe < demax)
                            {
                               
                                sheetItog.Cells[1, iraschet + 11] = techDe;
                                ((Excel.Range)sheetItog.Cells[4, iraschet + 11]).Font.Color = Color.White;
                                sheetItog.Cells[3, 11 + (iraschet * 4)] = "n" + (iraschet + 1).ToString();
                                sheetItog.Cells[3, 12 + (iraschet * 4)] = "E" + (iraschet + 1).ToString() + " при δ = " + techDe.ToString();


                                for (int i = 0; i < n; i++)
                                {
                                    
                                    sheetItog.Cells[i + 4, 11 + (iraschet * 4)] = i + 1;
                                    ((Excel.Range)sheetItog.Cells[i + 4, 13 + (iraschet * 4)]).FormulaR1C1 = "=1-(1-R1C" + (iraschet + 11).ToString() + ")^RC[-2]";
                                    ((Excel.Range)sheetItog.Cells[i + 4, 13 + (iraschet * 4)]).Font.Color = Color.White;
                                    ((Excel.Range)sheetItog.Cells[i + 4, 14 + (iraschet * 4)]).FormulaR1C1 = "=1/((1+R" + (4 + iitog).ToString() + "C3)^RC[-3])";
                                    ((Excel.Range)sheetItog.Cells[i + 4, 14 + (iraschet * 4)]).Font.Color = Color.White;
                                    ((Excel.Range)sheetItog.Cells[i + 4, 12 + (iraschet * 4)]).FormulaR1C1 = "= (R" + (4 + iitog).ToString() + "C2 / (R1C" + (iraschet + 11).ToString() + " * R" + (4 + iitog).ToString() + "C4)) * (SUM(R[-" + i.ToString() + "]C[1]:RC[1]) / SUM(R[-" + i.ToString() + "]C[2]:RC[2]))";
                                    sheetItog.Cells[i + 4, 12 + (iraschet * 4)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                }

                              
                                seriesItog = seriesCollectionItog.NewSeries();
                                rng1Itog = sheetItog.Range[sheetItog.Cells[4, 11 + (iraschet * 4)], sheetItog.Cells[n + 3, 11 + (iraschet * 4)]];
                                seriesItog.XValues = rng1Itog;
                                rng1Itog = sheetItog.Range[sheetItog.Cells[4, 12 + (iraschet * 4)], sheetItog.Cells[n + 3, 12 + (iraschet * 4)]];
                                seriesItog.Values = rng1Itog;
                                seriesItog.Name = "E" + (iraschet + 1).ToString() + " при δ=" + techDe + " e=" + ee + " r=" + r + " L=" + l + " N=" + n;



                                techDe = techDe + deshag;
                                j++;
                                iraschet++;
                            }

                           
                            sheetItog.Cells[1, iraschet + 11] = demax;
                            ((Excel.Range)sheetItog.Cells[4, iraschet + 11]).Font.Color = Color.White;
                            sheetItog.Cells[3, 11 + (iraschet * 4)] = "n" + (iraschet + 1).ToString();
                            sheetItog.Cells[3, 12 + (iraschet * 4)] = "E" + (iraschet + 1).ToString() + " при δ = " + demax.ToString();


                            for (int i = 0; i < n; i++)
                            {
                              
                                sheetItog.Cells[i + 4, 11 + (iraschet * 4)] = i + 1;
                                ((Excel.Range)sheetItog.Cells[i + 4, 13 + (iraschet * 4)]).FormulaR1C1 = "=1-(1-R1C" + (iraschet + 11).ToString() + ")^RC[-2]";
                                ((Excel.Range)sheetItog.Cells[i + 4, 13 + (iraschet * 4)]).Font.Color = Color.White;
                                ((Excel.Range)sheetItog.Cells[i + 4, 14 + (iraschet * 4)]).FormulaR1C1 = "=1/((1+R" + (4 + iitog).ToString() + "C3)^RC[-3])";
                                ((Excel.Range)sheetItog.Cells[i + 4, 14 + (iraschet * 4)]).Font.Color = Color.White;
                                ((Excel.Range)sheetItog.Cells[i + 4, 12 + (iraschet * 4)]).FormulaR1C1 = "= (R" + (4 + iitog).ToString() + "C2 / (R1C" + (iraschet + 11).ToString() + " * R" + (4 + iitog).ToString() + "C4)) * (SUM(R[-" + i.ToString() + "]C[1]:RC[1]) / SUM(R[-" + i.ToString() + "]C[2]:RC[2]))";
                                sheetItog.Cells[i + 4, 12 + (iraschet * 4)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


                            }

                            seriesItog = seriesCollectionItog.NewSeries();
                            rng1Itog = sheetItog.Range[sheetItog.Cells[4, 11 + (iraschet * 4)], sheetItog.Cells[n + 3, 11 + (iraschet * 4)]];
                            seriesItog.XValues = rng1Itog;
                            rng1Itog = sheetItog.Range[sheetItog.Cells[4, 12 + (iraschet * 4)], sheetItog.Cells[n + 3, 12 + (iraschet * 4)]];
                            seriesItog.Values = rng1Itog;
                            seriesItog.Name = "E" + (iraschet + 1).ToString() + " при δ=" + techDe + " e=" + ee + " r=" + r + " L=" + l + " N=" + n;

                            iitog++;
                            iraschet++;
                        };
                        sheetItog.Cells[1, 1] = " ";
                        nomerlista++;
                    };
                };





                nomerlista = 0;
                for (int nomerrascheta = 0; nomerrascheta < vsegoRaschetov; nomerrascheta++)
                {

                    if (Program.SpisokEcoIntens[nomerrascheta, 3] != 0 && Program.SpisokEcoIntens[nomerrascheta, 1] != 0)
                    {

                        double ee = 0, r = 0, l = 0, n = 0, de = 0, demax = 0, deshag = 0;
                        ee = Program.SpisokEcoIntens[nomerrascheta, 0];
                        r = Program.SpisokEcoIntens[nomerrascheta, 2];
                        l = Program.SpisokEcoIntens[nomerrascheta, 1];
                        n = Program.SpisokEcoIntens[nomerrascheta, 4];
                        de = Program.SpisokEcoIntens[nomerrascheta, 3];
                        demax = Program.SpisokEcoIntens[nomerrascheta, 5];
                        deshag = Program.SpisokEcoIntens[nomerrascheta, 6];



                        //ex.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;
                        //ex.ScreenUpdating = true;
                        //ex.DisplayAlerts = true;
                        //ex.UserControl = true;
                        //ex.EnableEvents = true;
                        //ex.UserControl = true;
                        var xlSheets = ex.Sheets as Excel.Sheets;

                        try
                        {
                            sheet = (Excel.Worksheet)ex.Worksheets.get_Item(nomerlista + 2);
                        }
                        catch
                        {
                            sheet = (Excel.Worksheet)xlSheets.Add(Type.Missing, xlSheets[nomerlista + 1], Type.Missing, Type.Missing);
                            sheet = (Excel.Worksheet)ex.Worksheets.get_Item(nomerlista + 2);

                        }


                        //sheet = (Excel.Worksheet)xlSheets.Add(Type.Missing, xlSheets[nomerrascheta + 1], Type.Missing, Type.Missing);
                        //MessageBox.Show(nomerrascheta.ToString());
                        sheet.Name = "Лист" + (nomerlista + 1).ToString();


                        sheet.Cells[1, 1] = "Расчет эко-интенсивности";



                        sheet.Cells[3, 1] = "e";
                        sheet.Cells[4, 1] = ee.ToString();
                        sheet.Cells[3, 2] = "r";
                        sheet.Cells[4, 2] = r.ToString();
                        sheet.Cells[3, 3] = "L";
                        sheet.Cells[4, 3] = l.ToString();
                        sheet.Cells[3, 4] = "N";
                        sheet.Cells[4, 4] = n.ToString();


                        Excel.Range rng;
                        if (demax == 0 || deshag == 0)
                        {
                            sheet.Cells[3, 5] = "δ";
                            sheet.Cells[4, 5] = de.ToString();

                            sheet.Cells[6, 1] = "n";
                            sheet.Cells[6, 2] = "E(n)";

                           

                            for (int i = 0; i < n; i++)
                            {
                                sheet.Cells[i + 7, 1] = i + 1;
                                ((Excel.Range)sheet.Cells[i + 7, 3]).FormulaR1C1 = "=1-(1-R4C5)^RC[-2]";
                                ((Excel.Range)sheet.Cells[i + 7, 3]).Font.Color = Color.White;
                                ((Excel.Range)sheet.Cells[i + 7, 4]).FormulaR1C1 = "=1/((1+R4C2)^RC[-3])";
                                ((Excel.Range)sheet.Cells[i + 7, 4]).Font.Color = Color.White;
                                ((Excel.Range)sheet.Cells[i + 7, 2]).FormulaR1C1 = "= (R4C1 / (R4C5 * R4C3)) * (SUM(R[-" + i.ToString() + "]C[1]:RC[1]) / SUM(R[-" + i.ToString() + "]C[2]:RC[2]))";
                                sheet.Cells[i + 7, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                               
                            }

                            Excel.ChartObjects chartObjs = (Excel.ChartObjects)sheet.ChartObjects();
                            Excel.ChartObject chartObj = chartObjs.Add(sheet.Cells[n + 7, 2].Left + 20, sheet.Cells[n + 7, 2].Top + 20, 600, 400);
                            Excel.Chart xlChart = chartObj.Chart;
                            xlChart.HasTitle = true;
                            xlChart.ChartTitle.Text = "Расчет эко-интенсивности";
                            xlChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).HasTitle = true;
                            xlChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = sheet.Cells[6, 1];
                            xlChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasTitle = true;
                            xlChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = sheet.Cells[6, 2];
                            xlChart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;
                            xlChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasMinorGridlines = false;
                            xlChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasMajorGridlines = false;
                            xlChart.ChartType = Excel.XlChartType.xlLine;

                            seriesCollection = xlChart.SeriesCollection();
                            series = seriesCollection.NewSeries();
                            var startCell = sheet.Cells[7, 1];
                            var endCell = sheet.Cells[n + 6, 1];
                            rng1 = sheet.Range[startCell, endCell];
                            series.XValues = rng1;
                            startCell = sheet.Cells[7, 2];
                            endCell = sheet.Cells[n + 6, 2];
                            rng1 = sheet.Range[startCell, endCell];
                            series.Values = rng1;
                            series.Name = "E при δ=" + de + " e=" + ee + " r=" + r + " L=" + l + " N=" + n;



                           

                            iitog++;
                            iraschet++;
                        }
                        else
                        {
                            sheet.Cells[3, 5] = "δ min";
                            sheet.Cells[4, 5] = de.ToString();
                            sheet.Cells[3, 6] = "δ max";
                            sheet.Cells[4, 6] = demax.ToString();
                            sheet.Cells[3, 7] = "δ шаг";
                            sheet.Cells[4, 7] = deshag.ToString();

                            sheet.Cells[3, 9] = "n";
                            sheet.Cells[3, 10] = "E(n)";

                           
                            ((Excel.Range)sheet.Cells[3, 9]).Font.Color = Color.White;
                            ((Excel.Range)sheet.Cells[3, 10]).Font.Color = Color.White;
                            int j = 0;
                            double techDe = de;

                            Excel.ChartObjects chartObjs = (Excel.ChartObjects)sheet.ChartObjects();
                            Excel.ChartObject chartObj = chartObjs.Add(sheet.Cells[n + 7, 2].Left + 20, sheet.Cells[n + 7, 2].Top + 20, 600, 400);
                            Excel.Chart xlChart = chartObj.Chart;
                            xlChart.HasTitle = true;
                            xlChart.ChartTitle.Text = "Расчет эко-интенсивности";
                            xlChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).HasTitle = true;
                            xlChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = sheet.Cells[3, 9];
                            xlChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasTitle = true;
                            xlChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = sheet.Cells[3, 10];
                            xlChart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;
                            xlChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasMinorGridlines = false;
                            xlChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasMajorGridlines = false;
                            xlChart.ChartType = Excel.XlChartType.xlLine;

                            seriesCollection = xlChart.SeriesCollection();

                            while (techDe < demax)
                            {
                                sheet.Cells[4, j + 9] = techDe;
                                ((Excel.Range)sheet.Cells[4, j + 9]).Font.Color = Color.White;
                                sheet.Cells[6, (j * 4) + 1] = "n" + (j + 1).ToString();
                                sheet.Cells[6, (j * 4) + 2] = "E" + (j + 1).ToString() + " при δ = " + techDe.ToString();

                                for (int i = 0; i < n; i++)
                                {
                                    sheet.Cells[i + 7, (j * 4) + 1] = i + 1;
                                    ((Excel.Range)sheet.Cells[i + 7, (j * 4) + 3]).FormulaR1C1 = "=1-(1-R4C" + (j + 9).ToString() + ")^RC[-2]";
                                    ((Excel.Range)sheet.Cells[i + 7, (j * 4) + 3]).Font.Color = Color.White;
                                    ((Excel.Range)sheet.Cells[i + 7, (j * 4) + 4]).FormulaR1C1 = "=1/((1+R4C2)^RC[-3])";
                                    ((Excel.Range)sheet.Cells[i + 7, (j * 4) + 4]).Font.Color = Color.White;
                                    ((Excel.Range)sheet.Cells[i + 7, (j * 4) + 2]).FormulaR1C1 = "= (R4C1 / (R4C" + (j + 9).ToString() + " * R4C3)) * (SUM(R[-" + i.ToString() + "]C[1]:RC[1]) / SUM(R[-" + i.ToString() + "]C[2]:RC[2]))";
                                    sheet.Cells[i + 7, (j * 4) + 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                }

                                series = seriesCollection.NewSeries();
                                rng1 = sheet.Range[sheet.Cells[7, (j * 4) + 1], sheet.Cells[n + 6, (j * 4) + 1]];
                                series.XValues = rng1;
                                rng1 = sheet.Range[sheet.Cells[7, (j * 4) + 2], sheet.Cells[n + 6, (j * 4) + 2]];
                                series.Values = rng1;
                                series.Name = "E" + (j + 1).ToString() + " при δ=" + techDe + " e=" + ee + " r=" + r + " L=" + l + " N=" + n;

                                techDe = techDe + deshag;
                                j++;
                                iraschet++;
                            }

                            sheet.Cells[4, j + 9] = demax;
                            ((Excel.Range)sheet.Cells[4, j + 9]).Font.Color = Color.White;
                            sheet.Cells[6, (j * 4) + 1] = "n" + (j + 1).ToString();
                            sheet.Cells[6, (j * 4) + 2] = "E" + (j + 1).ToString() + " при δ = " + demax.ToString();

                            for (int i = 0; i < n; i++)
                            {
                                sheet.Cells[i + 7, (j * 4) + 1] = i + 1;
                                ((Excel.Range)sheet.Cells[i + 7, (j * 4) + 3]).FormulaR1C1 = "=1-(1-R4C" + (j + 9).ToString() + ")^RC[-2]";
                                ((Excel.Range)sheet.Cells[i + 7, (j * 4) + 3]).Font.Color = Color.White;
                                ((Excel.Range)sheet.Cells[i + 7, (j * 4) + 4]).FormulaR1C1 = "=1/((1+R4C2)^RC[-3])";
                                ((Excel.Range)sheet.Cells[i + 7, (j * 4) + 4]).Font.Color = Color.White;
                                ((Excel.Range)sheet.Cells[i + 7, (j * 4) + 2]).FormulaR1C1 = "= (R4C1 / (R4C" + (j + 9).ToString() + " * R4C3)) * (SUM(R[-" + i.ToString() + "]C[1]:RC[1]) / SUM(R[-" + i.ToString() + "]C[2]:RC[2]))";
                                sheet.Cells[i + 7, (j * 4) + 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                            }

                            series = seriesCollection.NewSeries();
                            rng1 = sheet.Range[sheet.Cells[7, (j * 4) + 1], sheet.Cells[n + 6, (j * 4) + 1]];
                            series.XValues = rng1;
                            rng1 = sheet.Range[sheet.Cells[7, (j * 4) + 2], sheet.Cells[n + 6, (j * 4) + 2]];
                            series.Values = rng1;
                            series.Name = "E" + (iraschet + 1).ToString() + " при δ=" + demax + " e=" + ee + " r=" + r + " L=" + l + " N=" + n;

                          
                            iitog++;
                            iraschet++;
                        };
                        sheet.Cells[1, 1] = " ";
                        nomerlista++;
                    };
                };






            }
            else
            {
                int vsegoRaschetov = Program.TecushiyConflictInteres[1];
                int nomerlista = 0;

                workBook = ex.Workbooks.Add();
                ex.Iteration = true;


                sheetItog = (Excel.Worksheet)ex.Worksheets.get_Item(1);
                sheetItog.Name = "Результат";
                sheetItog.Cells[1, 1] = "Результирующая диаграмма";
                sheetItog.Cells[3, 1] = "№";
                sheetItog.Cells[3, 2] = "p";
                sheetItog.Cells[3, 3] = "c";
                sheetItog.Cells[3, 4] = "δ";
                sheetItog.Cells[3, 5] = "e";
                sheetItog.Cells[3, 6] = "d min";
                sheetItog.Cells[3, 7] = "d max";
                sheetItog.Cells[3, 8] = "d шаг";



                Excel.ChartObjects chartObjsItog = (Excel.ChartObjects)sheetItog.ChartObjects();
                Excel.ChartObject chartObjItog = chartObjsItog.Add(70, 300, 800, 400);
                Excel.Chart xlChartItog = chartObjItog.Chart;
                xlChartItog.HasTitle = true;
                xlChartItog.ChartTitle.Text = "Расчет конфликта интересов";
                xlChartItog.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).HasTitle = true;
                xlChartItog.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = "d";
                xlChartItog.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasTitle = true;
                xlChartItog.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = "β";
                xlChartItog.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;
                xlChartItog.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasMinorGridlines = false;
                xlChartItog.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasMajorGridlines = false;
                xlChartItog.ChartType = Excel.XlChartType.xlXYScatterLinesNoMarkers;
                seriesCollectionItog = xlChartItog.SeriesCollection();


                for (int nomerrascheta = 0; nomerrascheta < vsegoRaschetov; nomerrascheta++)
                {
                    double p = 0, c = 0, de = 0, d = 0, ee = 0, dmax = 0, dshag = 0;
                    p = Program.SpisokConflictInteres[nomerrascheta, 0];
                    c = Program.SpisokConflictInteres[nomerrascheta, 1];
                    de = Program.SpisokConflictInteres[nomerrascheta, 2];
                    d = Program.SpisokConflictInteres[nomerrascheta, 3];
                    ee = Program.SpisokConflictInteres[nomerrascheta, 4];
                    dmax = Program.SpisokConflictInteres[nomerrascheta, 5];
                    dshag = Program.SpisokConflictInteres[nomerrascheta, 6];

                    tochkaconflicta = (p - c) / ee;
                    tochkaconflictastart = (de * (p - c)) / ee;
                    if (tochkaconflictastart < tochkaconflicta && dshag != 0)
                    {

                        
                        sheetItog.Cells[4 + nomerlista, 1] = (iitog + 1).ToString();
                        sheetItog.Cells[4 + nomerlista, 2] = p.ToString();
                        sheetItog.Cells[4 + nomerlista, 3] = c.ToString();
                        sheetItog.Cells[4 + nomerlista, 4] = de.ToString();
                        sheetItog.Cells[4 + nomerlista, 5] = ee.ToString();
                        sheetItog.Cells[4 + nomerlista, 6] = d.ToString();
                        sheetItog.Cells[4 + nomerlista, 7] = dmax.ToString();
                        sheetItog.Cells[4 + nomerlista, 8] = dshag.ToString();
                        sheetItog.Cells[3, 11 + (nomerlista * 2)] = "d" + (iraschet + 1).ToString();
                        sheetItog.Cells[3, 12 + (nomerlista * 2)] = "β" + (iraschet + 1).ToString() + " p=" + p + " c=" + c + " δ=" + de + " e=" + ee + " d min" + d + " d max" + dmax + " d шаг" + dshag;


                        int n = 0;
                        double techDe = de;
                        double znachenietochkaconflicta = 0;

                        double znachenied = 0;
                        if (tochkaconflictastart > d)
                        {
                            znachenied = tochkaconflictastart + 0.0001;
                        }
                        else
                        {
                            znachenied = d;
                        };



                        if (tochkaconflicta < dmax)
                        {
                            znachenietochkaconflicta = tochkaconflicta;
                            if (((tochkaconflicta - znachenied) % dshag) == 0)
                            {
                                n = Convert.ToInt32(((tochkaconflicta - znachenied) / dshag) - (tochkaconflicta - znachenied) % dshag);
                            }
                            else
                            {
                                n = Convert.ToInt32(((tochkaconflicta - znachenied) / dshag) - (tochkaconflicta - znachenied) % dshag + 1);
                            }
                        }
                        int i = 0;
                        double techuchee = znachenied;
                        while (techuchee < znachenietochkaconflicta)
                        {
                           
                            sheetItog.Cells[4 + i, 11 + (nomerlista * 2)] = techuchee;
                            ((Excel.Range)sheetItog.Cells[4 + i, 12 + (nomerlista * 2)]).FormulaR1C1 = "=(LN(1-((R" + (4 + nomerlista).ToString() + "C4*(R" + (4 + nomerlista).ToString() + "C2-R" + (4 + nomerlista).ToString() + "C3))/(RC[-1]*R" + (4 + nomerlista).ToString() + "C5)))/LN(1-R" + (4 + nomerlista).ToString() + "C4))-1";

                            techuchee = techuchee + dshag;
                            i = i + 1;
                        };
                       
                        sheetItog.Cells[4 + i, 11 + (nomerlista * 2)] = znachenietochkaconflicta;
                        ((Excel.Range)sheetItog.Cells[4 + i, 12 + (nomerlista * 2)]).FormulaR1C1 = "=(LN(1-((R" + (4 + nomerlista).ToString() + "C4*(R" + (4 + nomerlista).ToString() + "C2-R" + (4 + nomerlista).ToString() + "C3))/(RC[-1]*R" + (4 + nomerlista).ToString() + "C5)))/LN(1-R" + (4 + nomerlista).ToString() + "C4))-1";

                       
                        seriesItog = seriesCollectionItog.NewSeries();
                        var startCellitog = sheetItog.Cells[4, 11 + (nomerlista * 2)];
                        var endCellitog = sheetItog.Cells[4 + i, 11 + (nomerlista * 2)];
                        rng1Itog = sheetItog.Range[startCellitog, endCellitog];
                        seriesItog.XValues = rng1Itog;
                        startCellitog = sheetItog.Cells[4, 12 + (nomerlista * 2)];
                        endCellitog = sheetItog.Cells[4 + i, 12 + (nomerlista * 2)];
                        rng1Itog = sheetItog.Range[startCellitog, endCellitog];
                        seriesItog.Values = rng1Itog;
                        seriesItog.Name = "β" + (nomerlista + 1).ToString() + " при p=" + p + " c=" + c + " δ=" + de + " e=" + ee + " d=" + d + ";" + (d + dshag).ToString() + "..." + znachenietochkaconflicta;
                        nomerlista++;

                    }
                };



                nomerlista = 0;

                for (int nomerrascheta = 0; nomerrascheta < vsegoRaschetov; nomerrascheta++)
                {
                    double p = 0, c = 0, de = 0, d = 0, ee = 0, dmax = 0, dshag = 0;
                    p = Program.SpisokConflictInteres[nomerrascheta, 0];
                    c = Program.SpisokConflictInteres[nomerrascheta, 1];
                    de = Program.SpisokConflictInteres[nomerrascheta, 2];
                    d = Program.SpisokConflictInteres[nomerrascheta, 3];
                    ee = Program.SpisokConflictInteres[nomerrascheta, 4];
                    dmax = Program.SpisokConflictInteres[nomerrascheta, 5];
                    dshag = Program.SpisokConflictInteres[nomerrascheta, 6];

                    tochkaconflicta = (p - c) / ee;
                    tochkaconflictastart = (de * (p - c)) / ee;
                    if (tochkaconflictastart < tochkaconflicta && dshag != 0)
                    {

                        var xlSheets = ex.Sheets as Excel.Sheets;
                        try
                        {
                            sheet = (Excel.Worksheet)ex.Worksheets.get_Item(nomerlista + 2);
                        }
                        catch
                        {
                            sheet = (Excel.Worksheet)xlSheets.Add(Type.Missing, xlSheets[nomerlista + 1], Type.Missing, Type.Missing);
                        }
                        sheet.Name = "Лист" + (nomerlista + 1).ToString();
                        //sheet = (Excel.Worksheet)ex.Worksheets.get_Item(nomerrascheta+1);
                        sheet.Cells[1, 1] = "Расчет конфликта интересов";
                        sheet.Cells[3, 1] = "p";
                        sheet.Cells[4, 1] = p.ToString();
                        sheet.Cells[3, 2] = "c";
                        sheet.Cells[4, 2] = c.ToString();
                        sheet.Cells[3, 3] = "δ";
                        sheet.Cells[4, 3] = de.ToString();
                        sheet.Cells[3, 4] = "e";
                        sheet.Cells[4, 4] = ee.ToString();
                        sheet.Cells[3, 5] = "d min";
                        sheet.Cells[4, 5] = d.ToString();
                        sheet.Cells[3, 6] = "d max";
                        sheet.Cells[4, 6] = dmax.ToString();
                        sheet.Cells[3, 7] = "d шаг";
                        sheet.Cells[4, 7] = dshag.ToString();

                        sheet.Cells[3, 9] = "Точка конфликта d";
                        ((Excel.Range)sheet.Cells[4, 9]).FormulaR1C1 = "=(RC[-8]-RC[-7])/RC[-5]";

                        sheet.Cells[6, 1] = "d";
                        sheet.Cells[6, 1] = "β(d)";

                       
                        int n = 0;
                        double techDe = de;
                        double znachenietochkaconflicta = 0;

                        double znachenied = 0;
                        if (tochkaconflictastart > d)
                        {
                            znachenied = tochkaconflictastart + 0.0001;
                        }
                        else
                        {
                            znachenied = d;
                        };



                        if (tochkaconflicta < dmax)
                        {
                            znachenietochkaconflicta = tochkaconflicta;
                            if (((tochkaconflicta - znachenied) % dshag) == 0)
                            {
                                n = Convert.ToInt32(((tochkaconflicta - znachenied) / dshag) - (tochkaconflicta - znachenied) % dshag);
                            }
                            else
                            {
                                n = Convert.ToInt32(((tochkaconflicta - znachenied) / dshag) - (tochkaconflicta - znachenied) % dshag + 1);
                            }
                        }
                        Excel.ChartObjects chartObjs = (Excel.ChartObjects)sheet.ChartObjects();
                        Excel.ChartObject chartObj = chartObjs.Add(sheet.Cells[n + 7, 2].Left + 20, sheet.Cells[n + 7, 2].Top + 20, 600, 400);
                        Excel.Chart xlChart = chartObj.Chart;
                        xlChart.HasTitle = true;
                        xlChart.ChartTitle.Text = "Расчет конфликта интересов";
                        xlChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).HasTitle = true;
                        xlChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = "d";
                        xlChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasTitle = true;
                        xlChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = "β";
                        xlChart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;
                        xlChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasMinorGridlines = false;
                        xlChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).HasMajorGridlines = false;
                        xlChart.ChartType = Excel.XlChartType.xlXYScatterLinesNoMarkers;
                        seriesCollection = xlChart.SeriesCollection();
                        int i = 0;
                        double techuchee = znachenied;
                        while (techuchee < znachenietochkaconflicta)
                        {
                            sheet.Cells[i + 7, 1] = techuchee;
                            ((Excel.Range)sheet.Cells[i + 7, 2]).FormulaR1C1 = "=(LN(1-((R4C3*(R4C1-R4C2))/(RC[-1]*R4C4)))/LN(1-R4C3))-1";
                            sheet.Cells[i + 7, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                            techuchee = techuchee + dshag;
                            i = i + 1;
                        };
                        sheet.Cells[i + 7, 1] = znachenietochkaconflicta;
                        ((Excel.Range)sheet.Cells[i + 7, 2]).FormulaR1C1 = "=(LN(1-((R4C3*(R4C1-R4C2))/(RC[-1]*R4C4)))/LN(1-R4C3))-1";

                        sheet.Cells[i + 7, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        series = seriesCollection.NewSeries();
                        rng1 = sheet.Range[sheet.Cells[7, 1], sheet.Cells[i + 7, 1]];
                        series.XValues = rng1;
                        rng1 = sheet.Range[sheet.Cells[7, 2], sheet.Cells[i + 7, 2]];
                        series.Values = rng1;
                        series.Name = "β при p=" + p + " c=" + c + " δ=" + de + " e=" + ee + " d=" + d + ";" + (d + dshag).ToString() + "..." + znachenietochkaconflicta;

                        nomerlista++;

                    }
                };
                
                
            };
            



            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel File(.xlsx)|*.xlsx";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.RestoreDirectory = true;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fileName = saveFileDialog1.FileName;
                workBook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                saveFileDialog1.Dispose();
                ex.Visible = true;
            }


            //ex.Quit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (Program.VidRascheta == false)
            {
                Program.RaschetEcoIntensiveForm.Show();
            }
            else
            {
                Program.ConflictInteresovForm.Show();
            }

            this.Close();
        }

        private void RaschetItog_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Program.VidRascheta == false)
            {
                Program.RaschetEcoIntensiveForm.Show();
            }
            else
            {
                Program.ConflictInteresovForm.Show();
            }
        }
    }
}
