using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;

namespace EcoIntensive
{
    public partial class ConflictInteresov : Form
    {
        object NavedenieObiect;

        private double[,] massivvichislenii = new double[2, 1000];
        public string cc;
        int tablestrok=0;
        double tochkaconflicta = 0;
        double tochkaconflictastart = 0;
        public ConflictInteresov()
        {
            InitializeComponent();

        }

        private void RaschetConflictInteresovToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Program.RaschetEcoIntensiveForm == null)
            { Program.RaschetEcoIntensiveForm = new RaschetEcoIntensive(); }
            Program.RaschetEcoIntensiveForm.Show();
            Program.VidRascheta = false;
            this.Hide();
        }

        private void VichislenieFormuli(double p, double c , double de, double d, double e, double dmax, double dshag)
        {
            
            tochkaconflictastart = (de*(p-c))/e;
            tochkaconflicta = (p - c) / e;
            double znachenied = 0;
            if (tochkaconflictastart>d)
            {
                znachenied = tochkaconflictastart+0.0001;
            }
            else
            {
                znachenied = d;
            };
            
            int j = 0;
            while (znachenied < dmax && znachenied < tochkaconflicta)
            {
                massivvichislenii[0, j] = znachenied;
                massivvichislenii[1, j] = (Math.Log(1-(de*(p-c))/(znachenied*e)) / Math.Log(1-de)) - 1;
                znachenied = znachenied + dshag;
                j++;
            }
            tablestrok = j;
            if (tochkaconflicta.ToString() == d.ToString())
            { tablestrok = j-1; }
            if (tochkaconflicta.ToString() != d.ToString())
            {
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
        }

        private void RaschetConflictInteresovtoolStrip_MouseUp(object sender, MouseEventArgs e)
        {
            DublirovatToolStripMenuItem.Enabled = false;
            YdalitToolStripMenuItem.Enabled = false;
            SochranitToolStripMenuItem3.Enabled = false;
            if (e.Button == MouseButtons.Right)
            {
                contextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right);
            }
        }

        public void DobavitRaschetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string kkk;
           ToolStripButton Btn = new ToolStripButton();
            Btn.Text = "Расчет конфликта интересов №" + (Program.TecushiyConflictInteres[2] + 1);
            Btn.AutoToolTip = false;
            // устанавливаем обработчик нажатия
            Btn.MouseUp += (sender1, args) =>
            {
                DublirovatToolStripMenuItem.Enabled = true;
                YdalitToolStripMenuItem.Enabled = true;
                SochranitToolStripMenuItem3.Enabled = true;
                NavedenieObiect = sender1;
            };
            Btn.Click += (sender1, args) =>
            {
                LabelTochki.Text = "";
                for (int i = 0; i < Program.TecushiyConflictInteres[1]; i++)
                {
                    if (Btn.Name.ToString() == Program.NameSpisokConflictInteres[i])
                    {
                        Program.TecushiyConflictInteres[0] = i;
                    }
                }
                kkk = Btn.Name.ToString();
                kkk = kkk.Replace("RaschetConflictInteresov", "");
                labelZagolovok.Text = "Расчет конфликта интересов №" + kkk;
                textBoxp.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 0].ToString();
                textBoxc.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 1].ToString();
                textBoxde.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 2].ToString();
                textBoxdmin.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 3].ToString();
                textBoxe.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 4].ToString();
                textBoxdmax.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 5].ToString();
                textBoxdshag.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 6].ToString();
                if (Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 7] == 1)
                {
                    Raschet_Click(sender, e);
                }

                //Btn.BackColor = SystemColors.MenuHighlight;
            };

            Btn.Name = "RaschetConflictInteresov" + (Program.TecushiyConflictInteres[2] + 1);
            kkk = Btn.Name.ToString();
            kkk = kkk.Replace("RaschetConflictInteresov", "");
            Program.NameSpisokConflictInteres[Program.TecushiyConflictInteres[1]] = Btn.Name;
            labelZagolovok.Text = "Расчет конфликта интересов №" + kkk;
            Program.TecushiyConflictInteres[0] = Program.TecushiyConflictInteres[1];
            Program.TecushiyConflictInteres[1] = Program.TecushiyConflictInteres[1] + 1;
            Program.TecushiyConflictInteres[2] = Program.TecushiyConflictInteres[2] + 1;
            RaschetConflictInterestoolStrip.Items.Add(Btn);
            OchistitDannieFormi();
            panel1.Visible = true;
            TablicaDannihRascheta.Visible = true;




        }

        private void YdalitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int NomerNaidenoiStroki = -1;
            ToolStripButton buttontech = NavedenieObiect as ToolStripButton;

            if (buttontech != null)
            {
                for (int i = 0; i <= Program.TecushiyConflictInteres[1]; i++)
                {
                    if (NomerNaidenoiStroki == -1)
                        if (buttontech.Name.ToString() == Program.NameSpisokConflictInteres[i])
                        {
                            NomerNaidenoiStroki = i;

                            break;
                        };
                }
                for (int i = NomerNaidenoiStroki; i < Program.TecushiyConflictInteres[1]; i++)
                {

                    Program.NameSpisokConflictInteres[i] = Program.NameSpisokConflictInteres[i + 1];
                    for (int j = 0; j < 20; j++)
                    {
                        Program.SpisokConflictInteres[i, j] = Program.SpisokConflictInteres[i + 1, j];
                    }
                }
                RaschetConflictInterestoolStrip.Items.Remove(NavedenieObiect as ToolStripButton);
                OchistitDannieFormi();
                Program.TecushiyConflictInteres[1] = Program.TecushiyConflictInteres[1] - 1;
            }
        }

        private void RaschetConflictInteresov_Load(object sender, EventArgs e)
        {
            panel1.Visible = false;
            TablicaDannihRascheta.Visible = false;
        }


        private void FuncTextBoxChanged(object sender, EventArgs e)
        {
            //string znachenie = textBoxe.Text.ToString();
            TextBox textboxelem = sender as TextBox;
            int i = -1;
            switch (textboxelem.Name)
            {
                case "textBoxp":
                    i = 0;
                    break;
                case "textBoxc":
                    i = 1;
                    break;
                case "textBoxde":
                    i = 2;
                    break;
                case "textBoxdmin":
                    i = 3;
                    break;
                case "textBoxe":
                    i = 4;
                    break;
                case "textBoxdmax":
                    i = 5;
                    break;
                case "textBoxdshag":
                    i = 6;
                    break;
            }
            if (textboxelem.Text == "" || textboxelem.Text == ",")
            {
                textboxelem.Text = "0";
            }

            if (Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], i] != double.Parse(textboxelem.Text))
            {
                Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], i] = double.Parse(textboxelem.Text); ;
                Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 7] = 0;
            }
            LabelTochki.Text = "";
            OchistitGraphic();

        }

        private void Raschet_Click(object sender, EventArgs e)
        {
            Random rnd = new Random();
            int i = 0;
            int r, g, b;

            OchistitGraphic();

            double p = 0, c = 0, de = 0, d = 0, ee = 0, dmax = 0, dshag = 0;
            p = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 0];
            c = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 1];
            de = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 2];
            d = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 3];
            ee = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 4];
            dmax = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 5];
            dshag = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 6];

            double tochkaconflictastart;
            double tochkaconflicta; 

            tochkaconflictastart = (de * (p - c)) / ee;
            tochkaconflicta = (p - c) / ee;

            if (tochkaconflictastart < tochkaconflicta && dshag != 0 && (((tochkaconflictastart <= dmax) && (tochkaconflictastart >= d)) || ((tochkaconflicta <= dmax) && (tochkaconflicta >= d))))
            {
                VichislenieFormuli(Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 0], Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 1], Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 2], Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 3], Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 4], Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 5], Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 6]);

                LabelTochki.Text = "";
                if (tochkaconflictastart != 0) LabelTochki.Text = LabelTochki.Text + "При d<" + tochkaconflictastart.ToString() + " момент конфликта интересов не наступает.";
                if (tochkaconflictastart != 0) LabelTochki.Text = LabelTochki.Text + "При d>" + tochkaconflicta.ToString() + " конфликт интересов наступает моментально.";

                TablicaDannihRascheta.Columns.Add("d" + (i + 1).ToString(), "d" + (i + 1).ToString());
                TablicaDannihRascheta.Columns.Add("be" + (i + 1).ToString(), "β" + (i + 1).ToString() + "(d" + (i + 1).ToString() + ")");

                r = rnd.Next(0, 255);
                g = rnd.Next(0, 255);
                b = rnd.Next(0, 255);
                chart1.Series.Add("Series" + (i).ToString());
                chart1.Series["Series" + i.ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                chart1.Series["Series" + i.ToString()].Color = Color.FromArgb(255, r, g, b);

                chart1.Legends.Add("Legends0");
                chart1.Legends["Legends0"].Docking = Docking.Left;
                chart1.Legends["Legends0"].IsDockedInsideChartArea = false;
                chart1.Legends["Legends0"].TableStyle = LegendTableStyle.Wide;
                chart1.Legends["Legends0"].Alignment = StringAlignment.Center;
                chart1.Series["Series" + i.ToString()].Legend = "Legends0";
                chart1.Series["Series" + i.ToString()].BorderWidth = 2;
                chart1.Series["Series" + i.ToString()].LegendText = "β при p=" + Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 0].ToString() + " c=" + Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 1].ToString() + " δ=" + Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 2].ToString() + " e=" + Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 4].ToString() + " d=" + Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 3].ToString() + ";" + (Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 3] + Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 6]).ToString() + "..." + Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 5].ToString();

                //chart1.Legends["Legends" + i.ToString()].Alignment = StringAlignment.Center;
                //chart1.Legends["Legends" + i.ToString()].Docking = Docking.Bottom;


                for (int j = 0; j <= tablestrok; j++)
                {
                    TablicaDannihRascheta.Rows.Add();
                    TablicaDannihRascheta.Rows[j].Cells[0].Value = Math.Round(massivvichislenii[0, j], 3);
                    TablicaDannihRascheta.Rows[j].Cells[1].Value = Math.Round(massivvichislenii[1, j], 3);

                    if (tochkaconflictastart != 0)
                        chart1.Series["Series" + i.ToString()].Points.AddXY(Math.Round(massivvichislenii[0, j], 3), Math.Round(massivvichislenii[1, j], 3));
                    else
                        chart1.Series["Series" + i.ToString()].Points.AddXY(Math.Round(massivvichislenii[0, j], 3), Math.Round(massivvichislenii[1, j], 3));
                }
                Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 7] = 1;


            }
        }

        private void SochranitToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (textBoxp.Text == "0" || textBoxc.Text == "0" || textBoxde.Text == "0" || textBoxdmin.Text == "0" || textBoxe.Text == "0" || textBoxdmax.Text == "0" || textBoxdshag.Text == "0" || textBoxp.Text == "" || textBoxc.Text == "" || textBoxde.Text == "" || textBoxdmin.Text == "" || textBoxe.Text == "" || textBoxdmax.Text == "" || textBoxdshag.Text == "")
                MessageBox.Show("Необходимо заполнить все поля");
            else
            {
                Program.VidRascheta = true;
                OknoSochranenia oknoSochranenia = new OknoSochranenia();
                oknoSochranenia.Show();
                
            }
           
        }

        private void Funct_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox textboxelem = sender as TextBox;
            string textboxstring = textboxelem.Text;
            int ColVoZapiatich = 0;
            int ColVoPeredZapiatich = 0;

            if (e.KeyChar == 46) e.KeyChar = ',';
            if (e.KeyChar == 44 && textboxstring.Length == 1 && textboxelem.SelectionLength == textboxelem.Text.Length) e.Handled = true; ;
            for (int i = 0; i < textboxstring.Length; i++)
            {
                if (textboxstring[i] == ',') ColVoZapiatich++;
                if (ColVoZapiatich > 0) ColVoPeredZapiatich++;
            }
            if (ColVoPeredZapiatich > 3) e.Handled = true;
            if (textboxelem.SelectionLength > 0) e.Handled = false;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 44 && e.KeyChar != 08 && e.KeyChar != 127 && e.KeyChar != 00 || (ColVoZapiatich == 1 && e.KeyChar == 44))
                e.Handled = true;


            if (e.KeyChar == 08 || e.KeyChar == 127) e.Handled = false;
        }
       
        private void OchistitDannieFormi()
        {
            textBoxp.Text = "0";
            textBoxc.Text = "0";
            textBoxde.Text = "0";
            textBoxdmin.Text = "0";
            textBoxe.Text = "0";
            textBoxdmax.Text = "0";
            textBoxdshag.Text = "0";
            LabelTochki.Text = "";
            OchistitGraphic();
        }

        private void OchistitGraphic()
        {
            chart1.Series.Clear();
            chart1.Legends.Clear();
            TablicaDannihRascheta.Rows.Clear();
            TablicaDannihRascheta.Columns.Clear();

        }

        private void DublirovatToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string kkk;
            int NomerNaidenoiStroki = -1;
            ToolStripButton buttontech = NavedenieObiect as ToolStripButton;


            for (int i = 0; i <= Program.TecushiyConflictInteres[1]; i++)
            {
                if (NomerNaidenoiStroki == -1)
                    if (buttontech.Name.ToString() == Program.NameSpisokConflictInteres[i])
                    {
                        NomerNaidenoiStroki = i;

                        break;
                    };
            }

            ToolStripButton Btn = new ToolStripButton();
            Btn.Text = "Расчет конфликта интересов №" + (Program.TecushiyConflictInteres[2] + 1);
            Btn.AutoToolTip = false;
            // устанавливаем обработчик нажатия
            Btn.MouseUp += (sender1, args) =>
            {
                DublirovatToolStripMenuItem.Enabled = true;
                YdalitToolStripMenuItem.Enabled = true;
                SochranitToolStripMenuItem3.Enabled = true;
                NavedenieObiect = sender1;
            };
            Btn.Click += (sender1, args) =>
            {
                LabelTochki.Text = "";
                for (int i = 0; i < Program.TecushiyConflictInteres[1]; i++)
                {
                    if (Btn.Name.ToString() == Program.NameSpisokConflictInteres[i])
                    {
                        Program.TecushiyConflictInteres[0] = i;
                    }
                }
                kkk = Btn.Name.ToString();
                kkk = kkk.Replace("RaschetConflictInteresov", "");
                labelZagolovok.Text = "Расчет конфликта интересов №" + kkk;
                textBoxp.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 0].ToString();
                textBoxc.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 1].ToString();
                textBoxde.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 2].ToString();
                textBoxdmin.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 3].ToString();
                textBoxe.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 4].ToString();
                textBoxdmax.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 5].ToString();
                textBoxdshag.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 6].ToString();
                if (Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 7] == 1)
                {
                    Raschet_Click(sender, e);
                }
                
                //Btn.BackColor = SystemColors.MenuHighlight;
            };

            Btn.Name = "RaschetConflictInteresov" + (Program.TecushiyConflictInteres[2] + 1);
            kkk = Btn.Name.ToString();
            kkk = kkk.Replace("RaschetConflictInteresov", "");
            Program.NameSpisokConflictInteres[Program.TecushiyConflictInteres[1]] = Btn.Name;
            labelZagolovok.Text = "Расчет конфликта интересов №" + kkk;
            Program.TecushiyConflictInteres[0] = Program.TecushiyConflictInteres[1];
            Program.TecushiyConflictInteres[1] = Program.TecushiyConflictInteres[1] + 1;
            Program.TecushiyConflictInteres[2] = Program.TecushiyConflictInteres[2] + 1;
            RaschetConflictInterestoolStrip.Items.Add(Btn);
            OchistitDannieFormi();
            panel1.Visible = true;
            TablicaDannihRascheta.Visible = true;

            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 0] = Program.SpisokConflictInteres[NomerNaidenoiStroki, 0];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 1] = Program.SpisokConflictInteres[NomerNaidenoiStroki, 1];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 2] = Program.SpisokConflictInteres[NomerNaidenoiStroki, 2];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 3] = Program.SpisokConflictInteres[NomerNaidenoiStroki, 3];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 4] = Program.SpisokConflictInteres[NomerNaidenoiStroki, 4];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 5] = Program.SpisokConflictInteres[NomerNaidenoiStroki, 5];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 6] = Program.SpisokConflictInteres[NomerNaidenoiStroki, 6];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 7] = Program.SpisokConflictInteres[NomerNaidenoiStroki, 7];
            textBoxp.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 0].ToString();
            textBoxc.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 1].ToString();
            textBoxde.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 2].ToString();
            textBoxdmin.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 3].ToString();
            textBoxe.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 4].ToString();
            textBoxdmax.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 5].ToString();
            textBoxdshag.Text = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 6].ToString();
           if (Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 7] == 1)
            {
                Raschet_Click(sender, e);
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Program.VidRascheta = true;
            RaschetItog newForm = new RaschetItog();
            newForm.Show();
            this.Hide();
        }

        private void ExportExcelToolStripMenuItem_Click_1(object sender, EventArgs e)
        {

            double p = 0, c = 0, de = 0, d = 0, ee = 0, dmax = 0, dshag = 0;
            p = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 0];
            c = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 1];
            de = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 2];
            d = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 3];
            ee = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 4];
            dmax = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 5];
            dshag = Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 6];
            double tochkaconflictastart;
            double tochkaconflicta;

            tochkaconflictastart = (de * (p - c)) / ee;
            tochkaconflicta = (p - c) / ee;

            if (tochkaconflictastart < tochkaconflicta && dshag != 0 && (((tochkaconflictastart <= dmax) && (tochkaconflictastart >= d)) || ((tochkaconflicta <= dmax) && (tochkaconflicta >= d))))
            {
                Excel.Application ex = new Excel.Application();
                Excel.Workbook workBook;
                Excel.Worksheet sheet;
                Excel.SeriesCollection seriesCollection;
                Excel.Series series;
                Excel.Range rng1;

                workBook = ex.Workbooks.Add();
                ex.Iteration = true;

                sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
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

                sheet.Cells[3, 10] = "max d когда конфликт наступает моментально";
                ((Excel.Range)sheet.Cells[4, 10]).FormulaR1C1 = "=(RC[-9]-RC[-8])/RC[-6]";

                sheet.Cells[3, 9] = "min d когда можно найти момент конфликта";
                ((Excel.Range)sheet.Cells[4, 9]).FormulaR1C1 = "=(RC[-6]*(RC[-8]-RC[-7]))/RC[-5]";

                sheet.Cells[6, 1] = "d";
                sheet.Cells[6, 2] = "β(d)";
                int n = 0;
                double techDe = de;
                tochkaconflicta = (p - c) / ee;
                tochkaconflictastart = (de * (p - c)) / ee;
                double znachenied = 0;
                if (tochkaconflictastart > d)
                {
                    znachenied = tochkaconflictastart + 0.0001;
                }
                else
                {
                    znachenied = d;
                };
                double znachenietochkaconflicta = 0;
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
                else
                {
                    znachenietochkaconflicta = dmax;
                    if (((dmax - znachenied) % dshag) == 0)
                    {
                        n = Convert.ToInt32(((dmax - znachenied) / dshag) - (dmax - znachenied) % dshag);
                    }
                    else
                    {
                        n = Convert.ToInt32(((dmax - znachenied) / dshag) - (dmax - znachenied) % dshag + 1);
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
                xlChart.ChartType = Excel.XlChartType.xlLine;
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
            }
        }

        private void ConflictInteresov_FormClosing(object sender, FormClosingEventArgs e)
        {
            Program.SpisokEcoIntens = null;
            Program.NameSpisokEcoIntens = null;
            Program.SpisokConflictInteres = null;
            Program.NameSpisokConflictInteres = null;
            Program.TecushiyEcoIntens = null;
            Program.TecushiyConflictInteres = null;
            Program.RaschetEcoIntensiveForm = null;
            Program.ConflictInteresovForm = null;
            Application.Exit();
        }

        private void StartEcoIntensive_Click(object sender, EventArgs e)
        {
            panel2.Visible = false;
            RaschetConflictInterestoolStrip.Visible = true;
            RaschetItogButton.Visible = true;
            label15.Visible = true;
            Program.sender11 = sender;
            Program.e11 = e;
        }

        private void FuncTextBoxClick(object sender, EventArgs e)
        {
            TextBox textboxelem = sender as TextBox;
            if (textboxelem.Text == "0")
            {
                textboxelem.SelectionStart = 0;
                textboxelem.SelectionLength = textboxelem.Text.Length;
            }
        }

        private void FuncTextBoxDoublelick(object sender, EventArgs e)
        {
            TextBox textboxelem = sender as TextBox;
            if (!String.IsNullOrEmpty(textboxelem.Text))
            {
                textboxelem.SelectionStart = 0;
                textboxelem.SelectionLength = textboxelem.Text.Length;
            }
        }

        private void FuncTextBoxFocusLeave(object sender, EventArgs e)
        {
            TextBox textboxelem = sender as TextBox;
            if (textboxelem.Text == "")
            {
                textboxelem.Text = "0";
            }
        }

        private void textBoxdmax_Leave(object sender, EventArgs e)
        {
            FuncTextBoxFocusLeave(sender, e);
            TextBox textboxelem = sender as TextBox;
            if (Convert.ToDouble(textBoxdmax.Text) < Convert.ToDouble(textBoxdmin.Text))
            {
                textBoxdmax.Text = textBoxdmin.Text;
            }
        }

        private void textBoxdshag_Leave(object sender, EventArgs e)
        {
            FuncTextBoxFocusLeave(sender, e);
            TextBox textboxelem = sender as TextBox;
            if (Convert.ToDouble(textBoxdshag.Text) == 0 && Convert.ToDouble(textBoxdmax.Text) != 0)
            {
                textBoxdshag.Text = "0,1";
            }
        }

        private void OtcritIzBDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Program.VidRascheta = true;
            DobavlenieIzBD newForm = new DobavlenieIzBD();
            newForm.Show();
            this.Hide();
        }
        public void DobavitRaschetDB(object sender, EventArgs e)
        {
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[1] - 1, 0] = Program.znacheniastroki[0];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[1] - 1, 1] = Program.znacheniastroki[1];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[1] - 1, 2] = Program.znacheniastroki[2];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[1] - 1, 3] = Program.znacheniastroki[3];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[1] - 1, 4] = Program.znacheniastroki[4];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[1] - 1, 5] = Program.znacheniastroki[5];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[1] - 1, 6] = Program.znacheniastroki[6];
            Program.SpisokConflictInteres[Program.TecushiyConflictInteres[1] - 1, 7] = Program.znacheniastroki[7];
            textBoxp.Text = Program.znacheniastroki[0].ToString();
            textBoxc.Text = Program.znacheniastroki[1].ToString();
            textBoxde.Text = Program.znacheniastroki[2].ToString();
            textBoxdmin.Text = Program.znacheniastroki[3].ToString();
            textBoxe.Text = Program.znacheniastroki[4].ToString();
            textBoxdmax.Text = Program.znacheniastroki[5].ToString();
            textBoxdshag.Text = Program.znacheniastroki[6].ToString();
            
            if (Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 7] == 1)
            {
                Raschet_Click(sender, e);
            }
        }
    }
}
