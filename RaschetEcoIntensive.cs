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
    public partial class RaschetEcoIntensive : Form
    {
        object NavedenieObiect;

        private double[,] massivvichislenii = new double[1000, 10000];
        bool boolinterval = false;
        public string cc;
        Random rnd = new Random();
        public RaschetEcoIntensive()
        {
            InitializeComponent();
            Program.TecushiyEcoIntens[0] = 0;
            Program.TecushiyEcoIntens[1] = 0;
            Program.TecushiyConflictInteres[0] = 0;
            Program.TecushiyConflictInteres[1] = 0;
            //RaschetEcoIntenstoolStrip.Visible = false;
            DublirovatToolStripMenuItem.Enabled = false;
            YdalitToolStripMenuItem.Enabled = false;
            SochranitToolStripMenuItem3.Enabled = false;
        }

        private void VichislenieFormuli(double e, double l , double r, double de, double n, int i)
        {
            for (int j = 0; j < n; j++)
            {
                double ZnachenieChislitel = 0;
                double ZnachenieZnamenatel = 0;
                for (int k = 1; k <= j+1; k++)
                {
                    ZnachenieChislitel = ZnachenieChislitel + (1 - Math.Pow((1 - de), k));
                    ZnachenieZnamenatel = ZnachenieZnamenatel + (1 / (Math.Pow((1 + r), k)));
                }
                massivvichislenii[i, j] = (e / (de * l)) * (ZnachenieChislitel / ZnachenieZnamenatel);
            }
            
        }

        //Контекстное еню списка правая кнопка мыши
        private void RaschetEcoIntenstoolStrip_MouseUp(object sender, MouseEventArgs e)
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
            Btn.Text = "Расчет эко-интенсивности №" + (Program.TecushiyEcoIntens[2] + 1);
            Btn.AutoToolTip = false;
           // устанавливаем обработчик нажатия
            Btn.MouseUp += (sender1, args) =>
            {
                NavedenieObiect = sender1;
                DublirovatToolStripMenuItem.Enabled = true;
                YdalitToolStripMenuItem.Enabled = true;
                SochranitToolStripMenuItem3.Enabled = true;
            };
            Btn.Click += (sender1, args) =>
            {
                
                for (int i = 0; i < Program.TecushiyEcoIntens[1]; i++)
                {
                    if (Btn.Name.ToString() == Program.NameSpisokEcoIntens[i])
                    {
                        Program.TecushiyEcoIntens[0] = i;
                    }
                }
                kkk= Btn.Name.ToString();
                kkk = kkk.Replace("RaschetEcointensStroka", "");
                labelZagolovok.Text = "Расчет эко-интенсивности №" + kkk;
                textBoxe.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 0].ToString();
                textBoxl.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 1].ToString();
                textBoxr.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 2].ToString();
                textBoxde.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 3].ToString();
                textBoxn.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4].ToString();
                textBoxdemax.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 5].ToString();
                textBoxdeshag.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 6].ToString();
                if (textBoxdemax.Text == "0" || textBoxdeshag.Text == "0" || textBoxdeshag.Text == "" || textBoxdemax.Text == "" )
                    VidDeStart();
                else
                {
                    if (boolinterval == false)
                    {
                        VidDeRaschrit();
                    };
                }
                if (Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 7] == 1)
                {
                    Raschet_Click(sender, e);
                };
                
                //Btn.BackColor = SystemColors.MenuHighlight;
            };

            Btn.Name = "RaschetEcointensStroka" + (Program.TecushiyEcoIntens[2] + 1);
            kkk = Btn.Name.ToString();
            kkk = kkk.Replace("RaschetEcointensStroka", "");
            Program.NameSpisokEcoIntens[Program.TecushiyEcoIntens[1]] = Btn.Name;
            labelZagolovok.Text = "Расчет эко-интенсивности №" + kkk;
            Program.TecushiyEcoIntens[0] = Program.TecushiyEcoIntens[1];
            Program.TecushiyEcoIntens[1] = Program.TecushiyEcoIntens[1] + 1;
            RaschetEcoIntenstoolStrip.Items.Add(Btn);
            OchistitDannieFormi();
            panel1.Visible = true;
            TablicaDannihRascheta.Visible = true;
            VidDeStart();
            Program.TecushiyEcoIntens[2] = Program.TecushiyEcoIntens[2] + 1;


        }

        private void YdalitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int NomerNaidenoiStroki = -1;
            ToolStripButton buttontech = NavedenieObiect as ToolStripButton;

            if (buttontech != null)
            {
                for (int i = 0; i <= Program.TecushiyEcoIntens[1]; i++)
                {
                    if (NomerNaidenoiStroki == -1)
                        if (buttontech.Name.ToString() == Program.NameSpisokEcoIntens[i])
                        {
                            NomerNaidenoiStroki = i;

                            break;
                        };
                }
                for (int i = NomerNaidenoiStroki; i < Program.TecushiyEcoIntens[1] - 1; i++)
                {

                    Program.NameSpisokEcoIntens[i] = Program.NameSpisokEcoIntens[i + 1];
                    for (int j = 0; j < 20; j++)
                    {
                        label1.Text = j.ToString();
                        Program.SpisokEcoIntens[i, j] = Program.SpisokEcoIntens[i + 1, j];
                    }
                }
                RaschetEcoIntenstoolStrip.Items.Remove(NavedenieObiect as ToolStripButton);
                OchistitDannieFormi();
                Program.TecushiyEcoIntens[1] = Program.TecushiyEcoIntens[1] - 1;
            }
        }

        private void RaschetEcoIntensive_Load(object sender, EventArgs e)
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
                case "textBoxe":
                    i = 0;
                    break;
                case "textBoxl":
                    i = 1;
                    break;
                case "textBoxr":
                    i = 2;
                    break;
                case "textBoxde":
                    i = 3;
                    break;
                case "textBoxn":
                    i = 4;
                    break;
                case "textBoxdemax":
                    i = 5;
                    break;
                case "textBoxdeshag":
                    i = 6;
                    break;
            }
            if (textboxelem.Text == "" || textboxelem.Text == ",")
            {
                textboxelem.Text = "0";
            }
            if (Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], i] != double.Parse(textboxelem.Text))
            {
                
                Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], i] = double.Parse(textboxelem.Text);
                Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 7] = 0;
               

            }
            OchistitGraphic();

        }

        public void Raschet_Click(object sender, EventArgs e)
        {

            int i = 0;
            int r, g, b;
            if (textBoxde.Text != "0" && textBoxl.Text != "0")
            { 
                OchistitGraphic();

                if (textBoxdemax.Text == "0" || textBoxdeshag.Text == "0")
                {
                    VichislenieFormuli(Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 0], Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 1], Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 2], Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 3], Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4], i);

                    TablicaDannihRascheta.Columns.Add("n" + (i + 1).ToString(), "n" + (i + 1).ToString());
                    TablicaDannihRascheta.Columns.Add("E" + (i + 1).ToString(), "E" + (i + 1).ToString());

                    r = rnd.Next(0, 255);
                    g = rnd.Next(0, 255);
                    b = rnd.Next(0, 255);
                    chart1.Series.Add("Series" + (i).ToString());
                    chart1.Series["Series" + i.ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                    chart1.Series["Series" + i.ToString()].Color = Color.FromArgb(255, r, g, b);


                    chart1.Legends.Add("Legends0");
                    chart1.Legends["Legends0"].Docking = Docking.Left;
                    chart1.Legends["Legends0"].IsDockedInsideChartArea = false;
                    //chart1.Legends["Legends" + (i).ToString()].TableStyle = LegendTableStyle.Wide;
                    chart1.Legends["Legends0"].Alignment = StringAlignment.Center;
                    chart1.Legends["Legends0"].LegendStyle = LegendStyle.Column;
                    chart1.Series["Series" + i.ToString()].Legend = "Legends0";
                    chart1.Series["Series" + i.ToString()].BorderWidth = 2;
                    chart1.Series["Series" + i.ToString()].LegendText = "E" + (i + 1).ToString() + " при δ=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 3].ToString() + " e=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 0].ToString() + " r=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 2].ToString() + " L=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 1].ToString() + " N=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4].ToString();

                    for (int j = 0; j < Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4]; j++)
                    {
                        //if (j < Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4] - 1) TablicaDannihRascheta.Rows.Add();
                        TablicaDannihRascheta.Rows.Add();
                        TablicaDannihRascheta.Rows[j].Cells[0].Value = (j + 1).ToString();
                        TablicaDannihRascheta.Rows[j].Cells[1].Value = String.Format("{0:#0.0##}", massivvichislenii[0, j]);

                        chart1.Series["Series" + i.ToString()].Points.AddXY(j + 1, massivvichislenii[0, j]);
                        //if (Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4] - j == 0) TablicaDannihRascheta.Rows.RemoveAt(j-1);
                    }
                    //TablicaDannihRascheta.Rows[int.Parse((Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4] - 1).ToString())].Cells[0].Value = (int.Parse((Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4] - 1).ToString()) + 1).ToString();
                    //TablicaDannihRascheta.Rows[int.Parse((Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4] - 1).ToString())].Cells[1].Value = massivvichislenii[0, int.Parse((Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4] - 1).ToString())];

                    //chart1.Series["Series" + i.ToString()].Points.AddXY(int.Parse((Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4] - 1).ToString()) + 1, massivvichislenii[0, int.Parse((Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4] - 1).ToString())]);

                }
                else
                {
                    chart1.Legends.Add("Legends0");
                    chart1.Legends["Legends0"].Docking = Docking.Left;
                    chart1.Legends["Legends0"].Alignment = StringAlignment.Center;
                    chart1.Legends["Legends0"].LegendStyle = LegendStyle.Column;
                    double techDe = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 3];
                    while (techDe < Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 5])
                    {

                        VichislenieFormuli(Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 0], Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 1], Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 2], techDe, Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4], i);

                        TablicaDannihRascheta.Columns.Add("n" + (i + 1).ToString(), "n" + (i + 1).ToString());
                        TablicaDannihRascheta.Columns.Add("E" + (i + 1).ToString(), "E" + (i + 1).ToString());

                        r = rnd.Next(0, 255);
                        g = rnd.Next(0, 255);
                        b = rnd.Next(0, 255);
                        chart1.Series.Add("Series" + (i).ToString());
                        chart1.Series["Series" + i.ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                        chart1.Series["Series" + i.ToString()].Color = Color.FromArgb(255, r, g, b);

                        chart1.Series["Series" + i.ToString()].Legend = "Legends0";
                        chart1.Series["Series" + i.ToString()].BorderWidth = 2;
                        chart1.Series["Series" + i.ToString()].LegendText = "E" + (i + 1).ToString() + " при δ=" + techDe + " e=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 0].ToString() + " r=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 2].ToString() + " L=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 1].ToString() + " N=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4].ToString();

                        if (TablicaDannihRascheta.Rows.Count < Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4])
                        {
                            while (Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4] - TablicaDannihRascheta.Rows.Count > 0)
                            {
                                TablicaDannihRascheta.Rows.Add();
                            };
                        }

                        for (int j = 0; j < Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4]; j++)
                        {

                            TablicaDannihRascheta.Rows[j].Cells[2 * i].Value = (j + 1).ToString();
                            TablicaDannihRascheta.Rows[j].Cells[2 * i + 1].Value = String.Format("{0:#0.0##}", massivvichislenii[i, j]);

                            chart1.Series["Series" + i.ToString()].Points.AddXY(j + 1, massivvichislenii[i, j]);
                        };
                        techDe = techDe + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 6];
                        i++;
                    }
                    VichislenieFormuli(Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 0], Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 1], Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 2], Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 5], Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4], i);

                    TablicaDannihRascheta.Columns.Add("n" + (i + 1).ToString(), "n" + (i + 1).ToString());
                    TablicaDannihRascheta.Columns.Add("E" + (i + 1).ToString(), "E" + (i + 1).ToString());

                    r = rnd.Next(0, 255);
                    g = rnd.Next(0, 255);
                    b = rnd.Next(0, 255);
                    chart1.Series.Add("Series" + (i).ToString());
                    chart1.Series["Series" + i.ToString()].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
                    chart1.Series["Series" + i.ToString()].Color = Color.FromArgb(255, r, g, b);
                    chart1.Series["Series" + i.ToString()].Legend = "Legends0";
                    chart1.Series["Series" + i.ToString()].BorderWidth = 2;
                    chart1.Series["Series" + i.ToString()].LegendText = "E" + (i + 1).ToString() + " при δ=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 5].ToString() + " e=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 0].ToString() + " r=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 2].ToString() + " L=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 1].ToString() + " N=" + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4].ToString();

                    if (TablicaDannihRascheta.Rows.Count < Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4])
                    {
                        while (Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4] - TablicaDannihRascheta.Rows.Count > 0)
                        {
                            TablicaDannihRascheta.Rows.Add();
                        };
                    }
                    for (int j = 0; j < Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4]; j++)
                    {

                        TablicaDannihRascheta.Rows[j].Cells[2 * i].Value = (j + 1).ToString();
                        TablicaDannihRascheta.Rows[j].Cells[2 * i + 1].Value = String.Format("{0:#0.0##}", massivvichislenii[i, j]);

                        chart1.Series["Series" + i.ToString()].Points.AddXY(j + 1, massivvichislenii[i, j]);
                    }
                    techDe = techDe + Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 6];

                };
                Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 7] = 1;
            }
            else MessageBox.Show("Показатели δ и l не могут быть равны нулю");

        }

        private void SochranitToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (textBoxe.Text == "0" || textBoxl.Text == "0" ||  textBoxde.Text == "0" || textBoxn.Text == "0" || textBoxe.Text == "" || textBoxl.Text == "" ||  textBoxde.Text == "" || textBoxn.Text == "")
                MessageBox.Show("Необходимо заполнить все поля");
            else
            {
                Program.VidRascheta = false;
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
            

            if (e.KeyChar == 08 || e.KeyChar == 127)  e.Handled = false;
            //if (textboxelem.Text.Length>6 && textboxelem.SelectionLength == 0) 
            //    e.Handled = true;
        }


        private void textBoxn_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox textboxelem = sender as TextBox;
            string textboxstring = textboxelem.Text;
            int ColVoSimvolov = 0;
            for (int i = 0; i < textboxstring.Length; i++)
            {
                ColVoSimvolov++;
            }
            if (ColVoSimvolov > 2) e.Handled = true;
            if (textboxelem.SelectionLength > 0) e.Handled = false;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 08 && e.KeyChar != 127 && e.KeyChar != 00)
                e.Handled = true;
            
            if (e.KeyChar == 08 || e.KeyChar == 127) e.Handled = false;

        }

        private void ShagButton_Click(object sender, EventArgs e)
        {
            if (boolinterval == false)
            {
                VidDeRaschrit();
            }
            else
            {
                VidDeStart();

            }


        }

        public void VidDeStart ()
        {
            label10.Text = "δ =";
            label13.Visible = false;
            label14.Visible = false;
            textBoxdemax.Visible = false;
            textBoxdeshag.Visible = false;
            textBoxdemax.Text = "0";
            textBoxdeshag.Text = "0";
            if (textBoxdemax.Text == "") Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 5] = double.Parse("0");
            else Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 5] = double.Parse(textBoxdemax.Text);
            if (textBoxdeshag.Text == "") Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 6] = double.Parse("0");
            else Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 6] = double.Parse(textBoxdeshag.Text);
            textBoxde.Location = new Point(49, 183);
            label5.Location = new Point(137, 186);
            label10.Text = "δ = ";
            ShagButton.Location = new Point(535, 183);
            boolinterval = false;
            ShagButton.Text = "Задать δ интервалом с шагом";
            textBoxn.TabIndex = 5;
            Raschet.TabIndex = 6;
            ShagButton.TabIndex = 7;
        }

        public void VidDeRaschrit()
        {
            label10.Text = "δ min =";
            label13.Text = "δ max =";
            label13.Visible = true;
            label14.Text = "δ шаг =";
            label14.Visible = true;
            textBoxdemax.Visible = true;
            textBoxdeshag.Visible = true;
            label5.Location = new Point(458, 186);
            textBoxde.Location = new Point(73, 183);
            ShagButton.Location = new Point(856, 183);
            boolinterval = true;
            ShagButton.Text = "Задать δ константой";
            textBoxdemax.TabIndex = 5;
            textBoxdeshag.TabIndex = 6;
            textBoxn.TabIndex = 7;
            Raschet.TabIndex = 8;
            ShagButton.TabIndex = 9;
        }

        public void OchistitDannieFormi()
        {
            textBoxe.Text = "0";
            textBoxl.Text = "0";
            textBoxr.Text = "0";
            textBoxde.Text = "0";
            textBoxn.Text = "0";
            textBoxdemax.Text = "0";
            textBoxdeshag.Text = "0";
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

            if (buttontech != null)
            {
                for (int i = 0; i <= Program.TecushiyEcoIntens[1]; i++)
                {
                    if (NomerNaidenoiStroki == -1)
                        if (buttontech.Name.ToString() == Program.NameSpisokEcoIntens[i])
                        {
                            NomerNaidenoiStroki = i;

                            break;
                        };
                }

                ToolStripButton Btn = new ToolStripButton();
                Btn.Text = "Расчет эко-интенсивности №" + (Program.TecushiyEcoIntens[2] + 1);
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
                    DublirovatToolStripMenuItem.Enabled = true;
                    YdalitToolStripMenuItem.Enabled = true;
                    SochranitToolStripMenuItem3.Enabled = true;
                    for (int i = 0; i < Program.TecushiyEcoIntens[1]; i++)
                    {
                        if (Btn.Name.ToString() == Program.NameSpisokEcoIntens[i])
                        {
                            Program.TecushiyEcoIntens[0] = i;
                        }
                    }
                    kkk = Btn.Name.ToString();
                    kkk = kkk.Replace("RaschetEcointensStroka", "");
                    labelZagolovok.Text = "Расчет эко-интенсивности №" + kkk;
                    textBoxe.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 0].ToString();
                    textBoxl.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 1].ToString();
                    textBoxr.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 2].ToString();
                    textBoxde.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 3].ToString();
                    textBoxn.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4].ToString();
                    textBoxdemax.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 5].ToString();
                    textBoxdeshag.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 6].ToString();
                    if (textBoxdemax.Text == "0" || textBoxdeshag.Text == "0" || textBoxdeshag.Text == "" || textBoxdemax.Text == "")
                        VidDeStart();
                    else
                    {
                        if (boolinterval == false)
                        {
                            VidDeRaschrit();
                        };
                    }
                    if (Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 7] == 1)
                    {
                        Raschet_Click(sender, e);
                    }

                    //Btn.BackColor = SystemColors.MenuHighlight;
                };

                Btn.Name = "RaschetEcointensStroka" + (Program.TecushiyEcoIntens[2] + 1);
                kkk = Btn.Name.ToString();
                kkk = kkk.Replace("RaschetEcointensStroka", "");
                Program.NameSpisokEcoIntens[Program.TecushiyEcoIntens[1]] = Btn.Name;
                labelZagolovok.Text = "Расчет эко-интенсивности №" + kkk;
                Program.TecushiyEcoIntens[0] = Program.TecushiyEcoIntens[1];
                Program.TecushiyEcoIntens[1] = Program.TecushiyEcoIntens[1] + 1;
                Program.TecushiyEcoIntens[2] = Program.TecushiyEcoIntens[2] + 1;
                RaschetEcoIntenstoolStrip.Items.Add(Btn);
                panel1.Visible = true;
                TablicaDannihRascheta.Visible = true;

                Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 0] = Program.SpisokEcoIntens[NomerNaidenoiStroki, 0];
                Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 1] = Program.SpisokEcoIntens[NomerNaidenoiStroki, 1];
                Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 2] = Program.SpisokEcoIntens[NomerNaidenoiStroki, 2];
                Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 3] = Program.SpisokEcoIntens[NomerNaidenoiStroki, 3];
                Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4] = Program.SpisokEcoIntens[NomerNaidenoiStroki, 4];
                Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 5] = Program.SpisokEcoIntens[NomerNaidenoiStroki, 5];
                Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 6] = Program.SpisokEcoIntens[NomerNaidenoiStroki, 6];
                Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 7] = Program.SpisokEcoIntens[NomerNaidenoiStroki, 7];
                textBoxe.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 0].ToString();
                textBoxl.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 1].ToString();
                textBoxr.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 2].ToString();
                textBoxde.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 3].ToString();
                textBoxn.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4].ToString();
                textBoxdemax.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 5].ToString();
                textBoxdeshag.Text = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 6].ToString();
                if (textBoxdemax.Text == "0" || textBoxdeshag.Text == "0" || textBoxdeshag.Text == "" || textBoxdemax.Text == "")
                    VidDeStart();
                else
                {
                    if (boolinterval == false)
                    {
                        VidDeRaschrit();
                    };
                }
                if (Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 7] == 1)
                {
                    Raschet_Click(sender, e);
                }
            }
           

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Program.VidRascheta = false;
            RaschetItog newForm = new RaschetItog();
            newForm.Show();
            this.Hide();
        }
        
        //Экспорт в excel
        private void ExportExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            double ee = 0, r = 0, l = 0, n = 0, de = 0, demax = 0, deshag = 0;

            
            ee = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 0];
            r = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 2];
            l = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 1];
            n = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4];
            de = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 3];
            demax = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 5];
            deshag = Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 6];
            if (de != 0 && l != 0)
            {
                Excel.Application ex = new Excel.Application();
                Excel.Workbook workBook;
                Excel.Worksheet sheet;
                Excel.SeriesCollection seriesCollection;
                Excel.Series series;
                Excel.Range rng1;


                workBook = ex.Workbooks.Add();

                ex.Iteration = true;
                //ex.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;
                //ex.ScreenUpdating = true;
                //ex.DisplayAlerts = true;
                //ex.UserControl = true;
                //ex.EnableEvents = true;
                //ex.UserControl = true;
                sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);


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
                        sheet.Cells[6, (j * 4) + 2] = "E" + (j + 1).ToString() + "(n" + (j + 1).ToString() + ") при δ = " + techDe.ToString();


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
                        series.Name = "E при δ=" + techDe + " e=" + ee + " r=" + r + " L=" + l + " N=" + n;

                        techDe = techDe + deshag;
                        j++;
                    }

                    sheet.Cells[4, j + 9] = demax;
                    ((Excel.Range)sheet.Cells[4, j + 9]).Font.Color = Color.White;
                    sheet.Cells[6, (j * 4) + 1] = "n" + (j + 1).ToString();
                    sheet.Cells[6, (j * 4) + 2] = "E" + (j + 1).ToString() + "(n" + (j + 1).ToString() + ") при δ = " + demax.ToString();


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
                    series.Name = "E при δ=" + demax + " e=" + ee + " r=" + r + " L=" + l + " N=" + n;
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
            };


        }

        private void OtcritIzBDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DobavlenieIzBD newForm = new DobavlenieIzBD();
            newForm.Show();
            this.Hide();
        }

        private void конфликтИнтересовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Program.ConflictInteresovForm == null)
            { Program.ConflictInteresovForm = new ConflictInteresov(); }
            Program.ConflictInteresovForm.Show();
            Program.VidRascheta = true;
            this.Hide();

        }

        private void RaschetEcoIntensive_FormClosing(object sender, FormClosingEventArgs e)
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
            RaschetEcoIntenstoolStrip.Visible = true;
            RaschetItogButton.Visible = true;
            label12.Visible = true;
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

        private void  FuncTextBoxFocusLeave(object sender, EventArgs e)
        {
            TextBox textboxelem = sender as TextBox;
            if (textboxelem.Text == "")
            {
                textboxelem.Text = "0";
            }
            if (textboxelem.Text == ",")
            {
                textboxelem.Text = "0";
            }
        }
        
        private void textBoxdemax_Leave(object sender, EventArgs e)
        {
            FuncTextBoxFocusLeave(sender, e);
            TextBox textboxelem = sender as TextBox;
            if (Convert.ToDouble(textBoxdemax.Text)< Convert.ToDouble(textBoxde.Text))
            {
                textBoxdemax.Text = textBoxde.Text;
            }
        }

        private void textBoxdeshag_Leave(object sender, EventArgs e)
        {
            FuncTextBoxFocusLeave(sender, e);
            TextBox textboxelem = sender as TextBox;
            if (Convert.ToDouble(textBoxdeshag.Text) == 0 && Convert.ToDouble(textBoxdemax.Text) != 0)
            {
                textBoxdeshag.Text = "0,1";
            }
        }


        public void DobavitRaschetDB(object sender, EventArgs e)
        {
            Program.SpisokEcoIntens[Program.TecushiyEcoIntens[1] - 1, 0] = Program.znacheniastroki[0];
            Program.SpisokEcoIntens[Program.TecushiyEcoIntens[1] - 1, 1] = Program.znacheniastroki[1];
            Program.SpisokEcoIntens[Program.TecushiyEcoIntens[1] - 1, 2] = Program.znacheniastroki[2];
            Program.SpisokEcoIntens[Program.TecushiyEcoIntens[1] - 1, 3] = Program.znacheniastroki[3];
            Program.SpisokEcoIntens[Program.TecushiyEcoIntens[1] - 1, 4] = Program.znacheniastroki[4];
            Program.SpisokEcoIntens[Program.TecushiyEcoIntens[1] - 1, 5] = Program.znacheniastroki[5];
            Program.SpisokEcoIntens[Program.TecushiyEcoIntens[1] - 1, 6] = Program.znacheniastroki[6];
            Program.SpisokEcoIntens[Program.TecushiyEcoIntens[1] - 1, 7] = Program.znacheniastroki[7];
            textBoxe.Text = Program.znacheniastroki[0].ToString();
            textBoxl.Text = Program.znacheniastroki[1].ToString();
            textBoxr.Text = Program.znacheniastroki[2].ToString();
            textBoxde.Text = Program.znacheniastroki[3].ToString();
            textBoxn.Text = Program.znacheniastroki[4].ToString();
            textBoxdemax.Text = Program.znacheniastroki[5].ToString();
            textBoxdeshag.Text = Program.znacheniastroki[6].ToString();
            if (textBoxdemax.Text == "0" || textBoxdeshag.Text == "0" || textBoxdeshag.Text == "" || textBoxdemax.Text == "")
                VidDeStart();
            else
            {
                if (boolinterval == false)
                {
                    VidDeRaschrit();
                };
            }
            if (Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 7] == 1)
            {
                Raschet_Click(sender, e);
            }
        }

        
    }
}
