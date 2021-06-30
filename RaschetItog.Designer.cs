
namespace EcoIntensive
{
    partial class RaschetItog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RaschetItog));
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.label12 = new System.Windows.Forms.Label();
            this.TableItog = new System.Windows.Forms.DataGridView();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.label1 = new System.Windows.Forms.Label();
            this.label25 = new System.Windows.Forms.Label();
            this.ExportExcel = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.TableItog)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.SuspendLayout();
            // 
            // imageList1
            // 
            this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Calibri", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label12.Location = new System.Drawing.Point(417, 9);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(329, 24);
            this.label12.TabIndex = 69;
            this.label12.Text = "Итоговый расчет эко-интенсивности ";
            // 
            // TableItog
            // 
            this.TableItog.AllowUserToAddRows = false;
            this.TableItog.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.TableItog.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TableItog.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.TableItog.EnableHeadersVisualStyles = false;
            this.TableItog.Location = new System.Drawing.Point(48, 100);
            this.TableItog.Name = "TableItog";
            this.TableItog.RowHeadersVisible = false;
            this.TableItog.Size = new System.Drawing.Size(346, 558);
            this.TableItog.TabIndex = 102;
            // 
            // chart1
            // 
            this.chart1.BorderlineColor = System.Drawing.Color.DimGray;
            this.chart1.BorderlineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Solid;
            chartArea1.AxisX.Title = "n";
            chartArea1.AxisY.Title = "E(n)";
            chartArea1.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea1);
            this.chart1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            legend1.Alignment = System.Drawing.StringAlignment.Center;
            legend1.Docking = System.Windows.Forms.DataVisualization.Charting.Docking.Bottom;
            legend1.Name = "Legend1";
            legend1.TitleAlignment = System.Drawing.StringAlignment.Far;
            this.chart1.Legends.Add(legend1);
            this.chart1.Location = new System.Drawing.Point(431, 100);
            this.chart1.Name = "chart1";
            this.chart1.Size = new System.Drawing.Size(893, 545);
            this.chart1.TabIndex = 103;
            this.chart1.Text = "chart1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(101, 78);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(219, 19);
            this.label1.TabIndex = 104;
            this.label1.Text = "Массив полученных значений";
            this.label1.Visible = false;
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Font = new System.Drawing.Font("Calibri", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label25.Location = new System.Drawing.Point(733, 73);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(311, 24);
            this.label25.TabIndex = 105;
            this.label25.Text = "График расчета эко-интенсивности";
            // 
            // ExportExcel
            // 
            this.ExportExcel.BackColor = System.Drawing.Color.SteelBlue;
            this.ExportExcel.FlatAppearance.BorderSize = 0;
            this.ExportExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ExportExcel.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ExportExcel.ForeColor = System.Drawing.Color.White;
            this.ExportExcel.Location = new System.Drawing.Point(1043, 651);
            this.ExportExcel.Name = "ExportExcel";
            this.ExportExcel.Size = new System.Drawing.Size(281, 30);
            this.ExportExcel.TabIndex = 106;
            this.ExportExcel.Text = "Экспортировать в Excel";
            this.ExportExcel.UseVisualStyleBackColor = false;
            this.ExportExcel.Click += new System.EventHandler(this.ExportExcel_Click);
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.SteelBlue;
            this.button3.FlatAppearance.BorderSize = 0;
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.Font = new System.Drawing.Font("Calibri", 12F);
            this.button3.ForeColor = System.Drawing.Color.White;
            this.button3.Location = new System.Drawing.Point(1165, 12);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(173, 30);
            this.button3.TabIndex = 107;
            this.button3.Text = "Вернуться к расчетам";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // RaschetItog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(1350, 689);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.ExportExcel);
            this.Controls.Add(this.label25);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.chart1);
            this.Controls.Add(this.TableItog);
            this.Controls.Add(this.label12);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "RaschetItog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Расчет эко-интенсивности";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.RaschetItog_FormClosed);
            this.Load += new System.EventHandler(this.RaschetItog_Load);
            ((System.ComponentModel.ISupportInitialize)(this.TableItog)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.DataGridView TableItog;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.Button ExportExcel;
        private System.Windows.Forms.Button button3;
    }
}