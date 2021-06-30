using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using System.IO;

namespace EcoIntensive
{
    public partial class DobavlenieIzBD : Form
    {


        SQLiteCommand Command;


        public DobavlenieIzBD()
        {
            InitializeComponent();
        }

        private void DobavlenieIzBD_Load(object sender, EventArgs e)
        {
            TablicaDannihRascheta.Columns.Add("project", "Проект");
            TablicaDannihRascheta.Columns.Add("location", "Локация");
            TablicaDannihRascheta.Columns.Add("Polzovateli", "Пользователь");
            TablicaDannihRascheta.Columns.Add("datasozd", "Дата");
            if (Program.VidRascheta == false)
            {
                TablicaDannihRascheta.Columns.Add("e", "e");
                TablicaDannihRascheta.Columns.Add("l", "l");
                TablicaDannihRascheta.Columns.Add("r", "r");
                TablicaDannihRascheta.Columns.Add("de", "δ");
                TablicaDannihRascheta.Columns.Add("n", "n");
                TablicaDannihRascheta.Columns.Add("demax", "δ max");
                TablicaDannihRascheta.Columns.Add("demin", "δ шаг");
            }
            else
            {
                TablicaDannihRascheta.Columns.Add("p", "p");
                TablicaDannihRascheta.Columns.Add("c", "c");
                TablicaDannihRascheta.Columns.Add("de", "δ");
                TablicaDannihRascheta.Columns.Add("d", "d");
                TablicaDannihRascheta.Columns.Add("e", "e");
                TablicaDannihRascheta.Columns.Add("dmax", "d max");
                TablicaDannihRascheta.Columns.Add("dmin", "d шаг");
            }
            SQLiteConnection Connect = null;
            Connect.Open();
            //Заполнение комбобокса пользователя
            Command = new SQLiteCommand
            {
                Connection = Connect,
                CommandText = @"select fio from Polzovateli"
            };

            SQLiteDataReader sqlReader = Command.ExecuteReader();
            while (sqlReader.Read())
            {
                comboBoxPolzovatel.Items.Add(sqlReader["fio"].ToString());
            }

            //Заполнение комбобокса локации 
            Command = new SQLiteCommand
            {
                Connection = Connect,
                CommandText = @"select location from Locations"
            };
            sqlReader = Command.ExecuteReader();
            while (sqlReader.Read())
            {
                comboBoxLocacia.Items.Add(sqlReader["location"].ToString());
            }

            //Заполнение комбобокса Проекта 
            Command = new SQLiteCommand
            {
                Connection = Connect,
                CommandText = @"select project from Projects"
            };
            sqlReader = Command.ExecuteReader();
            while (sqlReader.Read())
            {
                comboBoxProect.Items.Add(sqlReader["project"].ToString());
            }
            Connect.Close();
            buttonFind_Click(sender, e);
        }

        private void buttonFind_Click(object sender, EventArgs e)
        {

            int idpolzovatel = -1;
            int idproect = -1;
            int idlocatia = -1;
            string commandText = "";

            SQLiteConnection Connect = null;
            Connect.Open();

            if (comboBoxPolzovatel.Text != "")
            {
                Command = new SQLiteCommand { Connection = Connect,CommandText = @"select id, fio from Polzovateli where fio like @perem" };
                Command.Parameters.AddWithValue("@perem", comboBoxPolzovatel.Text);
                SQLiteDataReader sqlReader = Command.ExecuteReader();
                while (sqlReader.Read())
                {
                    idpolzovatel = Convert.ToInt32(sqlReader["id"]);
                };
            };
            if (comboBoxProect.Text != "")
            {
                Command = new SQLiteCommand { Connection = Connect, CommandText = @"select id, project from Projects where project like @perem" };
                Command.Parameters.AddWithValue("@perem", comboBoxProect.Text);
                SQLiteDataReader sqlReader = Command.ExecuteReader();
                while (sqlReader.Read())
                {
                    idproect = Convert.ToInt32(sqlReader["id"]);
                };
               
            };
            if (comboBoxLocacia.Text != "")
            {
                Command = new SQLiteCommand { Connection = Connect, CommandText = @"select id, location from Locations where location like @perem"};
                Command.Parameters.AddWithValue("@perem", comboBoxLocacia.Text);
                SQLiteDataReader sqlReader = Command.ExecuteReader();
                while (sqlReader.Read())
                {
                    idlocatia = Convert.ToInt32(sqlReader["id"]);
                };
            };
            if (Program.VidRascheta == false)
            {
                bool indicator = false;
                commandText = commandText + @"SELECT raschetEcoIntensive.id,Projects.project, Locations.location, Polzovateli.fio, date,e,l, r, de, n,demax,deshag FROM raschetEcoIntensive 
                                            LEFT JOIN Polzovateli ON raschetEcoIntensive.polzovatel = Polzovateli.id
                                            LEFT JOIN Locations ON raschetEcoIntensive.location = Locations.id
                                            LEFT JOIN Projects ON raschetEcoIntensive.project = Projects.id";

                if (comboBoxPolzovatel.Text != "" || comboBoxLocacia.Text != "" || comboBoxProect.Text != "" || dateStart.Checked == true || dateEnd.Checked == true)
                {
                    commandText = commandText + @" where ";
                };

                if (comboBoxPolzovatel.Text != "")
                {
                    commandText = commandText + @"raschetEcoIntensive.polzovatel like @idpolzovatel";
                    indicator = true;
                }
                if (comboBoxProect.Text != "")
                {
                    if (indicator == true)
                    {
                        commandText = commandText + " and ";
                    }
                    commandText = commandText + @"raschetEcoIntensive.project like @idproect";
                    indicator = true; 
                }

                if (comboBoxLocacia.Text != "")
                {
                    if (indicator == true)
                    {
                        commandText = commandText + " and ";
                    }
                    commandText = commandText + @"raschetEcoIntensive.location like @idlocatia";
                    indicator = true;
                }
                if (dateStart.Checked == true)
                {
                    if (indicator == true)
                    {
                        commandText = commandText + " and ";
                    }
                    commandText = commandText + @"raschetEcoIntensive.date >= @dateStart";
                    indicator = true;
                }
                if (dateEnd.Checked == true)
                {
                    if (indicator == true)
                    {
                        commandText = commandText + " and ";
                    }
                    commandText = commandText + @"raschetEcoIntensive.date <= @dateEnd";
                    indicator = true;
                }

                //if (comboBoxPolzovatel.Text != "")
                //{
                //    commandText = commandText + @"raschetEcoIntensive.polzovatel like @idpolzovatel";
                //    //Command.Parameters.AddWithValue("@idpolzovatel", idpolzovatel);
                //    if (comboBoxProect.Text != "")
                //    {
                //        commandText = commandText + @" and raschetEcoIntensive.project like @idproect";
                //        //Command.Parameters.AddWithValue("@idproect", idproect);
                //    };
                //    if (comboBoxLocacia.Text != "")
                //    {
                //        commandText = commandText + @"and raschetEcoIntensive.location like @idlocatia";
                //        //Command.Parameters.AddWithValue("@idlocatia", idlocatia);
                //    };
                //}
                //else
                //{
                //    if (comboBoxProect.Text != "")
                //    {
                //        commandText = commandText + @"raschetEcoIntensive.project like @idproect";
                //        //Command.Parameters.AddWithValue("@idproect", idproect);
                //        if (comboBoxLocacia.Text != "")
                //        {
                //            commandText = commandText + @"and raschetEcoIntensive.location like @idlocatia";
                //            //Command.Parameters.AddWithValue("@idlocatia", idlocatia);
                //        };
                //    }
                //    else
                //    {
                //        if (comboBoxLocacia.Text != "")
                //        {
                //            commandText = commandText + @"raschetEcoIntensive.location like @idlocatia";
                //            //Command.Parameters.AddWithValue("@idlocatia", idlocatia);
                //        };
                //    };

                //};
                //Добавить фильтр даты

                Command = new SQLiteCommand { Connection = Connect,  CommandText = commandText };
                Command.Parameters.AddWithValue("@idpolzovatel", idpolzovatel);
                Command.Parameters.AddWithValue("@idproect", idproect);
                Command.Parameters.AddWithValue("@idlocatia", idlocatia);
                if (dateStart.Checked == true)
                {
                    Command.Parameters.AddWithValue("@dateStart", (dateStart.Value.ToString("yyyy-MM-dd") + " 00:00:00 "));
                }
                if (dateEnd.Checked == true)
                {
                    Command.Parameters.AddWithValue("@dateEnd", (dateEnd.Value.ToString("yyyy-MM-dd") + " 23:59:59 "));
                }

                SQLiteDataReader sqlReader = Command.ExecuteReader();
                int i = 0;
                TablicaDannihRascheta.Rows.Clear();
                while (sqlReader.Read())
                {
                    TablicaDannihRascheta.Rows.Add();
                    TablicaDannihRascheta.Rows[i].Cells[0].Value = sqlReader[0].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[1].Value = sqlReader[1].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[2].Value = sqlReader[2].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[3].Value = sqlReader[3].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[4].Value = sqlReader[4].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[5].Value = sqlReader[5].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[6].Value = sqlReader[6].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[7].Value = sqlReader[7].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[8].Value = sqlReader[8].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[9].Value = sqlReader[9].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[10].Value = sqlReader[10].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[11].Value = sqlReader[11].ToString();
                    i++;
                }
                sqlReader.Close();

                Connect.Close();

            }
            else 
            {
                bool indicator = false;

                commandText = commandText + @"SELECT rashetConflictInteres.id,Projects.project, Locations.location, Polzovateli.fio, date,p,c, de, d, e,dmax,dshag FROM rashetConflictInteres 
                                            LEFT JOIN Polzovateli ON rashetConflictInteres.polzovatel = Polzovateli.id
                                            LEFT JOIN Locations ON rashetConflictInteres.location = Locations.id
                                            LEFT JOIN Projects ON rashetConflictInteres.project = Projects.id";

                if (comboBoxPolzovatel.Text != "" || comboBoxLocacia.Text != "" || comboBoxProect.Text != "" || dateStart.Checked == true || dateEnd.Checked == true)
                {
                    commandText = commandText + @" where ";
                };

                if (comboBoxPolzovatel.Text != "")
                {
                    commandText = commandText + @"rashetConflictInteres.polzovatel like @idpolzovatel";
                    indicator = true;
                }
                if (comboBoxProect.Text != "")
                {
                    if (indicator == true)
                    {
                        commandText = commandText + " and ";
                    }
                    commandText = commandText + @"rashetConflictInteres.project like @idproect";
                    indicator = true;
                }

                if (comboBoxLocacia.Text != "")
                {
                    if (indicator == true)
                    {
                        commandText = commandText + " and ";
                    }
                    commandText = commandText + @"rashetConflictInteres.location like @idlocatia";
                    indicator = true;
                }
                if (dateStart.Checked == true)
                {
                    if (indicator == true)
                    {
                        commandText = commandText + " and ";
                    }
                    commandText = commandText + @"rashetConflictInteres.date >= @dateStart";
                    indicator = true;
                }
                if (dateEnd.Checked == true)
                {
                    if (indicator == true)
                    {
                        commandText = commandText + " and ";
                    }
                    commandText = commandText + @"rashetConflictInteres.date <= @dateEnd";
                    indicator = true;
                }




                //Добавить фильтр даты

                Command = new SQLiteCommand { Connection = Connect, CommandText = commandText };
                Command.Parameters.AddWithValue("@idpolzovatel", idpolzovatel);
                Command.Parameters.AddWithValue("@idproect", idproect);
                Command.Parameters.AddWithValue("@idlocatia", idlocatia);

                if (dateStart.Checked == true)
                {
                    Command.Parameters.AddWithValue("@dateStart", (dateStart.Value.ToString("yyyy-MM-dd") + " 00:00:00 "));
                }
                if (dateEnd.Checked == true)
                {
                    Command.Parameters.AddWithValue("@dateEnd", (dateEnd.Value.ToString("yyyy-MM-dd") + " 23:59:59 "));
                }

                SQLiteDataReader sqlReader = Command.ExecuteReader();
                int i = 0;
                TablicaDannihRascheta.Rows.Clear();
                while (sqlReader.Read())
                {
                    TablicaDannihRascheta.Rows.Add();
                    TablicaDannihRascheta.Rows[i].Cells[0].Value = sqlReader[0].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[1].Value = sqlReader[1].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[2].Value = sqlReader[2].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[3].Value = sqlReader[3].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[4].Value = sqlReader[4].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[5].Value = sqlReader[5].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[6].Value = sqlReader[6].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[7].Value = sqlReader[7].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[8].Value = sqlReader[8].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[9].Value = sqlReader[9].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[10].Value = sqlReader[10].ToString();
                    TablicaDannihRascheta.Rows[i].Cells[11].Value = sqlReader[11].ToString();
                    i++;
                }
                sqlReader.Close();

                Connect.Close();
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = TablicaDannihRascheta.CurrentCell.RowIndex;
            Program.znacheniastroki[0] = Convert.ToDouble(TablicaDannihRascheta[5, row].Value);
            Program.znacheniastroki[1] = Convert.ToDouble(TablicaDannihRascheta[6, row].Value);
            Program.znacheniastroki[2] = Convert.ToDouble(TablicaDannihRascheta[7, row].Value);
            Program.znacheniastroki[3] = Convert.ToDouble(TablicaDannihRascheta[8, row].Value);
            Program.znacheniastroki[4] = Convert.ToDouble(TablicaDannihRascheta[9, row].Value);
            Program.znacheniastroki[5] = Convert.ToDouble(TablicaDannihRascheta[10, row].Value);
            Program.znacheniastroki[6] = Convert.ToDouble(TablicaDannihRascheta[11, row].Value);

            for (int i = 0; i < TablicaDannihRascheta.Rows.Count; i++)
            {
                TablicaDannihRascheta.Rows[i].DefaultCellStyle.BackColor = SystemColors.ControlLightLight;
                TablicaDannihRascheta.Rows[i].DefaultCellStyle.ForeColor = SystemColors.ControlText;
            };
            
            TablicaDannihRascheta.Rows[row].DefaultCellStyle.BackColor = Color.RoyalBlue;
            TablicaDannihRascheta.Rows[row].DefaultCellStyle.ForeColor = Color.White;
        }

        private void dobavlenieButton_Click(object sender, EventArgs e)
        {
            if (Program.VidRascheta == false)
            {
                Program.RaschetEcoIntensiveForm.DobavitRaschetToolStripMenuItem_Click(Program.sender11, Program.e11);
                Program.RaschetEcoIntensiveForm.DobavitRaschetDB(Program.sender11, Program.e11);
            }
            else
            {
                Program.ConflictInteresovForm.DobavitRaschetToolStripMenuItem_Click(Program.sender11, Program.e11);
                Program.ConflictInteresovForm.DobavitRaschetDB(Program.sender11, Program.e11);
            }

            MessageBox.Show("Расчет добавлен!");

        }

        private void DobavlenieIzBD_FormClosed(object sender, FormClosedEventArgs e)
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



        private void DobavlenieIzBD_FormClosedbutton(object sender, EventArgs e)
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
    }
}