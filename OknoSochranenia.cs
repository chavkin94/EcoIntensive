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
    public partial class OknoSochranenia : Form
    {
        string commandText;
        SQLiteCommand Command;
        int counter = 0;
        public OknoSochranenia()
        {
            InitializeComponent();
        }

        private void OknoSochranenia_Load(object sender, EventArgs e)
        {
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

        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            int idpolzovatel = -1;
            int idproect = -1;
            int idlocatia = -1;

            SQLiteConnection Connect = null;
            Connect.Open();
            
            if (comboBoxPolzovatel.Text != "")
            {
                Command = new SQLiteCommand
                {
                    Connection = Connect,
                    CommandText = @"select id, fio from Polzovateli where fio like @perem"
                };
                Command.Parameters.AddWithValue("@perem", comboBoxPolzovatel.Text);
                SQLiteDataReader sqlReader = Command.ExecuteReader();
                counter = 0;
                while (sqlReader.Read())
                {
                    counter++;
                };
                if (counter == 0)
                {
                    commandText = "INSERT INTO [Polzovateli] ([fio]) VALUES (@perem)";
                    Command = new SQLiteCommand(commandText, Connect);
                    Command.Parameters.AddWithValue("@perem", comboBoxPolzovatel.Text);
                    Command.ExecuteNonQuery();
                    Command = new SQLiteCommand
                    {
                        Connection = Connect,
                        CommandText = @"select id, fio from Polzovateli where fio like @perem"
                    };
                    Command.Parameters.AddWithValue("@perem", comboBoxPolzovatel.Text);
                    sqlReader = Command.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        idpolzovatel = Convert.ToInt32(sqlReader["id"]);
                    };
                }
                else
                {
                    SQLiteCommand Command1 = new SQLiteCommand
                    {
                        Connection = Connect,
                        CommandText = @"select id, fio from Polzovateli where fio like @perem"
                    };
                    Command1.Parameters.AddWithValue("@perem", comboBoxPolzovatel.Text);
                    SQLiteDataReader sqlReader1 = Command1.ExecuteReader();
                    while (sqlReader1.Read())
                    {
                        idpolzovatel = Convert.ToInt32(sqlReader1["id"]);
                    };
                };
            };

            if (comboBoxProect.Text != "")
            {
                Command = new SQLiteCommand
                {
                    Connection = Connect,
                    CommandText = @"select id, project from Projects where project like @perem"
                };
                Command.Parameters.AddWithValue("@perem", comboBoxProect.Text);
                SQLiteDataReader sqlReader = Command.ExecuteReader();
                counter = 0;
                while (sqlReader.Read())
                {
                    counter++;
                };
                if (counter == 0)
                {
                    commandText = "INSERT INTO [Projects] ([project]) VALUES (@perem)";
                    Command = new SQLiteCommand(commandText, Connect);
                    Command.Parameters.AddWithValue("@perem", comboBoxProect.Text);
                    Command.ExecuteNonQuery();
                    Command = new SQLiteCommand
                    {
                        Connection = Connect,
                        CommandText = @"select id, project from Projects where project like @perem"
                    };
                    Command.Parameters.AddWithValue("@perem", comboBoxProect.Text);
                    sqlReader = Command.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        idproect = Convert.ToInt32(sqlReader["id"]);
                    };
                }
                else
                {
                    SQLiteCommand Command3 = new SQLiteCommand
                    {
                        Connection = Connect,
                        CommandText = @"select id, project from Projects where project like @perem"
                    };
                    Command3.Parameters.AddWithValue("@perem", comboBoxProect.Text);
                    SQLiteDataReader sqlReader3 = Command3.ExecuteReader();
                    while (sqlReader3.Read())
                    {
                        idproect = Convert.ToInt32(sqlReader3["id"]);
                    };
                };
            };

            if (comboBoxLocacia.Text != "")
            {
                Command = new SQLiteCommand
                {
                    Connection = Connect,
                    CommandText = @"select id, location from Locations where location like @perem"
                };
                Command.Parameters.AddWithValue("@perem", comboBoxLocacia.Text);
                SQLiteDataReader sqlReader = Command.ExecuteReader();
                counter = 0;
                while (sqlReader.Read())
                {
                    counter++;
                };
                if (counter == 0)
                {

                    commandText = "INSERT INTO [Locations] ([location]) VALUES (@perem)";
                    Command = new SQLiteCommand(commandText, Connect);
                    Command.Parameters.AddWithValue("@perem", comboBoxLocacia.Text);
                    Command.ExecuteNonQuery();
                    Command = new SQLiteCommand
                    {
                        Connection = Connect,
                        CommandText = @"select id, location from Locations where location like @perem"
                    };
                    Command.Parameters.AddWithValue("@perem", comboBoxLocacia.Text);
                    sqlReader = Command.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        idlocatia = Convert.ToInt32(sqlReader["id"]);
                    };
                }
                else
                {
                    SQLiteCommand Command2 = new SQLiteCommand
                    {
                        Connection = Connect,
                        CommandText = @"select id, location from Locations where location like @perem"
                    };
                    Command2.Parameters.AddWithValue("@perem", comboBoxLocacia.Text);
                    SQLiteDataReader sqlReader2 = Command2.ExecuteReader();
                    while (sqlReader2.Read())
                    {
                        idlocatia = Convert.ToInt32(sqlReader2["id"]);
                    };
                };
            };

            if (Program.VidRascheta == false)
            {
                commandText = "INSERT INTO [raschetEcoIntensive] ([project], [location], [polzovatel], [date], [e], [l], [r], [de], [n], [demax], [deshag]) VALUES (@project, @location, @polzovatel, @date, @e, @l, @r, @de, @n, @demax, @deshag)";
                Command = new SQLiteCommand(commandText, Connect);
                if (idlocatia >= 0)
                    Command.Parameters.AddWithValue("@location", idlocatia);
                else
                    Command.Parameters.AddWithValue("@location", null);
                if (idproect >= 0)
                    Command.Parameters.AddWithValue("@project", idproect);
                else
                    Command.Parameters.AddWithValue("@project", null);
                if (idpolzovatel >= 0)
                    Command.Parameters.AddWithValue("@polzovatel", idpolzovatel);
                else
                    Command.Parameters.AddWithValue("@polzovatel", null);
                Command.Parameters.AddWithValue("@date", DateTime.Now);
                Command.Parameters.AddWithValue("@e", Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 0]);
                Command.Parameters.AddWithValue("@l", Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 1]);
                Command.Parameters.AddWithValue("@r", Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 2]);
                Command.Parameters.AddWithValue("@de", Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 3]);
                Command.Parameters.AddWithValue("@n", Convert.ToInt32((Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 4])));
                Command.Parameters.AddWithValue("@demax", Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 5]);
                Command.Parameters.AddWithValue("@deshag", Program.SpisokEcoIntens[Program.TecushiyEcoIntens[0], 6]);
                Command.ExecuteNonQuery();
            }
            else
            {

                commandText = "INSERT INTO [rashetConflictInteres] ([project], [location], [polzovatel], [date], [p], [c], [de], [d], [e], [dmax], [dshag]) VALUES (@project, @location, @polzovatel, @date, @p, @c, @de, @d, @e, @dmax, @dshag)";
                Command = new SQLiteCommand(commandText, Connect);
                if (idlocatia >= 0)
                    Command.Parameters.AddWithValue("@location", idlocatia);
                else
                    Command.Parameters.AddWithValue("@location", null);
                if (idproect >= 0)
                    Command.Parameters.AddWithValue("@project", idproect);
                else
                    Command.Parameters.AddWithValue("@project", null);
                if (idpolzovatel >= 0)
                    Command.Parameters.AddWithValue("@polzovatel", idpolzovatel);
                else
                    Command.Parameters.AddWithValue("@polzovatel", null);
                Command.Parameters.AddWithValue("@date", DateTime.Now);
                Command.Parameters.AddWithValue("@p", Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 0]);
                Command.Parameters.AddWithValue("@c", Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 1]);
                Command.Parameters.AddWithValue("@de", Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 2]);
                Command.Parameters.AddWithValue("@d", Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 3]);
                Command.Parameters.AddWithValue("@e", Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 4]);
                Command.Parameters.AddWithValue("@dmax", Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 5]);
                Command.Parameters.AddWithValue("@dshag", Program.SpisokConflictInteres[Program.TecushiyConflictInteres[0], 6]);
                Command.ExecuteNonQuery();
            }
            Connect.Close();
            

        }

        private void OknoSochranenia_FormClosed(object sender, FormClosedEventArgs e)
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
