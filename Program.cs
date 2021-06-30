using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using System.IO;

namespace EcoIntensive
{
    static class Program
    {
       

        [STAThread]
        static void Main()
        {

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            RaschetEcoIntensiveForm = new RaschetEcoIntensive();
            TecushiyEcoIntens[2] = 0;
            Application.Run(RaschetEcoIntensiveForm);

        }

        public static double[,] SpisokEcoIntens = new double[1000, 200];
        public static string[] NameSpisokEcoIntens = new string[1000];
        public static double[,] SpisokConflictInteres = new double[1000, 200];
        public static string[] NameSpisokConflictInteres = new string[1000];
        public static bool VidRascheta = false;
        public static int[] TecushiyEcoIntens = new int[3];
        public static int[] TecushiyConflictInteres = new int[3];

        public static RaschetEcoIntensive RaschetEcoIntensiveForm ;
        public static ConflictInteresov ConflictInteresovForm;
        public static double[] znacheniastroki = new double[10];

        public static object sender11;
        public static EventArgs e11;

    }
}
