using System;
using System.Linq;
using System.Windows.Forms;


namespace one034
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {


            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Logowanie());
            //if(zalogowano)
            //MessageBox.Show(args0 +" i "+ args1);
            Application.Run(new ONE034Form());
        }
    }
}