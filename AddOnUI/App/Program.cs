using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using AddOnUI.App;

namespace AddOnUI
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Main oMain = new Main();

            System.Windows.Forms.Application.Run();
        }
    }
}
