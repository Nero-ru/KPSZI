using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPSZI
{
    public partial class splashForm : Form
    {
        public splashForm()
        {
            InitializeComponent();
            //this.FormClosing += new FormClosingEventHandler(beforeClosing);

            this.Opacity = 0;
            bool fadingIn = true;

            Timer timer = new Timer();
            timer.Tick += new EventHandler((s, e1) =>
            {
                if (fadingIn)
                {
                    if ((Opacity += 0.05d) >= 1)
                    {
                        fadingIn = false;
                        timer.Stop();
                    }
                }
            });
            timer.Interval = 25;
            timer.Start();
        }

        public void beforeClosing(object sender, FormClosingEventArgs e)
        {
            this.Opacity = 100;
            bool fadingOut = true;

            Timer timer = new Timer();
            timer.Tick += new EventHandler((s, e1) =>
            {
                if (fadingOut)
                {
                    if ((Opacity -= 0.10d) <= 0)
                    {
                        fadingOut = false;
                        timer.Stop();
                    }
                }
            });
            timer.Interval = 50;
            timer.Start();
        }
    }
}
