using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 ab = new AboutBox1();
            ab.Show();
        }

        Student st;

        private void button1_Click(object sender, EventArgs e)
        {
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            if (st == null || st.IsDisposed)
            {
                st = new Student();
                st.Show();
            }
            this.Cursor = System.Windows.Forms.Cursors.Default;
            
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        Group gr;

        private void button2_Click(object sender, EventArgs e)
        {
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            if (gr == null || gr.IsDisposed)
            { 
                gr = new Group();
                gr.Show();
            }
            this.Cursor = System.Windows.Forms.Cursors.Default;
        }
    }
}
