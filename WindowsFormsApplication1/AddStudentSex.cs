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
    public partial class AddStudentSex : Form
    {
        public AddStudentSex()
        {
            InitializeComponent();
        }

        public string ChooseSex()
        {
            string sex = "";
            if (radioButton1.Checked == true)
                sex = "муж.";
            if (radioButton2.Checked == true)
                sex = "жен.";
            return sex;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Hide();
        }
    }
}
