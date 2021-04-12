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
    public partial class Setting : Form
    {
        public Setting()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var settings = Properties.Settings.Default;
            settings.birth = checkBox2.Checked;
            settings.goden = checkBox4.Checked;
            settings.katgodn = checkBox5.Checked;
            settings.indx = checkBox6.Checked;
            settings.city = checkBox7.Checked;
            settings.district = checkBox8.Checked;
            settings.country = checkBox9.Checked;
            settings.street = checkBox10.Checked;
            settings.house = checkBox11.Checked;
            settings.flat = checkBox12.Checked;
            settings.phone = checkBox13.Checked;
            settings.homePhone = checkBox14.Checked;
            settings.citizenship = checkBox15.Checked;
            settings.passpSeries = checkBox16.Checked;
            settings.passpNumber = checkBox17.Checked;
            settings.passpKemVidan = checkBox18.Checked;
            settings.passpDate = checkBox19.Checked;
            settings.groupName = checkBox20.Checked;
            settings.otdelenie = checkBox21.Checked;
            settings.comm = checkBox22.Checked;
            settings.prikazNumIn = checkBox24.Checked;
            settings.dateIn = checkBox25.Checked;
            settings.prikazNumOut = checkBox23.Checked;
            settings.dateOut = checkBox26.Checked;
            settings.prichinaOut = checkBox27.Checked;
            settings.kval = checkBox28.Checked;
            settings.prikazNumKval = checkBox29.Checked;
            settings.akademotpusk = checkBox30.Checked;
            settings.Note = checkBox31.Checked;
            Close();
        }

        private void Setting_Load(object sender, EventArgs e)
        {
            var settings = Properties.Settings.Default;
            checkBox2.Checked = settings.birth;
            checkBox4.Checked = settings.goden;
            checkBox5.Checked = settings.katgodn;
            checkBox6.Checked = settings.indx;
            checkBox7.Checked = settings.city;
            checkBox8.Checked = settings.district;
            checkBox9.Checked = settings.country;
            checkBox10.Checked = settings.street;
            checkBox11.Checked = settings.house;
            checkBox12.Checked = settings.flat;
            checkBox13.Checked = settings.phone;
            checkBox14.Checked = settings.homePhone;
            checkBox15.Checked = settings.citizenship;
            checkBox16.Checked = settings.passpSeries;
            checkBox17.Checked = settings.passpNumber;
            checkBox18.Checked = settings.passpKemVidan;
            checkBox19.Checked = settings.passpDate;
            checkBox20.Checked = settings.groupName;
            checkBox21.Checked = settings.otdelenie;
            checkBox22.Checked = settings.comm;
            checkBox24.Checked = settings.prikazNumIn;
            checkBox25.Checked = settings.dateIn;
            checkBox23.Checked = settings.prikazNumOut;
            checkBox26.Checked = settings.dateOut;
            checkBox27.Checked = settings.prichinaOut;
            checkBox28.Checked = settings.kval;
            checkBox29.Checked = settings.prikazNumKval;
            checkBox30.Checked = settings.akademotpusk;
            checkBox31.Checked = settings.Note;

            //DetermineDBase dbase = new DetermineDBase();
            //label3.Text = dbase.Determine();
        }
    }
}
