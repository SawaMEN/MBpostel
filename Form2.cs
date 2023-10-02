using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NBpostel
{

    public struct Switsh    // Что передавать
    {
        public string lab; //лейбла
        public int prostin, pododel, navoloch, mpol, bpol, hal;
        public string type;
    }

    public partial class Form2 : Form
    {
        static public Switsh opis;
        public Form2()
        {
            InitializeComponent();
        }


        public delegate void del();
        public del delMethod;

        private void Form2_Load(object sender, EventArgs e)
        {
            label1.Text = opis.lab;

            numericUpDown1.Value = opis.prostin;
            numericUpDown2.Value = opis.navoloch;
            numericUpDown3.Value = opis.pododel;
            numericUpDown4.Value = opis.mpol;
            numericUpDown5.Value = opis.bpol;
            numericUpDown6.Value = opis.hal;

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            INIManager config = new INIManager(System.IO.Directory.GetCurrentDirectory() + "\\config.ini");
            config.WritePrivateString(opis.type, "prostin", Convert.ToString(numericUpDown1.Value));
            config.WritePrivateString(opis.type, "navoloch", Convert.ToString(numericUpDown2.Value));
            config.WritePrivateString(opis.type, "pododel", Convert.ToString(numericUpDown3.Value));
            config.WritePrivateString(opis.type, "mpol", Convert.ToString(numericUpDown4.Value));
            config.WritePrivateString(opis.type, "bpol", Convert.ToString(numericUpDown5.Value));
            config.WritePrivateString(opis.type, "hal", Convert.ToString(numericUpDown6.Value));

            

            //Form2.close();
            //Form1.ActiveForm.Invalidate(); 
            //Form1.ActiveForm.Update();
            delMethod();
            MessageBox.Show("Готово");
            this.Close();

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
