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
        //public int c_prostin, c_pododel, c_navoloch, c_mpol, c_bpol, c_hal;

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
            
            if (opis.type == "snali")
            {
                numericUpDown1.Value = 0;
                numericUpDown2.Value = 0;
                numericUpDown3.Value = 0;
                numericUpDown4.Value = 0;
                numericUpDown5.Value = 0;
                numericUpDown6.Value = 0;
            }
            else
            if (opis.type == "postelili")
            {
                
                numericUpDown1.Value = 30 - opis.prostin;
                numericUpDown2.Value = 30 - opis.navoloch;
                numericUpDown3.Value = 30 - opis.pododel;
                numericUpDown4.Value = 0;
                numericUpDown5.Value = 0;
                numericUpDown6.Value = 0;
            }
            else
            {
                numericUpDown1.Value = opis.prostin;
                numericUpDown2.Value = opis.navoloch;
                numericUpDown3.Value = opis.pododel;
                numericUpDown4.Value = opis.mpol;
                numericUpDown5.Value = opis.bpol;
                numericUpDown6.Value = opis.hal;
            }
            
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

            if (opis.type == "privoz")
            {
                config.WritePrivateString("pratch", "prostin", Convert.ToString(opis.prostin - numericUpDown1.Value));
                config.WritePrivateString("pratch", "navoloch", Convert.ToString(opis.navoloch - numericUpDown2.Value));
                config.WritePrivateString("pratch", "pododel", Convert.ToString(opis.pododel - numericUpDown3.Value));
                config.WritePrivateString("pratch", "mpol", Convert.ToString(opis.mpol - numericUpDown4.Value));
                config.WritePrivateString("pratch", "bpol", Convert.ToString(opis.bpol - numericUpDown5.Value));
                config.WritePrivateString("pratch", "hal", Convert.ToString(opis.hal - numericUpDown6.Value));
            } else {
                if (opis.type == "uvoz") {
                    // Уменьшаем грязное
                    config.WritePrivateString("grazn", "prostin", Convert.ToString(opis.prostin - numericUpDown1.Value));
                    config.WritePrivateString("grazn", "navoloch", Convert.ToString(opis.navoloch - numericUpDown2.Value));
                    config.WritePrivateString("grazn", "pododel", Convert.ToString(opis.pododel - numericUpDown3.Value));
                    config.WritePrivateString("grazn", "mpol", Convert.ToString(opis.mpol - numericUpDown4.Value));
                    config.WritePrivateString("grazn", "bpol", Convert.ToString(opis.bpol - numericUpDown5.Value));
                    config.WritePrivateString("grazn", "hal", Convert.ToString(opis.hal - numericUpDown6.Value));

                    //увеличиваем в прачке
                    config.WritePrivateString("pratch", "prostin", Convert.ToString(Int32.Parse(config.GetPrivateString("pratch", "prostin")) + numericUpDown1.Value));
                    config.WritePrivateString("pratch", "navoloch", Convert.ToString(Int32.Parse(config.GetPrivateString("pratch", "navoloch")) + numericUpDown2.Value));
                    config.WritePrivateString("pratch", "pododel", Convert.ToString(Int32.Parse(config.GetPrivateString("pratch", "pododel")) + numericUpDown3.Value));
                    config.WritePrivateString("pratch", "mpol", Convert.ToString(Int32.Parse(config.GetPrivateString("pratch", "mpol")) + numericUpDown4.Value));
                    config.WritePrivateString("pratch", "bpol", Convert.ToString(Int32.Parse(config.GetPrivateString("pratch", "bpol")) + numericUpDown5.Value));
                    config.WritePrivateString("pratch", "hal", Convert.ToString(Int32.Parse(config.GetPrivateString("pratch", "hal")) + numericUpDown6.Value));
                } else {
                    if (opis.type == "snali") {
                        // Увеличиваем грязное
                        config.WritePrivateString("grazn", "prostin", Convert.ToString(opis.prostin + numericUpDown1.Value));
                        config.WritePrivateString("grazn", "navoloch", Convert.ToString(opis.navoloch + numericUpDown2.Value));
                        config.WritePrivateString("grazn", "pododel", Convert.ToString(opis.pododel + numericUpDown3.Value));
                        config.WritePrivateString("grazn", "mpol", Convert.ToString(opis.mpol + numericUpDown4.Value));
                        config.WritePrivateString("grazn", "bpol", Convert.ToString(opis.bpol + numericUpDown5.Value));
                        config.WritePrivateString("grazn", "hal", Convert.ToString(opis.hal + numericUpDown6.Value));

                        // Уменьшаем на койках
                        config.WritePrivateString("koik", "prostin", Convert.ToString(Convert.ToInt32(config.GetPrivateString("koik", "prostin")) - numericUpDown1.Value));
                        config.WritePrivateString("koik", "navoloch", Convert.ToString(Int32.Parse(config.GetPrivateString("koik", "navoloch")) - numericUpDown2.Value));
                        config.WritePrivateString("koik", "pododel", Convert.ToString(Int32.Parse(config.GetPrivateString("koik", "pododel")) - numericUpDown3.Value));
                        config.WritePrivateString("koik", "mpol", Convert.ToString(Int32.Parse(config.GetPrivateString("koik", "mpol")) - numericUpDown4.Value));
                        config.WritePrivateString("koik", "bpol", Convert.ToString(Int32.Parse(config.GetPrivateString("koik", "bpol")) - numericUpDown5.Value));
                        config.WritePrivateString("koik", "hal", Convert.ToString(Int32.Parse(config.GetPrivateString("koik", "hal")) - numericUpDown6.Value));
                    } else {
                        if (opis.type == "postelili") {
                            //Прибавляем на койки
                            config.WritePrivateString("koik", "prostin", Convert.ToString(Convert.ToInt32(config.GetPrivateString("koik", "prostin")) + numericUpDown1.Value));
                            config.WritePrivateString("koik", "navoloch", Convert.ToString(Int32.Parse(config.GetPrivateString("koik", "navoloch")) + numericUpDown2.Value));
                            config.WritePrivateString("koik", "pododel", Convert.ToString(Int32.Parse(config.GetPrivateString("koik", "pododel")) + numericUpDown3.Value));
                            config.WritePrivateString("koik", "mpol", Convert.ToString(Int32.Parse(config.GetPrivateString("koik", "mpol")) + numericUpDown4.Value));
                            config.WritePrivateString("koik", "bpol", Convert.ToString(Int32.Parse(config.GetPrivateString("koik", "bpol")) + numericUpDown5.Value));
                            config.WritePrivateString("koik", "hal", Convert.ToString(Int32.Parse(config.GetPrivateString("koik", "hal")) + numericUpDown6.Value));
                        } else {
                            config.WritePrivateString(opis.type, "prostin", Convert.ToString(numericUpDown1.Value));
                            config.WritePrivateString(opis.type, "navoloch", Convert.ToString(numericUpDown2.Value));
                            config.WritePrivateString(opis.type, "pododel", Convert.ToString(numericUpDown3.Value));
                            config.WritePrivateString(opis.type, "mpol", Convert.ToString(numericUpDown4.Value));
                            config.WritePrivateString(opis.type, "bpol", Convert.ToString(numericUpDown5.Value));
                            config.WritePrivateString(opis.type, "hal", Convert.ToString(numericUpDown6.Value));
                        }
                        
                    }


                    
                }
            }



            

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
