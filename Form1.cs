using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using AutoUpdaterDotNET;

namespace NBpostel
{
    public partial class Form1 : Form
    {
        static string dir = System.IO.Directory.GetCurrentDirectory();
        static INIManager config = new INIManager(dir + "\\config.ini");

        
        //Всего
        int vsego_prostin = Convert.ToInt32(config.GetPrivateString("vsego", "prostin"));
        int vsego_pododel = Convert.ToInt32(config.GetPrivateString("vsego", "pododel"));
        int vsego_navoloch = Convert.ToInt32(config.GetPrivateString("vsego", "navoloch"));
        int vsego_mpol = Convert.ToInt32(config.GetPrivateString("vsego", "mpol"));
        int vsego_bpol = Convert.ToInt32(config.GetPrivateString("vsego", "bpol"));
        int vsego_hal = Convert.ToInt32(config.GetPrivateString("vsego", "hal"));

        //Грязного
        int grazn_prostin = Int32.Parse(config.GetPrivateString("grazn", "prostin"));
        int grazn_pododel = Int32.Parse(config.GetPrivateString("grazn", "pododel"));
        int grazn_navoloch = Int32.Parse(config.GetPrivateString("grazn", "navoloch"));
        int grazn_mpol = Int32.Parse(config.GetPrivateString("grazn", "mpol"));
        int grazn_bpol = Int32.Parse(config.GetPrivateString("grazn", "bpol"));
        int grazn_hal = Int32.Parse(config.GetPrivateString("grazn", "hal"));

        //В прачке
        int pratch_prostin = Int32.Parse(config.GetPrivateString("pratch", "prostin"));
        int pratch_pododel = Int32.Parse(config.GetPrivateString("pratch", "pododel"));
        int pratch_navoloch = Int32.Parse(config.GetPrivateString("pratch", "navoloch"));
        int pratch_mpol = Int32.Parse(config.GetPrivateString("pratch", "mpol"));
        int pratch_bpol = Int32.Parse(config.GetPrivateString("pratch", "bpol"));
        int pratch_hal = Int32.Parse(config.GetPrivateString("pratch", "hal"));

        //На койках
        int koik_prostin = Convert.ToInt32(config.GetPrivateString("koik", "prostin"));
        int koik_pododel = Int32.Parse(config.GetPrivateString("koik", "pododel"));
        int koik_navoloch = Int32.Parse(config.GetPrivateString("koik", "navoloch"));
        int koik_mpol = Int32.Parse(config.GetPrivateString("koik", "mpol"));
        int koik_bpol = Int32.Parse(config.GetPrivateString("koik", "bpol"));
        int koik_hal = Int32.Parse(config.GetPrivateString("koik", "hal"));

        //Вес
        int ves_prostin = Int32.Parse(config.GetPrivateString("ves", "prostin"));
        int ves_pododel = Int32.Parse(config.GetPrivateString("ves", "pododel"));
        int ves_navoloch = Int32.Parse(config.GetPrivateString("ves", "navoloch"));
        int ves_mpol = Int32.Parse(config.GetPrivateString("ves", "mpol"));
        int ves_bpol = Int32.Parse(config.GetPrivateString("ves", "bpol"));
        int ves_hal = Int32.Parse(config.GetPrivateString("ves", "hal"));

        //Settings
        int settings_act = Int32.Parse(config.GetPrivateString("settings", "act"));
        string company_act = (config.GetPrivateString("settings", "company")).ToString();

        public void update()
        {

            //Всего
            int vsego_prostin = Convert.ToInt32(config.GetPrivateString("vsego", "prostin"));
            int vsego_pododel = Convert.ToInt32(config.GetPrivateString("vsego", "pododel"));
            int vsego_mpol = Convert.ToInt32(config.GetPrivateString("vsego", "mpol"));
            int vsego_bpol = Convert.ToInt32(config.GetPrivateString("vsego", "bpol"));
            int vsego_hal = Convert.ToInt32(config.GetPrivateString("vsego", "hal"));

            //Грязного
            int grazn_prostin = Int32.Parse(config.GetPrivateString("grazn", "prostin"));
            int grazn_pododel = Int32.Parse(config.GetPrivateString("grazn", "pododel"));
            int grazn_navoloch = Int32.Parse(config.GetPrivateString("grazn", "navoloch"));
            int grazn_mpol = Int32.Parse(config.GetPrivateString("grazn", "mpol"));
            int grazn_bpol = Int32.Parse(config.GetPrivateString("grazn", "bpol"));
            int grazn_hal = Int32.Parse(config.GetPrivateString("grazn", "hal"));

            //В прачке
            int pratch_prostin = Int32.Parse(config.GetPrivateString("pratch", "prostin"));
            int pratch_pododel = Int32.Parse(config.GetPrivateString("pratch", "pododel"));
            int pratch_navoloch = Int32.Parse(config.GetPrivateString("pratch", "navoloch"));
            int pratch_mpol = Int32.Parse(config.GetPrivateString("pratch", "mpol"));
            int pratch_bpol = Int32.Parse(config.GetPrivateString("pratch", "bpol"));
            int pratch_hal = Int32.Parse(config.GetPrivateString("pratch", "hal"));

            //На койках
            int koik_prostin = Convert.ToInt32(config.GetPrivateString("koik", "prostin"));
            int koik_pododel = Int32.Parse(config.GetPrivateString("koik", "pododel"));
            int koik_navoloch = Int32.Parse(config.GetPrivateString("koik", "navoloch"));
            int koik_mpol = Int32.Parse(config.GetPrivateString("koik", "mpol"));
            int koik_bpol = Int32.Parse(config.GetPrivateString("koik", "bpol"));
            int koik_hal = Int32.Parse(config.GetPrivateString("koik", "hal"));

            //Вес
            int ves_prostin = Int32.Parse(config.GetPrivateString("ves", "prostin"));
            int ves_pododel = Int32.Parse(config.GetPrivateString("ves", "pododel"));
            int ves_navoloch = Int32.Parse(config.GetPrivateString("ves", "navoloch"));
            int ves_mpol = Int32.Parse(config.GetPrivateString("ves", "mpol"));
            int ves_bpol = Int32.Parse(config.GetPrivateString("ves", "bpol"));
            int ves_hal = Int32.Parse(config.GetPrivateString("ves", "hal"));

            //Settings
            int settings_act = Int32.Parse(config.GetPrivateString("settings", "act"));
            string company_act = (config.GetPrivateString("settings", "company")).ToString();

            //Грязного
            label7.Text = Convert.ToString(grazn_prostin);
            label8.Text = Convert.ToString(grazn_navoloch);
            label9.Text = Convert.ToString(grazn_pododel);
            label10.Text = Convert.ToString(grazn_mpol);
            label11.Text = Convert.ToString(grazn_bpol);
            label12.Text = Convert.ToString(grazn_hal);

            //В прачке
            label18.Text = Convert.ToString(pratch_prostin);
            label17.Text = Convert.ToString(pratch_navoloch);
            label16.Text = Convert.ToString(pratch_pododel);
            label15.Text = Convert.ToString(pratch_mpol);
            label14.Text = Convert.ToString(pratch_bpol);
            label13.Text = Convert.ToString(pratch_hal);

            //На койках
            label48.Text = Convert.ToString(koik_prostin);
            label47.Text = Convert.ToString(koik_navoloch);
            label46.Text = Convert.ToString(koik_pododel);
            label45.Text = Convert.ToString(koik_mpol);
            label44.Text = Convert.ToString(koik_bpol);
            label43.Text = Convert.ToString(koik_hal);

            //Settings
            textBox7.Text = company_act;
            numericUpDown7.Value = settings_act;

            //чистых
            int chist_prostin = (vsego_prostin - grazn_prostin - pratch_prostin - koik_prostin);
            int chist_pododel = (vsego_pododel - grazn_pododel - pratch_pododel - koik_pododel);
            int chist_navoloch = (vsego_navoloch - grazn_navoloch - pratch_navoloch - koik_navoloch);
            int chist_mpol = (vsego_mpol - grazn_mpol - pratch_mpol - koik_mpol);
            int chist_bpol = (vsego_bpol - grazn_bpol - pratch_bpol - koik_bpol);
            int chist_hal = (vsego_hal - grazn_hal - pratch_hal - koik_hal);

            label36.Text = Convert.ToString(chist_prostin);
            label35.Text = Convert.ToString(chist_navoloch);
            label34.Text = Convert.ToString(chist_pododel);
            label33.Text = Convert.ToString(chist_mpol);
            label32.Text = Convert.ToString(chist_bpol);
            label31.Text = Convert.ToString(chist_hal);

            //MessageBox.Show("Готово");

            //log
            listBox1.Items.Clear();
            string[] LogFile = File.ReadLines(dir + "\\log.txt").Reverse().Take(30).ToArray();
            for (int i = 0; i < LogFile.Length; i++)
            {
                listBox1.Items.Add(LogFile[i]);
            }

            //var lines = File.ReadAllLines(dir+ "\\log.txt").Reverse();
            //listBox1.Items.Add(lines);
            //listBox1.Items.AddRange(File.ReadAllLines(dir + "\\log.txt", Encoding.Default));
            //listBox1.Items.AddRange(File.ReadAllLines(dir + "\\log.txt", Encoding.Default));
        }

        public Form1()
        {
            InitializeComponent();

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void label29_Click(object sender, EventArgs e)
        {

        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        public void Form1_Load(object sender, EventArgs e)
        {
            update();
            AutoUpdater.Start("https://github.com/SawaMEN/MBpostel/raw/main/update.xml");
            label25.Text = "MBpostel v" + Application.ProductVersion.ToString() + " ©️RGBcorp, 2023";

            /* using (StreamReader r = new StreamReader("users.txt", Encoding.Default))
                while (!r.EndOfStream)
                    comboBox1.Items.Add(r.ReadLine()); */
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label36_Click(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            config.WritePrivateString("settings", "company", textBox7.Text);
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();    // создаем объект класса Form2 
            Form2.opis.lab = "Редактируем колличество грязного белья";
            newForm.Text = "Настройка грязного белья";

            Form2.opis.prostin = Int32.Parse(config.GetPrivateString("grazn", "prostin")); ;
            Form2.opis.pododel = Int32.Parse(config.GetPrivateString("grazn", "pododel"));
            Form2.opis.navoloch = Int32.Parse(config.GetPrivateString("grazn", "navoloch"));
            Form2.opis.mpol = Int32.Parse(config.GetPrivateString("grazn", "mpol"));
            Form2.opis.bpol = Int32.Parse(config.GetPrivateString("grazn", "bpol"));
            Form2.opis.hal = Int32.Parse(config.GetPrivateString("grazn", "hal"));
            Form2.opis.type = "grazn";

            newForm.delMethod = update;
            newForm.Show();
            //newForm.ShowDialog();           // Вызов формы-диалога
            //textBox1.Text = LogPar.login + "/" + LogPar.parol; // Результат
            newForm.delMethod = update;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();    // создаем объект класса Form2 
            Form2.opis.lab = "Редактируем колличество белья в прачке";
            newForm.Text = "Настройка белья в прачке";

            Form2.opis.prostin = Int32.Parse(config.GetPrivateString("pratch", "prostin"));
            Form2.opis.pododel = Int32.Parse(config.GetPrivateString("pratch", "pododel"));
            Form2.opis.navoloch = Int32.Parse(config.GetPrivateString("pratch", "navoloch"));
            Form2.opis.mpol = Int32.Parse(config.GetPrivateString("pratch", "mpol"));
            Form2.opis.bpol = Int32.Parse(config.GetPrivateString("pratch", "bpol"));
            Form2.opis.hal = Int32.Parse(config.GetPrivateString("pratch", "hal"));
            Form2.opis.type = "pratch";

            newForm.delMethod = update;
            newForm.ShowDialog();
        }

        private void numericUpDown7_ValueChanged(object sender, EventArgs e)
        {
            config.WritePrivateString("settings", "act", Convert.ToString(numericUpDown7.Value));
            //update();
        } 

        private void button5_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();    // создаем объект класса Form2 
            Form2.opis.lab = "Сколько белья привезли?";
            newForm.Text = "Привезли бельё";

            Form2.opis.prostin = Int32.Parse(config.GetPrivateString("pratch", "prostin"));
            Form2.opis.pododel = Int32.Parse(config.GetPrivateString("pratch", "pododel"));
            Form2.opis.navoloch = Int32.Parse(config.GetPrivateString("pratch", "navoloch"));
            Form2.opis.mpol = Int32.Parse(config.GetPrivateString("pratch", "mpol"));
            Form2.opis.bpol = Int32.Parse(config.GetPrivateString("pratch", "bpol"));
            Form2.opis.hal = Int32.Parse(config.GetPrivateString("pratch", "hal"));
            Form2.opis.type = "privoz";

            newForm.delMethod = update;
            newForm.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();
            Form2.opis.lab = "Редактируем вес белья";
            newForm.Text = "Настройка веса белья";

            Form2.opis.prostin = Int32.Parse(config.GetPrivateString("ves", "prostin")); ;
            Form2.opis.pododel = Int32.Parse(config.GetPrivateString("ves", "pododel"));
            Form2.opis.navoloch = Int32.Parse(config.GetPrivateString("ves", "navoloch"));
            Form2.opis.mpol = Int32.Parse(config.GetPrivateString("ves", "mpol"));
            Form2.opis.bpol = Int32.Parse(config.GetPrivateString("ves", "bpol"));
            Form2.opis.hal = Int32.Parse(config.GetPrivateString("ves", "hal"));
            Form2.opis.type = "ves";

            newForm.delMethod = update;
            newForm.Show();
            //newForm.ShowDialog();           // Вызов формы-диалога
            //textBox1.Text = LogPar.login + "/" + LogPar.parol; // Результат
            newForm.delMethod = update;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();    // создаем объект класса Form2 
            Form2.opis.lab = "Сколько белья увезла прачка?";
            newForm.Text = "Увезли бельё";

            Form2.opis.prostin = Int32.Parse(config.GetPrivateString("grazn", "prostin")); ;
            Form2.opis.pododel = Int32.Parse(config.GetPrivateString("grazn", "pododel"));
            Form2.opis.navoloch = Int32.Parse(config.GetPrivateString("grazn", "navoloch"));
            Form2.opis.mpol = Int32.Parse(config.GetPrivateString("grazn", "mpol"));
            Form2.opis.bpol = Int32.Parse(config.GetPrivateString("grazn", "bpol"));
            Form2.opis.hal = Int32.Parse(config.GetPrivateString("grazn", "hal"));
            Form2.opis.type = "uvoz";

            newForm.delMethod = update;
            newForm.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();    // создаем объект класса Form2 
            Form2.opis.lab = "Сколько белья убрали в грязное?";
            newForm.Text = "Убираем бельё в мешок";

            Form2.opis.prostin = Int32.Parse(config.GetPrivateString("grazn", "prostin"));
            Form2.opis.pododel = Int32.Parse(config.GetPrivateString("grazn", "pododel"));
            Form2.opis.navoloch = Int32.Parse(config.GetPrivateString("grazn", "navoloch"));
            Form2.opis.mpol = Int32.Parse(config.GetPrivateString("grazn", "mpol"));
            Form2.opis.bpol = Int32.Parse(config.GetPrivateString("grazn", "bpol"));
            Form2.opis.hal = Int32.Parse(config.GetPrivateString("grazn", "hal"));
            Form2.opis.type = "snali";

            newForm.delMethod = update;
            newForm.Show();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            button4.Enabled = false;
            
            // Создаем приложение Word
            Word.Application wordApp = new Word.Application();

            // Открываем документ
            Word.Document doc = wordApp.Documents.Open(dir + "\\act.docx");

            //Сохраняем копию документа
            doc.SaveAs(dir + "\\acts\\act_"+ config.GetPrivateString("settings", "act") + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".docx");

            // Получаем все закладки в документе
            Word.Bookmarks bookmarks = doc.Bookmarks;

            // Заполняем данные в закладках
            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "num_dogovor")
                    bookmark.Range.Text = config.GetPrivateString("settings", "act");


            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "company_act")
                    bookmark.Range.Text = config.GetPrivateString("settings", "company");

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "ves_navoloch")
                    bookmark.Range.Text = config.GetPrivateString("ves", "navoloch");

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "ves_pododel")
                    bookmark.Range.Text = config.GetPrivateString("ves", "pododel");

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "ves_prostin")
                    bookmark.Range.Text = config.GetPrivateString("ves", "prostin");

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "ves_mpol")
                    bookmark.Range.Text = config.GetPrivateString("ves", "mpol");

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "ves_bpol")
                    bookmark.Range.Text = config.GetPrivateString("ves", "bpol");

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "ves_hal")
                    bookmark.Range.Text = config.GetPrivateString("ves", "hal");


            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "grazn_navoloch")
                    bookmark.Range.Text = config.GetPrivateString("grazn", "navoloch");

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "grazn_pododel")
                    bookmark.Range.Text = config.GetPrivateString("grazn", "pododel");

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "grazn_prostin")
                    bookmark.Range.Text = config.GetPrivateString("grazn", "prostin");

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "grazn_mpol")
                    bookmark.Range.Text = config.GetPrivateString("grazn", "mpol");

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "grazn_bpol")
                    bookmark.Range.Text = config.GetPrivateString("grazn", "bpol");

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "grazn_hal")
                    bookmark.Range.Text = config.GetPrivateString("grazn", "hal");


            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "grazn_vsego")
                    bookmark.Range.Text = Convert.ToString( Int32.Parse(config.GetPrivateString("grazn", "prostin")) +
                                                            Int32.Parse(config.GetPrivateString("grazn", "pododel")) +
                                                            Int32.Parse(config.GetPrivateString("grazn", "navoloch")) +
                                                            Int32.Parse(config.GetPrivateString("grazn", "mpol")) +
                                                            Int32.Parse(config.GetPrivateString("grazn", "bpol")) +
                                                            Int32.Parse(config.GetPrivateString("grazn", "hal")));

            int sum_navoloch, sum_pododel, sum_prostin, sum_mpol, sum_hal, sum_bpol, sum_vsego;

            sum_navoloch = ves_navoloch * Int32.Parse(config.GetPrivateString("grazn", "navoloch"));
            sum_pododel = ves_pododel * Int32.Parse(config.GetPrivateString("grazn", "pododel"));
            sum_prostin = ves_prostin * Int32.Parse(config.GetPrivateString("grazn", "prostin"));
            sum_mpol = ves_mpol * Int32.Parse(config.GetPrivateString("grazn", "mpol"));
            sum_bpol = ves_bpol * Int32.Parse(config.GetPrivateString("grazn", "bpol"));
            sum_hal = ves_hal * Int32.Parse(config.GetPrivateString("grazn", "hal"));
            sum_vsego = sum_navoloch + sum_pododel + sum_prostin + sum_mpol + sum_bpol + sum_hal;

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "sum_navoloch")
                    bookmark.Range.Text = Convert.ToString(sum_navoloch);

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "sum_pododel")
                    bookmark.Range.Text = Convert.ToString(sum_pododel);

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "sum_prostin")
                    bookmark.Range.Text = Convert.ToString(sum_prostin);

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "sum_mpol")
                    bookmark.Range.Text = Convert.ToString(sum_mpol);

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "sum_hal")
                    bookmark.Range.Text = Convert.ToString(sum_hal);

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "sum_bpol")
                    bookmark.Range.Text = Convert.ToString(sum_bpol);

            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "sum_vsego")
                    bookmark.Range.Text = Convert.ToString(sum_vsego);


            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "date")
                    bookmark.Range.Text =  DateTime.Now.ToString("dd.MM.yyyy");


            //doc.SaveAs("C:\\template2.docx");
            // Закрываем документ и приложение Word
            doc.Close();
            wordApp.Quit();

            MessageBox.Show("Акт №" + config.GetPrivateString("settings", "act") + " от " + DateTime.Now.ToString("dd.MM.yyyy") + "\n Успешно создан!");
            button4.Enabled = true;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer", dir+"\\acts\\");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();    // создаем объект класса Form2 
            Form2.opis.lab = "Сколько белья постелили на кровать?";
            newForm.Text = "Кладем на койки";


            Form2.opis.prostin = Convert.ToInt32(config.GetPrivateString("koik", "prostin"));
            Form2.opis.pododel = Int32.Parse(config.GetPrivateString("koik", "pododel"));
            Form2.opis.navoloch = Int32.Parse(config.GetPrivateString("koik", "navoloch"));
            Form2.opis.mpol = Int32.Parse(config.GetPrivateString("koik", "mpol"));
            Form2.opis.bpol = Int32.Parse(config.GetPrivateString("koik", "bpol"));
            Form2.opis.hal = Int32.Parse(config.GetPrivateString("koik", "hal"));


            Form2.opis.type = "postelili";

            newForm.delMethod = update;
            newForm.Show();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();    // создаем объект класса Form2 
            Form2.opis.lab = "Редактируем колличество белья";
            newForm.Text = "Сколько всего белья";

            Form2.opis.prostin = Int32.Parse(config.GetPrivateString("vsego", "prostin"));
            Form2.opis.pododel = Int32.Parse(config.GetPrivateString("vsego", "pododel"));
            Form2.opis.navoloch = Int32.Parse(config.GetPrivateString("vsego", "navoloch"));
            Form2.opis.mpol = Int32.Parse(config.GetPrivateString("vsego", "mpol"));
            Form2.opis.bpol = Int32.Parse(config.GetPrivateString("vsego", "bpol"));
            Form2.opis.hal = Int32.Parse(config.GetPrivateString("vsego", "hal"));
            Form2.opis.type = "vsego";

            newForm.delMethod = update;
            newForm.ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            config.WritePrivateString("grazn", "prostin", Convert.ToString(Convert.ToInt32(config.GetPrivateString("grazn", "prostin")) + 1));
            config.WritePrivateString("grazn", "navoloch", Convert.ToString(Int32.Parse(config.GetPrivateString("grazn", "navoloch")) + 1));
            config.WritePrivateString("grazn", "pododel", Convert.ToString(Int32.Parse(config.GetPrivateString("grazn", "pododel")) + 1));
            
            using (var writer = new StreamWriter("log.txt", true))
                writer.WriteLine(DateTime.Now.ToString("dd//MM//yy") + " Поменяли 1 комплект");

            update();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            config.WritePrivateString("koik", "mpol", Convert.ToString(Convert.ToInt32(config.GetPrivateString("koik", "mpol")) + 1));

            using (var writer = new StreamWriter("log.txt", true))
                writer.WriteLine(DateTime.Now.ToString("dd//MM//yy") + " Выдали полотенце");

            update();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            config.WritePrivateString("grazn", "mpol", Convert.ToString(Int32.Parse(config.GetPrivateString("grazn", "mpol")) + 1));
            config.WritePrivateString("koik", "mpol", Convert.ToString(Convert.ToInt32(config.GetPrivateString("koik", "mpol")) - 1));

            using (var writer = new StreamWriter("log.txt", true))
                writer.WriteLine(DateTime.Now.ToString("dd//MM//yy") + " Забрали полотенце");

            update();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            config.WritePrivateString("koik", "bpol", Convert.ToString(Convert.ToInt32(config.GetPrivateString("koik", "bpol")) + 1));

            using (var writer = new StreamWriter("log.txt", true))
                writer.WriteLine(DateTime.Now.ToString("dd//MM//yy") + " Выдали большое полотенце");

            update();
        }

        private void label25_Click_1(object sender, EventArgs e)
        {

        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button15_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process txt = new System.Diagnostics.Process();
            txt.StartInfo.FileName = "notepad.exe";
            txt.StartInfo.Arguments = dir+"\\log.txt";
            txt.Start();
        }
    }

    public class INIManager
    {
        //Конструктор, принимающий путь к INI-файлу
        public INIManager(string aPath)
        {
            path = aPath;
        }

        //Конструктор без аргументов (путь к INI-файлу нужно будет задать отдельно)
        public INIManager() : this("") { }

        //Возвращает значение из INI-файла (по указанным секции и ключу) 
        public string GetPrivateString(string aSection, string aKey)
        {
            //Для получения значения
            StringBuilder buffer = new StringBuilder(SIZE);

            //Получить значение в buffer
            GetPrivateString(aSection, aKey, null, buffer, SIZE, path);

            //Вернуть полученное значение
            return buffer.ToString();
        }

        //Пишет значение в INI-файл (по указанным секции и ключу) 
        public void WritePrivateString(string aSection, string aKey, string aValue)
        {
            //Записать значение в INI-файл
            WritePrivateString(aSection, aKey, aValue, path);
        }

        //Возвращает или устанавливает путь к INI файлу
        public string Path { get { return path; } set { path = value; } }

        //Поля класса
        private const int SIZE = 1024; //Максимальный размер (для чтения значения из файла)
        private string path = null; //Для хранения пути к INI-файлу

        //Импорт функции GetPrivateProfileString (для чтения значений) из библиотеки kernel32.dll
        [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileString")]
        private static extern int GetPrivateString(string section, string key, string def, StringBuilder buffer, int size, string path);

        //Импорт функции WritePrivateProfileString (для записи значений) из библиотеки kernel32.dll
        [DllImport("kernel32.dll", EntryPoint = "WritePrivateProfileString")]
        private static extern int WritePrivateString(string section, string key, string str, string path);
    }
}
