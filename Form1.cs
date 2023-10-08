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


            //Всего
            numericUpDown1.Value = vsego_prostin;
            numericUpDown2.Value = vsego_navoloch;
            numericUpDown3.Value = vsego_pododel;
            numericUpDown4.Value = vsego_mpol;
            numericUpDown5.Value = vsego_bpol;
            numericUpDown6.Value = vsego_hal;

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
            int chist_prostin = (vsego_prostin - grazn_prostin - pratch_prostin);
            int chist_pododel = (vsego_pododel - grazn_pododel - pratch_pododel);
            int chist_navoloch = (vsego_navoloch - grazn_navoloch - pratch_navoloch);
            int chist_mpol = (vsego_mpol - grazn_mpol - pratch_mpol);
            int chist_bpol = (vsego_bpol - grazn_bpol - pratch_bpol);
            int chist_hal = (vsego_hal - grazn_hal - pratch_hal);

            label36.Text = Convert.ToString(chist_prostin);
            label35.Text = Convert.ToString(chist_navoloch);
            label34.Text = Convert.ToString(chist_pododel);
            label33.Text = Convert.ToString(chist_mpol);
            label32.Text = Convert.ToString(chist_bpol);
            label31.Text = Convert.ToString(chist_hal);

            //MessageBox.Show("Готово");
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
            //INIManager config = new INIManager(dir + "\\config.ini");

            /*
            //Всего
            numericUpDown1.Value = vsego_prostin;
            numericUpDown2.Value = vsego_navoloch;
            numericUpDown3.Value = vsego_pododel;
            numericUpDown4.Value = vsego_mpol;
            numericUpDown5.Value = vsego_bpol;
            numericUpDown6.Value = vsego_hal;

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
            int chist_prostin = (vsego_prostin - grazn_prostin - pratch_prostin);
            int chist_pododel = (vsego_pododel - grazn_pododel - pratch_pododel);
            int chist_navoloch = (vsego_navoloch - grazn_navoloch - pratch_navoloch);
            int chist_mpol = (vsego_mpol - grazn_mpol - pratch_mpol);
            int chist_bpol = (vsego_bpol - grazn_bpol - pratch_bpol);
            int chist_hal = (vsego_hal - grazn_hal - pratch_hal);

            label36.Text = Convert.ToString(chist_prostin);
            label35.Text = Convert.ToString(chist_navoloch);
            label34.Text = Convert.ToString(chist_pododel);
            label33.Text = Convert.ToString(chist_mpol);
            label32.Text = Convert.ToString(chist_bpol);
            label31.Text = Convert.ToString(chist_hal);

            */
            update();

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

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            config.WritePrivateString("vsego", "prostin", Convert.ToString(numericUpDown1.Value));
            update();
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            config.WritePrivateString("vsego", "navoloch", Convert.ToString(numericUpDown2.Value));
            update();
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            config.WritePrivateString("vsego", "pododel", Convert.ToString(numericUpDown3.Value));
            update();
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            config.WritePrivateString("vsego", "mpol", Convert.ToString(numericUpDown4.Value));
            update();
        }

        private void numericUpDown5_ValueChanged(object sender, EventArgs e)
        {
            config.WritePrivateString("vsego", "bpol", Convert.ToString(numericUpDown5.Value));
            update();
        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {
            config.WritePrivateString("vsego", "hal", Convert.ToString(numericUpDown6.Value));
            update();
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
            newForm.delMethod = update;
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
            newForm.delMethod = update;
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
            newForm.delMethod = update;
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            // Создаем приложение Word
            Word.Application wordApp = new Word.Application();

            // Открываем документ
            Word.Document doc = wordApp.Documents.Open(dir + "\\act.docx");

            //Сохраняем копию документа
            doc.SaveAs(dir + "\\acts\\act_"+ Convert.ToString(settings_act) + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".docx");

            // Получаем все закладки в документе
            Word.Bookmarks bookmarks = doc.Bookmarks;

            // Заполняем данные в закладках
            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "num_dogovor")
                    bookmark.Range.Text = Convert.ToString(settings_act);


            foreach (Word.Bookmark bookmark in bookmarks)
                if (bookmark.Name == "company_act")
                    bookmark.Range.Text = company_act;

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
                    bookmark.Range.Text = Convert.ToString(grazn_prostin + grazn_navoloch + grazn_pododel + grazn_mpol + grazn_bpol + grazn_hal);

            int sum_navoloch, sum_pododel, sum_prostin, sum_mpol, sum_hal, sum_bpol, sum_vsego;

            sum_navoloch = ves_navoloch * grazn_navoloch;
            sum_pododel = ves_pododel * grazn_pododel;
            sum_prostin = ves_prostin * grazn_prostin;
            sum_mpol = ves_mpol * grazn_mpol;
            sum_bpol = ves_bpol * grazn_bpol;
            sum_hal = ves_hal * grazn_hal;
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

        }

        private void button9_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer", dir+"\\acts\\");
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
