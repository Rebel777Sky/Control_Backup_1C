using System;
using System.Windows.Forms;

namespace TaskScheduler
{
    public partial class MainForm : Form
    {
        //Получить список баз данных на кластере
        //Получить список пользователей
        //Создать лог соединений и детализацией
        //Отключить пользователей от базы с фильтром игнорирования
        //Создания бэкап из конфигуратора 1С. Настройка хранения бэкапов (неделя,месяц,год)
        //Расписание запуска приложения, отключения и создания бэкап
            
        MainFormControllers mfc = new MainFormControllers();

        public MainForm()
        {
            InitializeComponent();
            Hide();
            mfc.setMainForm(this);
        }

        public void ClearLog()
        {
            textBoxLog.Text = "";
        }

        public void setPlatformVersion(string text)
        {
            label1.Text = text;
        }

        public void setConfigVersion(string text)
        {
            label2.Text = text;
        }

        public void setAccessUpdate(string text)
        {
            label3.Text = text;
        }

        public string getNameBase()
        {
            return comboBox1.Text;
        }

        public void addNameBase(string Name)
        {
            comboBox1.Items.Add(Name);
        }

        public int findNameBase(string Name)
        {
            return comboBox1.FindString(Name);
        }

        public int countNameBase()
        {
            return comboBox1.Items.Count;
        }

        public void setLogText(string text)
        {
            textBoxLog.Text = text;
        }

        public void addLogText(string text)
        {
            textBoxLog.Text += text;
        }

        private void ResizeForm()
        {
            if (this.WindowState == FormWindowState.Normal)
            {
                this.ShowInTaskbar = false;
                this.WindowState = FormWindowState.Minimized;
            }
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                Hide();
                notifyIcon1.Visible = true;
            }
        }

        public string DefineLocalName()
        {
            string Name = "";
            Name = Environment.MachineName;
            return Name;
        }

        public void OpenFormAbout()
        {
            AboutBox1 form_2 = new AboutBox1();
            form_2.ShowDialog();  //* Модальный режим
        }

        private void button2_Click(object sender, EventArgs e)
        {
            mfc.DiscconnectUsers(comboBox1.Text);         
        }

        private void button1_Click(object sender, EventArgs e)
        {
            mfc.GetBaseInfo();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            mfc.GetListSessions(comboBox1.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            mfc.UpdateInfoBase(comboBox1.Text, true);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            mfc.UpdateInfoBase(comboBox1.Text, false);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            bool start_service_auto = Properties.Settings.Default.start_service_auto;
            bool tray_auto = Properties.Settings.Default.tray_auto;
            if (start_service_auto) { mfc.StartTheard(); }
            if (tray_auto) { ResizeForm(); }
            mfc.ListBase();
        }

        private void ParametersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFormParametrs(0);
        }

        public void OpenFormParametrs(int LMenu)
        {
            Parameters form_2 = new Parameters();
            form_2.form_1 = this; //* Установить ссылку
            form_2.ListMenu = LMenu;
            form_2.ShowDialog();  //* Модальный режим
        }

        private void CloseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFormAbout();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            mfc.StartTheard();
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = false;
            this.ShowInTaskbar = true;
        }

        private void OptionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFormParametrs(0);
        }

        private void HelpToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            OpenFormAbout();
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

    }

}

