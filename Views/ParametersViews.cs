using System;
using System.Data;
using System.Windows.Forms;
using System.IO;

namespace TaskScheduler
{
    public partial class Parameters : Form
    {
        public MainForm form_1; /// Добавляем ссылку на вызывающее окно
        public int ListMenu;
        string myDirectory = Directory.GetCurrentDirectory();        //присваиваем значение текущей директории

        ParametersControllers m_pc = new ParametersControllers();
        MailData m_maildata;

        public Parameters()
        {
            InitializeComponent();
            m_pc.setParametersForm(this);
        }

        public bool getTableBaseContains(string text)
        {
            return dataSet1.Tables["TableBases"].Rows.Contains(text);
        }

        public DataRow getTableBaseNewRow()
        {
            return dataSet1.Tables["TableBases"].NewRow();
        }

        public void setTableBaseAdd(DataRow text)
        {
            dataSet1.Tables["TableBases"].Rows.Add(text);
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            if (ListMenu > 0)
            {
                listBoxMenu.SetSelected(ListMenu, true);
            }
            else
            {
                listBoxMenu.SetSelected(0, true);
            }
            ///Почтовые параметры
            textBoxSmtpServer.Text = Properties.Settings.Default.smtpServer;
            textBoxFrom.Text = Properties.Settings.Default.from;
            textBoxMailUser.Text = Properties.Settings.Default.mail_u;
            textBoxMailDomain.Text = Properties.Settings.Default.mail_d;
            textBoxMailPass.Text = Properties.Settings.Default.mail_p;
            textBoxMailto.Text = Properties.Settings.Default.mailto;
            ///
            textBoxCodeBlok.Text = Properties.Settings.Default.CodeBlocking;
            textBoxPathBackup.Text = Properties.Settings.Default.PathBackup;
            textBoxPathLog.Text = Properties.Settings.Default.PathLog;
            textBoxPathExe1c.Text = Properties.Settings.Default.PathExe1c;
            textBoxL1c.Text = Properties.Settings.Default.user_1c;
            textBoxP1c.Text = Properties.Settings.Default.pas_1c;
            textBoxLC.Text = Properties.Settings.Default.user_claster;
            textBoxPC.Text = Properties.Settings.Default.pas_claster;
            textBoxNameSer.Text = Properties.Settings.Default.name_server;
            textBoxMode.Text = Properties.Settings.Default.mode;
            textBoxPortServer.Text = Properties.Settings.Default.port_server;
            chbx_AutoService.Checked = Properties.Settings.Default.start_service_auto;
            chbx_tray.Checked = Properties.Settings.Default.tray_auto;

            textBoxClasterName.Text = Properties.Settings.Default.claster_name;

            textBoxMode.ReadOnly = true;
            textBoxCodeBlok.ReadOnly = true;

            if (File.Exists(myDirectory + @"\base.xml"))
            {
                dataSet1.ReadXml(myDirectory + @"\base.xml", XmlReadMode.Auto);
            }
        }

        private void SelectFileEXE()
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "Файлы 1с |*.exe"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBoxPathExe1c.Text = dialog.FileName;
            }
        }

        /// Путь до исполняемого файла 1с
        private void button2_Click(object sender, EventArgs e)
        {
            SelectFileEXE();
        }

        private void SaveParameters()
        {
            ///Почтовые параметры
            Properties.Settings.Default.smtpServer = textBoxSmtpServer.Text;
            Properties.Settings.Default.mail_u = textBoxMailUser.Text;
            Properties.Settings.Default.mail_p = textBoxMailPass.Text;
            Properties.Settings.Default.mail_d = textBoxMailDomain.Text;
            Properties.Settings.Default.from = textBoxFrom.Text;
            Properties.Settings.Default.mailto = textBoxMailto.Text;
            ///
            Properties.Settings.Default.CodeBlocking = textBoxCodeBlok.Text;
            Properties.Settings.Default.PathBackup = textBoxPathBackup.Text;
            Properties.Settings.Default.PathLog = textBoxPathLog.Text;
            Properties.Settings.Default.PathExe1c = textBoxPathExe1c.Text;
            Properties.Settings.Default.user_1c = textBoxL1c.Text;
            Properties.Settings.Default.pas_1c = textBoxP1c.Text;
            Properties.Settings.Default.user_claster = textBoxLC.Text;
            Properties.Settings.Default.pas_claster = textBoxPC.Text;
            Properties.Settings.Default.name_server = textBoxNameSer.Text;
            Properties.Settings.Default.mode = textBoxMode.Text;
            Properties.Settings.Default.port_server = textBoxPortServer.Text;
            Properties.Settings.Default.start_service_auto = chbx_AutoService.Checked;
            Properties.Settings.Default.tray_auto = chbx_tray.Checked;
            Properties.Settings.Default.claster_name = textBoxClasterName.Text;
            Properties.Settings.Default.Save();

            var myDirectory = Directory.GetCurrentDirectory();        //присваиваем значение текущей директории
            dataSet1.WriteXml(myDirectory + @"\base.xml", XmlWriteMode.WriteSchema);
        }

        private void b_Ok_Click(object sender, EventArgs e)
        {
            SaveParameters();
            Close();
        }

        private void butCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void listBoxMenu_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxMenu.SelectedIndex == 0)
            {
                panel1.Enabled = true;
                panel1.Visible = true;
                labTitle.Text = listBoxMenu.Text;
            }
            else
            {
                panel1.Enabled = false;
                panel1.Visible = false;
            }
            if (listBoxMenu.SelectedIndex == 1)
            {
                panel2.Enabled = true;
                panel2.Visible = true;
                labTitle.Text = listBoxMenu.Text;
            }
            else
            {
                panel2.Enabled = false;
                panel2.Visible = false;
            }
            if (listBoxMenu.SelectedIndex == 2)
            {
                panel3.Enabled = true;
                panel3.Visible = true;
                labTitle.Text = listBoxMenu.Text;
            }
            else
            {
                panel3.Enabled = false;
                panel3.Visible = false;
            }
            if (listBoxMenu.SelectedIndex == 3)
            {
                panel4.Enabled = true;
                panel4.Visible = true;
                labTitle.Text = listBoxMenu.Text;
            }
            else
            {
                panel4.Enabled = false;
                panel4.Visible = false;
            }
            if (listBoxMenu.SelectedIndex == 4)
            {
                panel5.Enabled = true;
                panel5.Visible = true;
                labTitle.Text = listBoxMenu.Text;
            }
            else
            {
                panel5.Enabled = false;
                panel5.Visible = false;
            }
        }

        /// <summary>
        /// Открытие формы "Планировщик"
        /// </summary>
        public void OpenFormTaskScheduler()
        {
            _TaskScheduler form_2 = new _TaskScheduler();
            form_2.ShowDialog();  //* Модальный режим
        }

        private void TestSendMail()
        {
            SaveParameters();
            m_maildata.smtpServer = Properties.Settings.Default.smtpServer;
            m_maildata.m_user = Properties.Settings.Default.mail_u;
            m_maildata.m_domain = Properties.Settings.Default.mail_d;
            m_maildata.from = Properties.Settings.Default.from;
            m_maildata.m_pass = Properties.Settings.Default.mail_p;
            m_maildata.mailto = Properties.Settings.Default.mailto;
            m_maildata.caption = "ВНИМАНИЕ!";
            m_maildata.message = "Тестовое сообщение.";
            if (String.IsNullOrEmpty(m_maildata.smtpServer) ||
                String.IsNullOrEmpty(m_maildata.from) ||
                String.IsNullOrEmpty(m_maildata.m_pass) ||
                String.IsNullOrEmpty(m_maildata.mailto)) { return; }
            new _Net().SendMail(m_maildata);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxPathBackup.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxPathLog.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            m_pc.UpdateTableBases();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFormTaskScheduler();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            TestSendMail();
        }

    }
}
