using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ControlBackup1C
{
    public partial class MainForm_1 : Form
    {
        //Получить список баз данных на кластере
        //Получить список пользователей
        //Создать лог соединений и детализацией
        //Отключить пользователей от базы с фильтром игнорирования
        //Создания бэкап из конфигуратора 1С. Настройка хранения бэкапов (неделя,месяц,год)
        //Расписание запуска приложения, отключения и создания бэкап

        V83.COMConnector com1s = new V83.COMConnector();

        string user_1c = "";
        string pas_1c = "";
        string user_claster = "";
        string pas_claster = "";
        string name_server = "";
        string port_server = "";
        string mode = "";
        string name_soft = "";
        ///session.AppID == mode
        ///Режим внешнего соединения: COMConnection
        ///Режим конфигуратора: Designer
        ///Режим предприятие: 1CV8C
        string myDirectory = Directory.GetCurrentDirectory();        //присваиваем значение текущей директории
        dynamic infoBase;

        public MainForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Получение информации с информационной базы
        /// </summary>
        private void GetBaseInfo()
        {
            string connectionStringClientServerDB = "Srvr='" + name_server + "';Ref='" + comboBox1.Text + "';Usr='" + user_1c + "';Pwd='" + pas_1c + "';";
            infoBase = com1s.Connect(connectionStringClientServerDB);

            var rv = infoBase.NewObject("СистемнаяИнформация");
            string recomendVersion = rv.ВерсияПриложения;
            string configurationVersion = infoBase.Метаданные.Версия;
            bool isConfigurationModify = infoBase.КонфигурацияИзменена();

            label1.Text = "Версия платформы клиента: " + recomendVersion;
            label2.Text = "Версия конфигурации ИБ: " + configurationVersion;
            label3.Text = "Доступно обновление конфигурации: " + isConfigurationModify;

            //int currentSession = infoBase.НомерСеансаИнформационнойБазы();
        }

        private void UpdateOption()
        {

        }

        ///Заполняем список баз
        ///
        public void ListBase()
        {
            if (String.IsNullOrEmpty(name_server) ||
                String.IsNullOrEmpty(user_claster) ||
                String.IsNullOrEmpty(pas_claster))
            { MessageBox.Show("Нет данных для соединения с кластером!"); return; }
            V83.IServerAgentConnection agent = com1s.ConnectAgent(name_server);
            Array clasters = agent.GetClusters();
            foreach (V83.IClusterInfo clasterInfo in clasters)
            {
                agent.Authenticate(clasterInfo, user_claster, pas_claster);
                Array sessions = agent.GetSessions(clasterInfo);
                V83.IInfoBaseShort infoBase1;
                foreach (V83.ISessionInfo session in sessions)
                {
                    infoBase1 = session.infoBase;
                    if (comboBox1.FindString(infoBase1.Name) == -1) { comboBox1.Items.Add(infoBase1.Name); }
                }
            }
        }

        /// <summary>
        /// Отключение Всех пользователей от базы
        /// </summary>
        private void Discconnect()
        {
            string clasterName = "Локальный кластер";
            textBoxLog.Text = "";

            V83.IServerAgentConnection agent = com1s.ConnectAgent(name_server);
            Array clasters = agent.GetClusters();

            foreach (V83.IClusterInfo clasterInfo in clasters)
            {
                if (String.Equals(clasterInfo.HostName, name_server) && String.Equals(clasterInfo.ClusterName, clasterName))
                {
                    agent.Authenticate(clasterInfo, user_claster, pas_claster);
                    Array sessions = agent.GetSessions(clasterInfo);
                    V83.IInfoBaseShort infoBase;
                    foreach (V83.ISessionInfo session in sessions)
                    {
                        infoBase = session.infoBase;
                        //if (String.Equals(infoBase.Name, db) && session.SessionID == sessionID)
                        //{
                        //    agent.TerminateSession(clasterInfo, session);
                        //}
                        if (String.Equals(infoBase.Name, comboBox1.Text))
                        {
                            textBoxLog.Text += "DataBase: " + infoBase.Name + "; Number session: " + Convert.ToString(session.SessionID) + "; user name: " + session.userName + " (" + session.AppID + ")" + "\r\n";
                        }
                        if (String.Equals(infoBase.Name, comboBox1.Text) && session.AppID != mode)
                        {
                            agent.TerminateSession(clasterInfo, session);
                        }
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Discconnect();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetBaseInfo();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string clasterName = "Локальный кластер";
            textBoxLog.Text = "";

            V83.IServerAgentConnection agent = com1s.ConnectAgent(name_server);
            Array clasters = agent.GetClusters();

            foreach (V83.IClusterInfo clasterInfo in clasters)
            {
                if (String.Equals(clasterInfo.HostName, name_server) && String.Equals(clasterInfo.ClusterName, clasterName))
                {
                    agent.Authenticate(clasterInfo, user_claster, pas_claster);
                    Array sessions = agent.GetSessions(clasterInfo);
                    V83.IInfoBaseShort infoBase;
                    foreach (V83.ISessionInfo session in sessions)
                    {
                        infoBase = session.infoBase;
                        if (String.Equals(infoBase.Name, comboBox1.Text))
                        {
                            if (session.AppID == "1CV8C") { name_soft = "Тонкий клиент"; }
                            else if (session.AppID == "COMConnection") { name_soft = "COM-соединение"; }
                            else if (session.AppID == "Designer") { name_soft = "Конфигуратор"; }
                            textBoxLog.Text += "Computer: " + session.Host + "; Number session: " + Convert.ToString(session.SessionID) + "; User name: " + session.userName + " (" + name_soft + ")" + "\r\n";
                        }
                    }
                }
            }

        }

        /// <summary>
        /// Блокировка базы
        /// </summary>
        /// значение "F"- для файловой "S" - для серв.
        /// 1.Путь к исполняему файлу 1С
        /// 2.Имя пользователя администратора
        /// 3.Пароль администратора
        /// "C:\Program Files (x86)\1cv8\8.3.10.\bin\1cv8.exe" ENTERPRISE / F "\У" / N "Администратор" / P "123" / CРазрешитьРаботуПользователей / UCКодРазрешения
        private void BlokingBase()
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            BlokingBase();
        }

        /// <summary>
        /// Разблокировка базы
        /// </summary>
        /// "C:\Program Files (x86)\1cv8\8.3.10.\bin\1cv8.exe" ENTERPRISE / F"C:\1С\infobase" / N "123" / P "123" / WA - / AU - / DisableStartupMessages / CРазрешитьРаботуПользователей / UCКодРазрешения
        private void UnlokingBase()
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            UnlokingBase();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //ListBase();
            //textBoxPathExe1c.Text = Properties.Settings.Default.PathExe1c;
            user_1c = Properties.Settings.Default.user_1c;
            pas_1c = Properties.Settings.Default.pas_1c;
            user_claster = Properties.Settings.Default.user_claster;
            pas_claster = Properties.Settings.Default.pas_claster;
            name_server = Properties.Settings.Default.name_server;
            mode = Properties.Settings.Default.mode;
            port_server = Properties.Settings.Default.port_server;
            
             StartTheard();
        }

        ///Резервное копирования базы 1С
        ///
        private void button6_Click(object sender, EventArgs e)
        {
            /// "C:\Program Files (x86)\1cv8\8.3.10...\bin\1cv8.exe" DESIGNER / F "C:\Users\Documents\1C\Trade2" / N "Админ" / P "админ" / DumpIB "D:\bat backups\%backup_date%.dt"
            /// /DumpResult — после ключа должно быть указано имя файла, в который будет записан результат работы конфигуратора. Число, ноль в случае успеха. 
            /// rem / DumpIB "D:\bat backups\%backup_date%.dt"
        }

        private void параметрыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFormParametrs(0);
        }

        /// <summary>
        /// Открытие формы "Настройки"
        /// </summary>
        public void OpenFormParametrs(int LMenu)
        {
            Parameters form_2 = new Parameters();
            form_2.form_1 = this; //* Установить ссылку
            form_2.ListMenu = LMenu;
            form_2.ShowDialog();  //* Модальный режим
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Открытие формы "Настройки"
        /// </summary>
        public void OpenFormAbout()
        {
            AboutBox1 form_2 = new AboutBox1();
            form_2.ShowDialog();  //* Модальный режим
        }

        private void опрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFormAbout();
        }

        /// <summary>
        /// Отправка письма на почтовый ящик
        /// </summary>
        /// <param name="_caption">Тема письма</param>
        /// <param name="_message">Сообщение</param>
        public void mailsend(string _caption, string _message)
        {
            string smtpServer = Properties.Settings.Default.smtpServer;
            string m_user = Properties.Settings.Default.mail_u;
            string m_domain = Properties.Settings.Default.mail_d;
            string from = Properties.Settings.Default.from;
            string m_pass = Properties.Settings.Default.mail_p;
            string mailto = Properties.Settings.Default.mailto;
            if (String.IsNullOrEmpty(smtpServer) ||
                String.IsNullOrEmpty(from) ||
                String.IsNullOrEmpty(m_pass) ||
                String.IsNullOrEmpty(mailto)) { return; }
            _Net _mail = new _Net();
            _mail.SendMail(smtpServer, m_user, m_domain, from, m_pass, mailto, _caption, _message);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            mailsend("ВНИМАНИЕ!", "Тестовое сообщение.");
        }

        public void StartTheard()
        {
            Parameters param = new Parameters();
            if (File.Exists(myDirectory + @"\base.xml"))
            {
                param.dataSet1.ReadXml(myDirectory + @"\base.xml", XmlReadMode.Auto);
            }
            string u = Properties.Settings.Default.user_1c;
            string pa = Properties.Settings.Default.pas_1c;
            string cb = Properties.Settings.Default.CodeBlocking;
            string Server1c = Properties.Settings.Default.name_server;
            string fileName = Properties.Settings.Default.PathExe1c;
            string PathBackup = Properties.Settings.Default.PathBackup;
            string PathLog = Properties.Settings.Default.PathLog;
            string arguments = "";
            int c_t = param.dataSet1.Tables.Count;
            for (int n_t = 0; n_t < c_t; n_t++)
            {
                if (param.dataSet1.Tables[n_t].TableName == "TableBases")
                {
                    int c_r = param.dataSet1.Tables[n_t].Rows.Count;
                    for (int n_r = 0; n_r < c_r; n_r++)
                    {
                        _Thread t1 = new _Thread();
                        /// Выполняется три метода
                        /// 1-выгнать всех пользователей и заблокировать базу
                        /// 2-выгрузить базу данных в файл *.dt
                        /// 3-разблокировать базу
                        /// 
                        List<Arguments> l_arg = new List<Arguments>();

                        string PathBase = param.dataSet1.Tables[n_t].Rows[n_r].ItemArray[0].ToString();
                        string PrifexBase = param.dataSet1.Tables[n_t].Rows[n_r].ItemArray[2].ToString();
                        string _date = new _Format().DateNow_();

                        /// Выгоняем пользователей
                        if (String.IsNullOrEmpty(Server1c))
                        {
                            arguments = @"" + " ENTERPRISE /F\"" + PathBase + "\" /N\"" + u + "\" /P\"" + pa + "\" /AU- /WA- /DisableStartupMessages /C\"ЗавершитьРаботуПользователей\" /Out\"" + PathLog + "\\" + PrifexBase + "_" + "log1.txt\"";
                        }
                        else
                        {
                            arguments = @"" + " ENTERPRISE /S\"" + Server1c + "\\" + PathBase + "\" /N\"" + u + "\" /P\"" + pa + "\" /AU- /WA- /DisableStartupMessages /C\"ЗавершитьРаботуПользователей\" /UC\"" + cb + "\" /Out\"" + PathLog + "\\" + PrifexBase + "_" + "log1.txt\"";
                        }
                        l_arg.Add(new Arguments() { Name = arguments });

                        /// Выгрузка ИБ
                        if (String.IsNullOrEmpty(Server1c))
                        {
                            arguments = @"" + " DESIGNER /F\"" + PathBase + "\" /N\"" + u + "\" /P\"" + pa + "\" /UC" + cb + " /DisableStartupMessages /DumpIB\"" + PathBackup + "\\" + PrifexBase + "_" + _date + ".dt" + "\" /DumpResult " + PathLog + "\\" + PrifexBase + "Result.txt\" /Out\"" + PathLog + "\\" + PrifexBase + "_" + "log2.txt\"";
                        }
                        else
                        {
                            arguments = @"" + " DESIGNER /S\"" + Server1c + "\\" + PathBase + "\" /N\"" + u + "\" /P\"" + pa + "\" /UC" + cb + " /DisableStartupMessages /DumpIB\"" + PathBackup + "\\" + PrifexBase + "_" + _date + ".dt" + "\" /DumpResult " + PathLog + "\\" + PrifexBase + "Result.txt\" /Out\"" + PathLog + "\\" + PrifexBase + "_" + "log2.txt\"";
                        }
                        l_arg.Add(new Arguments() { Name = arguments });

                        /// Запускаем пользователей
                        if (String.IsNullOrEmpty(Server1c))
                        {
                            arguments = @"" + " ENTERPRISE /F\"" + PathBase + "\" /N\"" + u + "\" /P\"" + pa + "\" /C\"РазрешитьРаботуПользователей\" /UC" + cb + " /Out\"" + PathLog + "\\" + PrifexBase + "_" + "log3.txt\"";
                        }
                        else
                        {
                            arguments = @"" + " ENTERPRISE /S\"" + Server1c + "\\" + PathBase + "\" /N\"" + u + "\" /P\"" + pa + "\" /C\"РазрешитьРаботуПользователей\" /UC" + cb + " /Out\"" + PathLog + "\\" + PrifexBase + "_" + "log3.txt\"";
                        }
                        l_arg.Add(new Arguments() { Name = arguments });

                        t1.myThread("1C8Backup", l_arg, PathBase, PrifexBase);

                  }
                }
            }
        }

        public string DefineLocalName()
        {
            string Name = "";
            Name = Environment.MachineName;
            return Name;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            StartTheard();
        }
    }

}

