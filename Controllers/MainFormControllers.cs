using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace TaskScheduler
{
    class MainFormControllers
    {
        MainForm m_mf;
        V83.COMConnector com1s = new V83.COMConnector();
        ControlClaster control_claster = new ControlClaster();

        string my_directory = Directory.GetCurrentDirectory();        //присваиваем значение текущей директории
        dynamic infoBase;

        string user_1c = "";
        string pas_1c = "";
        string user_claster = "";
        string pas_claster = "";
        string name_server = "";
        string port_server = "";
        string mode = "";
        string name_soft = "";
        string clasterName = "";
        string key_access = "";
        string file_name = "";
        string path_backup = "";
        string path_log = "";

        MailData m_maildata;

        private void setInfoParameters()
        {
            user_1c = Properties.Settings.Default.user_1c;
            pas_1c = Properties.Settings.Default.pas_1c;
            user_claster = Properties.Settings.Default.user_claster;
            pas_claster = Properties.Settings.Default.pas_claster;
            name_server = Properties.Settings.Default.name_server;
            clasterName = Properties.Settings.Default.claster_name;
            mode = Properties.Settings.Default.mode;
            port_server = Properties.Settings.Default.port_server;
            key_access = Properties.Settings.Default.CodeBlocking;            
            file_name = Properties.Settings.Default.PathExe1c;
            path_backup = Properties.Settings.Default.PathBackup;
            path_log = Properties.Settings.Default.PathLog;

            m_maildata.smtpServer = Properties.Settings.Default.smtpServer;
            m_maildata.m_user = Properties.Settings.Default.mail_u;
            m_maildata.m_domain = Properties.Settings.Default.mail_d;
            m_maildata.from = Properties.Settings.Default.from;
            m_maildata.m_pass = Properties.Settings.Default.mail_p;
            m_maildata.mailto = Properties.Settings.Default.mailto;
        }

        public void setMainForm(MainForm mf)
        {
            m_mf = mf;
        }

        private void setPlatformVersion(string text)
        {
            m_mf.setPlatformVersion(text);
        }

        private void setConfigVersion(string text)
        {
            m_mf.setConfigVersion(text);
        }

        private void setAccessUpdate(string text)
        {
            m_mf.setAccessUpdate(text);
        }

        private string getNameBase()
        {
            return m_mf.getNameBase();
        }

        private void setLogText(string text)
        {
            m_mf.setLogText(text);
        }

        private void addLogText(string text)
        {
            m_mf.addLogText(text);
        }

        private void ClearLog()
        {
            m_mf.ClearLog();
        }

        private int findBaseName(string Name)
        {
            return m_mf.findNameBase(Name);
        }

        private void addBaseName(string Name)
        {
            m_mf.addNameBase(Name);
        }

        /// <summary>
        /// Получение информации с информационной базы
        /// </summary>
        public void GetBaseInfo()
        {
            ClearLog();
            setInfoParameters();
            string connectionStringClientServerDB = "Srvr='" + name_server + "';Ref='" + getNameBase() + "';Usr='" + user_1c + "';Pwd='" + pas_1c + "';";
            try
            {
                infoBase = com1s.Connect(connectionStringClientServerDB);

                var rv = infoBase.NewObject("СистемнаяИнформация");
                string recomendVersion = rv.ВерсияПриложения;
                string configurationVersion = infoBase.Метаданные.Синоним + " (" + infoBase.Метаданные.Версия + ")";
                bool isConfigurationModify = infoBase.КонфигурацияИзменена();

                setPlatformVersion("Версия платформы клиента: " + recomendVersion);
                setConfigVersion("Конфигурация ИБ: " + configurationVersion);
                setAccessUpdate("Доступно обновление конфигурации: " + isConfigurationModify);
            }
            catch (System.Runtime.InteropServices.COMException e) { setLogText(e.Message + " Возможно ИБ заблокирована."); }

        }

        public void StartTheard()
        {
            Parameters param = new Parameters();
            if (File.Exists(my_directory + @"\base.xml"))
            {
                param.dataSet1.ReadXml(my_directory + @"\base.xml", XmlReadMode.Auto);
            }
            setInfoParameters();
            if (!string.IsNullOrEmpty(port_server)) { name_server = name_server.Trim() + ":" + port_server.Trim().Replace("1540", "1541"); }
            string arguments = "";
            int c_t = param.dataSet1.Tables.Count;
            _Thread t1 = new _Thread();
            for (int n_t = 0; n_t < c_t; n_t++)
            {
                if (param.dataSet1.Tables[n_t].TableName == "TableBases")
                {
                    int c_r = param.dataSet1.Tables[n_t].Rows.Count;
                    for (int n_r = 0; n_r < c_r; n_r++)
                    {
                        /// Выполнить три метода
                        /// 1-заблокировать ИБ и отключить всех пользователей
                        /// 2-выгрузить базу данных в файл *.dt
                        /// 3-разблокировать ИБ на подключения пользователей
                        /// 
                        List<Arguments> l_arg = new List<Arguments>();

                        string PathBase = param.dataSet1.Tables[n_t].Rows[n_r].ItemArray[0].ToString();
                        string PrifexBase = param.dataSet1.Tables[n_t].Rows[n_r].ItemArray[2].ToString();
                        string _date = new _Format().DateNow_();

                        /// Заблокировать ИБ и отключить всех пользователей
                        if (!string.IsNullOrEmpty(name_server))
                        {
                            arguments = @"" + PathBase;
                            l_arg.Add(new Arguments() { Name = arguments });
                        }
                        /*{
                            arguments = @"" + " ENTERPRISE /F\"" + PathBase + "\" /N\"" + user_1c + "\" /P\"" + pas_1c + "\" /DisableStartupMessages /C\"ЗавершитьРаботуПользователей\" /Out\"" + path_log + "\\" + PrifexBase + "_" + "log1.txt\"";
                        }
                        else
                        {
                            arguments = @"" + " ENTERPRISE /S\"" + name_server + "\\" + PathBase + "\" /N\"" + user_1c + "\" /P\"" + pas_1c + "\" /DisableStartupMessages /C\"ЗавершитьРаботуПользователей\" /UC\"" + key_access + "\" /Out\"" + path_log + "\\" + PrifexBase + "_" + "log1.txt\"";
                        }
                        l_arg.Add(new Arguments() { Name = arguments });
                        */


                        /// Выгрузка ИБ
                        if (string.IsNullOrEmpty(name_server))
                        {
                            arguments = @"" + " DESIGNER /F\"" + PathBase + "\" /N\"" + user_1c + "\" /P\"" + pas_1c + "\" /UC" + key_access + " /DisableStartupMessages /DumpIB\"" + path_backup + "\\" + PrifexBase + "_" + _date + ".dt" + "\" /DumpResult " + path_log + "\\" + PrifexBase + "Result.txt\" /Out\"" + path_log + "\\" + PrifexBase + "_" + "log2.txt\"";
                        }
                        else
                        {
                            arguments = @"" + " DESIGNER /S\"" + name_server + "\\" + PathBase + "\" /N\"" + user_1c + "\" /P\"" + pas_1c + "\" /UC" + key_access + " /DisableStartupMessages /DumpIB\"" + path_backup + "\\" + PrifexBase + "_" + _date + ".dt" + "\" /DumpResult " + path_log + "\\" + PrifexBase + "Result.txt\" /Out\"" + path_log + "\\" + PrifexBase + "_" + "log2.txt\"";
                        }
                        l_arg.Add(new Arguments() { Name = arguments });


                        /// Разблокировать ИБ на подключения пользователей
                        if (!string.IsNullOrEmpty(name_server))
                        {
                            arguments = @"" + PathBase;
                            l_arg.Add(new Arguments() { Name = arguments });
                        }
                        /*{
                            arguments = @"" + " ENTERPRISE /F\"" + PathBase + "\" /N\"" + user_1c + "\" /P\"" + pas_1c + "\" /C\"РазрешитьРаботуПользователей\" /UC" + key_access + " /Out\"" + path_log + "\\" + PrifexBase + "_" + "log3.txt\"";
                        }
                        else
                        {
                            arguments = @"" + " ENTERPRISE /S\"" + name_server + "\\" + PathBase + "\" /N\"" + user_1c + "\" /P\"" + pas_1c + "\" /C\"РазрешитьРаботуПользователей\" /UC" + key_access + " /Out\"" + path_log + "\\" + PrifexBase + "_" + "log3.txt\"";
                        }
                        l_arg.Add(new Arguments() { Name = arguments });
                        */

                        t1.myThread("1C8Backup", l_arg, PathBase, PrifexBase, m_maildata);

                    }
                }
            }
        }

        public void GetListSessions(String DBName)
        {
            ClearLog();
            setInfoParameters();
            if (String.IsNullOrEmpty(clasterName)) { setLogText("Укажите в параметрах имя кластера."); return; }
            if (String.IsNullOrEmpty(name_server)) { setLogText("Укажите в параметрах имя сервера."); return; }

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
                        agent.UpdateInfoBase(clasterInfo, infoBase);
                        if (infoBase.Name == DBName)
                        {
                            if (session.AppID == "1CV8C") { name_soft = "Тонкий клиент"; }
                            else if (session.AppID == "1CV8") { name_soft = "Толстый клиент"; }
                            else if (session.AppID == "COMConnection") { name_soft = "COM-соединение"; }
                            else if (session.AppID == "Designer") { name_soft = "Конфигуратор"; }
                            else if (session.AppID == "SrvrConsole") { name_soft = "Консоль кластера"; }
                            else if (session.AppID == "COMConsole") { name_soft = "COM-администратор"; }

                            addLogText("Computer: " + session.Host + "; Number session: " + Convert.ToString(session.SessionID) + "; User name: " + session.userName + " (" + name_soft + ")" + "\r\n");
                        }
                    }
                }
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(agent);
        }

        /// <summary>
        /// Отключение Всех пользователей от базы
        /// </summary>
        public void DiscconnectUsers(String DBName)
        {
            control_claster.DiscconnectUsers(DBName);
        }

        /// <summary>
        /// Обновление свойства базы
        /// </summary>
        public void UpdateInfoBase(String DBName, bool connect_denied)
        {
            control_claster.UpdateInfoBase(DBName, connect_denied);
        }

        ///Заполняем список баз
        ///
        public void ListBase()
        {
            //ClearLog();
            setInfoParameters();
            if (String.IsNullOrEmpty(name_server))
            { setLogText("Проверьте параметры настроек для соединения с кластером сервера 1С!"); return; }
            try
            {
                V83.IServerAgentConnection agent = com1s.ConnectAgent(name_server + ":" + port_server);
                Array clasters = agent.GetClusters();
                foreach (V83.IClusterInfo clasterInfo in clasters)
                {
                    agent.Authenticate(clasterInfo, user_claster, pas_claster); //имя пользователя и пароль, пустое значение
                    Array DataBases = agent.GetInfoBases(clasterInfo);
                    foreach (V83.IInfoBaseShort DataBase in DataBases)
                    {
                        var session = agent.GetInfoBaseSessions(clasterInfo, DataBase);

                        if (findBaseName(DataBase.Name) == -1) { addBaseName(DataBase.Name); }
                    }
                }
            }
            catch
            {
                setLogText("Нет соединения с кластером сервера 1С!");
            }

        }

    }
}
