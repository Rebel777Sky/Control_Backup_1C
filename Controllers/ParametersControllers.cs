using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace TaskScheduler
{
    class ParametersControllers
    {
        
        string user_1c = "";
        string pas_1c = "";
        string user_claster = "";
        string pas_claster = "";
        string name_server = "";
        string port_server = "";
        string mode = "";
        string clasterName = "";
        string key_access = "";
        string file_name = "";
        string path_backup = "";
        string path_log = "";
        List<string> list_base_name = new List<string>();

        V83.COMConnector com1s = new V83.COMConnector();
        Parameters m_pc;
        MailData m_maildata;

        public void setParametersForm(Parameters pf)
        {
            m_pc = pf;
        }

        private bool getTableBaseContains(string text)
        {
            return m_pc.getTableBaseContains(text);
        }

        private DataRow getTableBaseNewRow()
        {
            return m_pc.getTableBaseNewRow();
        }

        private void setTableBaseAdd(DataRow text)
        {
            m_pc.setTableBaseAdd(text);
        }

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

        public void UpdateTableBases()
        {
            ListBase();
            try
            {
                int c_i = list_base_name.Count;
                for (int c_ii = 0; c_ii < c_i; c_ii++)
                {
                    DataRow custumerRow = getTableBaseNewRow();
                    custumerRow[0] = list_base_name[c_ii].ToString();
                    custumerRow[1] = "v83";
                    custumerRow[2] = "";
                    if (!getTableBaseContains(custumerRow[0].ToString()))
                    {
                        setTableBaseAdd(custumerRow);
                    }
                }
                MessageBox.Show("Список ИБ обновлен!");
            }catch(System.NullReferenceException e) { MessageBox.Show(e.Message); }
        }

        ///Заполняем список баз
        ///
        public void ListBase()
        {
            setInfoParameters();
            if (String.IsNullOrEmpty(name_server))
            { MessageBox.Show("Проверьте параметры настроек для соединения с кластером сервера 1С!"); return; }
            try
            {
                V83.IServerAgentConnection agent = com1s.ConnectAgent(name_server + ":" + port_server);
                Array clasters = agent.GetClusters();
                foreach (V83.IClusterInfo clasterInfo in clasters)
                {
                    agent.Authenticate(clasterInfo, "", ""); //имя пользователя и пароль, пустое значение
                    Array DataBases = agent.GetInfoBases(clasterInfo);
                    int index_ = 0;
                    foreach (V83.IInfoBaseShort DataBase in DataBases)
                    {
                        var session = agent.GetInfoBaseSessions(clasterInfo, DataBase);
                                                
                        if (list_base_name.Count == 0)
                        {
                            list_base_name.Add(DataBase.Name);
                        }else if(list_base_name.Contains(DataBase.Name) == false)
                        {
                            list_base_name.Add(DataBase.Name);
                        }
                        index_++;
                    }
                }
            }
            catch
            {
                MessageBox.Show("Нет соединения с кластером сервера 1С!");
            }

        }

    }
}
