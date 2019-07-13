using System;

namespace TaskScheduler
{
    public class ControlClaster
    {
        V83.COMConnector com1s1 = new V83.COMConnector();

        string user_claster = Properties.Settings.Default.user_claster;
        string pas_claster = Properties.Settings.Default.pas_claster;
        string name_server = Properties.Settings.Default.name_server;
        string key_access = Properties.Settings.Default.CodeBlocking;
        string clasterName = Properties.Settings.Default.claster_name;
        string user_1c = Properties.Settings.Default.user_1c;
        string pas_1c = Properties.Settings.Default.pas_1c;

        /// <summary>
        /// Отключение Всех пользователей от базы
        /// </summary>
        public void DiscconnectUsers(String DBName)
        {
            string textBoxLog = "";
            V83.IServerAgentConnection agent = com1s1.ConnectAgent(name_server);
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

                        if (infoBase.Name == DBName)
                        {
                            textBoxLog += "DataBase: " + infoBase.Name + "; Number session: " + Convert.ToString(session.SessionID) + "; user name: " + session.userName + " (" + session.AppID + ")" + "\r\n";
                        }
                        if ((infoBase.Name == DBName) && session.AppID != "Designer" && session.AppID != "SrvrConsole")
                        {
                            agent.TerminateSession(clasterInfo, session);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Обновление свойства базы
        /// </summary>
        /// значение "F"- для файловой "S" - для серв.
        /// 1.Путь к исполняему файлу 1С
        /// 2.Имя пользователя администратора
        /// 3.Пароль администратора
        /// "C:\Program Files (x86)\1cv8\8.3.10.\bin\1cv8.exe" ENTERPRISE / F "\У" / N "Администратор" / P "123" / CРазрешитьРаботуПользователей / UCКодРазрешения
        public void UpdateInfoBase(String DBName, bool connect_denied)
        {
            V83.IServerAgentConnection AgentConnection = com1s1.ConnectAgent(name_server);
            Array clasters = AgentConnection.GetClusters();

            foreach (V83.IClusterInfo clasterInfo in clasters)
            {
                if (String.Equals(clasterInfo.HostName, name_server) && String.Equals(clasterInfo.ClusterName, clasterName))
                {
                    AgentConnection.Authenticate(clasterInfo, user_claster, pas_claster);

                    Array WorkingProcess = AgentConnection.GetWorkingProcesses(clasterInfo);
                    String ConnectString = "";
                    foreach (V83.IWorkingProcessInfo WorkingProces in WorkingProcess)
                    {
                        ConnectString = WorkingProces.HostName + ":" + WorkingProces.MainPort;
                    }

                    V83.IInfoBaseInfo infoBase_;
                    V83.IWorkingProcessConnection WorkingProcessConnection1 = com1s1.ConnectWorkingProcess(ConnectString);
                    Array Bases_ = WorkingProcessConnection1.GetInfoBases();
                    foreach (V83.IInfoBaseInfo base_ in Bases_)
                    {
                        if (base_.Name == DBName)
                        {
                            infoBase_ = base_;

                            V83.IInfoBaseInfo new_info = WorkingProcessConnection1.CreateInfoBaseInfo();
                            new_info = infoBase_;
                            if (String.IsNullOrEmpty(new_info.PermissionCode) == true)
                            {
                                new_info.PermissionCode = key_access;
                            }
                            new_info.DeniedMessage = "Проводяться регламентные работы!";
                            new_info.ConnectDenied = connect_denied;
                            WorkingProcessConnection1.AddAuthentication(user_1c, pas_1c);
                            WorkingProcessConnection1.UpdateInfoBase(new_info);
                            break;
                        }
                    }
                    Array sessions = AgentConnection.GetSessions(clasterInfo);
                    foreach (V83.ISessionInfo session in sessions)
                    {
                        if (session.AppID == "COMConsole") { AgentConnection.TerminateSession(clasterInfo, session); }
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(AgentConnection);
                }
            }
        }

    }
}
