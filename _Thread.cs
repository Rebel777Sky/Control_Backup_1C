using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;



namespace TaskScheduler
{
    class _Thread
    {
        Thread thread;
        string fileName = Properties.Settings.Default.PathExe1c;
        public string InfoLog = "";
        public string PathBase = "";
        public string PrifexBase = "";
        string PathLog = Properties.Settings.Default.PathLog;
        ControlClaster control_claster = new ControlClaster();

        MailData m_maildata;

        /// <summary>
        /// Запуск функции в отдельном потоке
        /// </summary>
        /// <param name="name">Наименование потока</param>
        /// <param name="arguments">Передача аргументов</param>
        public void myThread(string name, List<Arguments> arguments, string PathBase_, string PrifexBase_, MailData m_maildata) /// Конструктор получает имя функции и аргументы
        {
            PathBase = PathBase_;
            PrifexBase = PrifexBase_;
            InfoLog += new _Format().DateTimeNow_() + " Выгрузка информационной базы: " + PathBase + " \r\n";

            thread = new Thread(this.func);
            thread.Name = name;
            thread.Start(arguments); /// Передача аргументов в поток         
        }

        /// <summary>
        /// Функция работает с информационной базой
        /// </summary>
        /// <param name="arguments">Передача аргументов</param>
        void func(object arguments)
        {
            string PathBackup = Properties.Settings.Default.PathBackup;
            string _date = new _Format().DateNow_();
            var argument = (List<Arguments>)arguments;
            int arg_count = argument.Count;
            for (int arg_index = 0; arg_index < arg_count; arg_index++)
            {
                if (arg_index == 0)
                {
                    BaseDisconnect(argument, arg_index);
                }
                else if (arg_index == 1)
                {
                    BaseUpload(argument, arg_index);
                }
                else if (arg_index == 2)
                {
                    BaseConnect(argument, arg_index, PathBackup, _date);
                }
            }
        }

        private void BaseUpload(List<Arguments> argument, int arg_index)
        {
            InfoLog += new _Format().DateTimeNow_() + " Время начала выгрузки информационной базы." + " \r\n";
            Process p = new Process();
            p.StartInfo.FileName = fileName;
            p.StartInfo.Arguments = argument[arg_index].Name;
            p.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
            p.Start();
            Console.WriteLine(p.Id);
            p.WaitForExit();/// Ждем завершения процесcа 

            string Path_ = @"" + PathLog + "\\" + PrifexBase + "_" + "log2.txt";
            if (File.Exists(Path_))
            {
                string[] ReadText = File.ReadAllLines(Path_, System.Text.Encoding.GetEncoding(1251));/// Выгрузка текстового файла в массив с указанной кодировкой
                foreach (string s in ReadText)
                {
                    InfoLog += s + " \r\n";
                }
            }
            InfoLog += new _Format().DateTimeNow_() + " Процесс завершен." + " \r\n";
        }

        private void BaseDisconnect(List<Arguments> argument, int arg_index)
        {
            InfoLog += new _Format().DateTimeNow_() + " Время начала процесса отключения пользователей от базы." + " \r\n";
            control_claster.UpdateInfoBase(argument[arg_index].Name, true);
            control_claster.DiscconnectUsers(argument[arg_index].Name);
            InfoLog += new _Format().DateTimeNow_() + " Процесс завершен." + " \r\n";
        }

        private void BaseConnect(List<Arguments> argument, int arg_index, string PathBackup, string _date)
        {
            InfoLog += new _Format().DateTimeNow_() + " Время начала подключения пользователей к работе в базе." + " \r\n";
            control_claster.UpdateInfoBase(argument[arg_index].Name, false);
            InfoLog += new _Format().DateTimeNow_() + " Процесс завершен." + " \r\n";
            InfoLog += " \r\n";
            InfoLog += @"Завершен процесс выгрузка информационной базы: " + PathBase + "! \r\n";
            InfoLog += @"Архив ИБ располагается локально на компьютере: " + new MainForm().DefineLocalName() + "! \r\n";
            InfoLog += @"Место расположение архива (" + PrifexBase + "_" + _date + ".dt" + ") : \"" + PathBackup + "\\\"! \r\n";
            FileStream fs;
            string fileName1_new = PathLog + "\\" + PrifexBase + "_" + "log_new.txt";
            if (File.Exists(fileName1_new)) { fs = new FileStream(PathLog + "\\" + PrifexBase + "_" + "log_new.txt", FileMode.Open, FileAccess.Write, FileShare.None); }
            else { fs = new FileStream(PathLog + "\\" + PrifexBase + "_" + "log_new.txt", FileMode.CreateNew, FileAccess.Write, FileShare.None); }
            StreamWriter writer = new StreamWriter(fs, System.Text.Encoding.GetEncoding(1251));
            writer.WriteLine(InfoLog);
            writer.Close();

            string Path_ = @"" + PathLog + "\\" + PrifexBase + "_" + "log_new.txt";
            if (File.Exists(Path_))
            {
                InfoLog = "";
                string[] ReadText = File.ReadAllLines(Path_, System.Text.Encoding.GetEncoding(1251));/// Выгрузка текстового файла в массив с указанной кодировкой
                foreach (string s in ReadText)
                {
                    InfoLog += s + " \r\n";
                }
            }

            m_maildata.caption = "Выгрузка информационной базы 1С. " + PrifexBase;
            m_maildata.message = InfoLog;
            if (String.IsNullOrEmpty(m_maildata.smtpServer) ||
                String.IsNullOrEmpty(m_maildata.from) ||
                String.IsNullOrEmpty(m_maildata.m_pass) ||
                String.IsNullOrEmpty(m_maildata.mailto)) { return; }
            new _Net().SendMail(m_maildata);//mailsend("Выгрузка информационной базы 1С. " + PrifexBase, InfoLog);
        }

    }

    class Arguments
    {
        public string Name { get; set; }
    }
}
