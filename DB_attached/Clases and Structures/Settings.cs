using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Net.Sockets;
using System.Collections.Concurrent;

namespace DB_attached
{
    [Serializable]
    public class Settings   //класс для работы с настройками программы
    {
        public string Site { get; set; }
        public string ProxyAddress { get; set; }
        public int ProxyPort { get; set; }
        public string Folder { get; set; }
        public string File { get; set; }
        [XmlIgnoreAttribute]
        public string Path
        {
            get
            {
                if (this.Folder == "Текущая папка")
                    return Directory.GetCurrentDirectory() + @"\" + this.File;
                else return this.Folder + this.File;
            }
        }
        public bool DownloadFile { get; set; }
        public int Threads { get; set; }
        public bool hidePassword { get; set; }
        public Credential[] Accounts { get; set; }
        public int EncryptLevel { get; set; }
        [XmlIgnoreAttribute]
        public bool TestPassed { get; set; }
        public bool RenameGender { get; set; }
        public bool ColumnAutoWidth { get; set; }
        public bool RenameColumnNames { get; set; }
        public bool AutoFilter { get; set; }
        public bool ColumnOrder { get; set; }
        public ColumnSynonim[] ColumnSynonims { get; set; }

        public Settings()
        {
            this.TestPassed = false;
        }

        public Credential[] CopyAccounts()  //создает копию учетных данных
        {
            Credential[] creds = new Credential[this.Accounts.Length];
            for (int i = 0; i < creds.Length; i++)
                creds[i] = this.Accounts[i].Copy();

            return creds;
        }
        public void SaveSettings()  //сохраняет настройки в xml
        {
            this.TestPassed = false;
            using (FileStream fs = new FileStream("Settings.xml", FileMode.Create))
            {
                XmlSerializer formatter = new XmlSerializer(typeof(Settings));
                formatter.Serialize(fs, this);
            }
        }
        public static Settings LoadSettings()  //загружает настройки из xml
        {
            using (FileStream fs = new FileStream("Settings.xml", FileMode.Open))
            {
                XmlSerializer formatter = new XmlSerializer(typeof(Settings));
                return (Settings)formatter.Deserialize(fs);
            }
        }
        public async Task<ConcurrentDictionary<string, bool>> Test()  //тестирование настроек
        {
            return await Task.Run(() =>
            {
                ConcurrentDictionary<string, bool> errors = new ConcurrentDictionary<string, bool>();

                int timeOut = 10000;

                //проверяем доступность прокси
                if (this.ProxyAddress != "")
                {

                    var client = new TcpClient();
                    try
                    {
                        if (client.ConnectAsync(this.ProxyAddress, this.ProxyPort).Wait(timeOut) == false)
                            throw new TimeoutException();
                    }
                    catch (Exception)
                    {
                        errors.TryAdd("proxy", true);
                    }
                    finally
                    {
                        client.Close();
                    }
                }
                if (!errors.ContainsKey("proxy"))
                    errors.TryAdd("proxy", false);

                //проверяем доступность сайта

                if (errors["proxy"] == false)
                {
                    try
                    {
                        HttpWebRequest webReq = (HttpWebRequest)HttpWebRequest.Create(this.Site);
                        webReq.Timeout = timeOut;

                        if (this.ProxyAddress != "")
                            webReq.Proxy = new WebProxy(this.ProxyAddress + ":" + this.ProxyPort);

                        webReq.GetResponse();
                        webReq.Abort();
                    }
                    catch (Exception)
                    {
                        errors.TryAdd("site", true);
                    }
                }
                if (!errors.ContainsKey("site"))
                    errors.TryAdd("site", false);

                //проверяем наличие папки
                if (!Directory.Exists(System.IO.Path.GetDirectoryName(this.Path)))
                    errors.TryAdd("folder", true);
                else
                    errors.TryAdd("folder", false);


                //проверяем наличие файла, если используется существующий файл
                if ((!this.DownloadFile) && (!System.IO.File.Exists(this.Path)))
                    errors.TryAdd("file", true);
                if (!errors.ContainsKey("file"))
                    errors.TryAdd("file", false);

                //проверяем кол-во потоков
                if ((this.Threads < 1) || (this.Threads > 50))
                    errors.TryAdd("threads", true);
                else
                    errors.TryAdd("threads", false);

                //проверяем логины    
                if ((errors["proxy"] == false) && (errors["site"] == false))
                {

                    Parallel.ForEach(this.Accounts, credential =>
                   {
                       using (WebSiteSRZ site = new WebSiteSRZ(this.Site, this.ProxyAddress, this.ProxyPort))
                       {
                           if (site.Authorize(credential).Result)
                           {
                               errors.TryAdd(credential.Login, false);
                               site.Logout();
                           }
                           else
                               errors.TryAdd(credential.Login, true);
                           
                       }
                   });
                }

                if (errors.Values.Contains(true))
                    this.TestPassed = false;
                else
                    this.TestPassed = true;

                return errors;
            });
        }
    }



}
