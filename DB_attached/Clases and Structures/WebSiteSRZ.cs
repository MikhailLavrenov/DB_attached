using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Threading;

namespace DB_attached
{
    public class WebSiteSRZ : IDisposable //класс для работы с сайтом СРЗ
    {        
        private WebProxy proxy;
        private CookieContainer cookieContainer;
        private HttpClientHandler handler;
        private HttpClient client;
        private FormUrlEncodedContent content;
        private HttpResponseMessage response;

        public WebSiteSRZ(string webAddress, string proxyAddress, int proxyPort)
        {
            cookieContainer = new CookieContainer();

            if (proxyAddress != "")
            {
                proxy = new WebProxy(proxyAddress + ":" + proxyPort.ToString());
                handler = new HttpClientHandler() { CookieContainer = cookieContainer, Proxy = proxy, UseProxy = true };
            }
            else
                handler = new HttpClientHandler() { CookieContainer = cookieContainer };

            client = new HttpClient(handler) { BaseAddress = new Uri(webAddress) };
        }

        public async Task<bool> Authorize(Credential credential)   //авторизуется на сайте
        {
            content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("lg", credential.Login),
                new KeyValuePair<string, string>("pw", credential.Password),
            });
            try
            {
                response = await client.PostAsync("data/user.ajax.logon.php", content);
                response.EnsureSuccessStatusCode();
                if (response.Content.ReadAsStringAsync().Result.Length > 5)
                    return false;
                else
                    return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public async Task<bool> Logout()   //выход с сайта
        {
            response = await client.GetAsync("?show=logoff");
            response.EnsureSuccessStatusCode();
            return true;
        }
        public Patient GetPatient(Patient patient)   //запрашивает пациента с сайта
        {
            try
            {
                content = new FormUrlEncodedContent(new[]
                {
                new KeyValuePair<string, string>("mode", "1"),
                new KeyValuePair<string, string>("person_enp", patient.polis),
            });

                response = client.PostAsync("data/reg.person.polis.search.php", content).Result;
                response.EnsureSuccessStatusCode();

                string result = response.Content.ReadAsStringAsync().Result;

                patient.fioFull = Converter.TextToFioFull(result);
                if (patient.fioFull == null)
                    return new Patient();
                patient.fioShort = Converter.TextToFioShort(result);
                return patient;
            }
            catch (Exception)
            {
                return new Patient();
            }
        }
        private class Job    //вспомогательный класс, описывает набор данных для 1 потока запросов к сайту: IDisposable
        {
            public int Nthread;
            public int Nrequests;
            public List<Credential> Accounts;
            public List<Patient> Patients;

            public Job()
            {
                Patients = new List<Patient>();
                Accounts = new List<Credential>();
            }
        }
        private static Job[] PrepareJobs(ConcurrentStack<string> patients, Credential[] Accounts, int nThreads)   //вспомогательный класс, равномерно делит входные данные на N потоков
        {
            if (patients.Count < nThreads)
            {
                nThreads = patients.Count / 5;
                if (nThreads < 1)
                    nThreads = 1;
            }

            Job[] jobs = new Job[nThreads];

            //Распределяем учетные записи и запросы равномерно
            int reqInThread = patients.Count / nThreads;
            int ostatok = patients.Count % nThreads;

            //Формируются данные для каждого потока
            for (int i = 0, k = 0, left; i < nThreads; i++)
            {
                jobs[i] = new Job();
                //Распределяем равномерно кол-во запросов для каждого потока
                jobs[i].Nthread = i;
                jobs[i].Nrequests = reqInThread;
                if (ostatok > 0)
                {
                    jobs[i].Nrequests++;
                    ostatok--;
                }
                //Формируем полисы
                string str;
                for (int j = 0; j < jobs[i].Nrequests; j++)
                {
                    patients.TryPop(out str);
                    jobs[i].Patients.Add(new Patient { polis = str });
                }

                //Формируем логины
                left = jobs[i].Nrequests;
                while (left != 0)
                {
                    //индекс меняется по кругу
                    if (Accounts.Count() == k)
                        k = 0;
                    if (Accounts[k] == null)
                    {
                        k++;
                        continue;
                    }

                    jobs[i].Accounts.Add(Accounts[k].Copy());
                    if (left >= Accounts[k].Requests)
                    {

                        left -= Accounts[k].Requests;
                        Accounts[k] = null;
                        k++;
                    }
                    else
                    {
                        jobs[i].Accounts[jobs[i].Accounts.Count() - 1].Requests = left;
                        Accounts[k].Requests -= left;
                        break;
                    }
                }
            }
            return jobs;
        }
        public async static Task<List<Patient>> GetPatients(ConcurrentStack<string> patients, Credential[] Accounts, int nThreads, Settings settings)   //запускает многопоточно запросы к сайту для поиска пациентов
        {
            return await Task.Run(() =>
            {
                Job[] jobs = WebSiteSRZ.PrepareJobs(patients, Accounts, nThreads);
                var result = new ConcurrentBag<Patient>();
                Task[] tasks = new Task[jobs.Count()];

                foreach (Job job in jobs)
                {
                    tasks[job.Nthread] = Task.Run(() =>
                    {
                        WebSiteSRZ site;
                        Patient patient;
                        int j = 0;

                        foreach (Credential cred in job.Accounts)
                        {
                            using (site = new WebSiteSRZ(settings.Site, settings.ProxyAddress, settings.ProxyPort))
                            {
                                for (int i = 0; i < 3; i++) //3 попытки на авторизацию с интервалом 10секунд
                                {
                                    if (site.Authorize(cred).Result)
                                    {
                                        while (cred.Requests > 0)
                                        {
                                            patient = site.GetPatient(job.Patients[j]);
                                            if (patient.polis != null)
                                            {
                                                result.Add(patient);
                                                cred.Requests--;
                                                j++;
                                            }
                                            else break;
                                        }
                                        site.Logout();
                                        break;
                                    }
                                    else
                                        Thread.Sleep(10000);
                                }
                            }
                        }
                    });
                }

                Task.WaitAll(tasks);

                return result.ToList<Patient>();
            });
        }
        public async Task<bool> DownloadFile(string excelFile, DateTime date)   //загружает файл прикрепленных пациентов на дату
        {
            //запрашиваем файл прикрепленных на дату
            content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("export_date_on", date.ToShortDateString().ToString()),
                new KeyValuePair<string, string>("exportlist_id", "25"),
            });

            response = await client.PostAsync("data/dbase.export.php", content);
            response.EnsureSuccessStatusCode();
            string resultText = response.Content.ReadAsStringAsync().Result;

            //загружаем zip архив в память
            response = await client.GetAsync(Converter.TextToRef(resultText));
            response.EnsureSuccessStatusCode();
            byte[] data = await response.Content.ReadAsByteArrayAsync();

            string dbfFile = Path.GetDirectoryName(excelFile) + "\\ATT_MO_temp_ERW3sdcxf1XCa.DBF";

            //извлекаем dbf файл из zip архива
            using (MemoryStream memoryStream = new MemoryStream(data))
            using (ZipArchive archive = new ZipArchive(memoryStream, ZipArchiveMode.Read))
                archive.Entries[0].ExtractToFile(dbfFile, true);
            

            //конвертируем dbf в excel
            Converter.DbfToExcel(dbfFile, excelFile);
            File.Delete(dbfFile);

            return true;
        }

        public void Dispose()
        {
            handler.Dispose();
            client.Dispose();
            content.Dispose();
            response.Dispose();
        }

    }


}
