using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Data.SqlServerCe;

namespace DB_attached
{

    public class CacheDB : IDisposable    //класс для работы с внутренней БД (используется для хранения найденных ФИО пациентов) 
    {
        private string connectionString;
        private SqlCeCommand command;
        private SqlCeConnection connection;
        private Patient[] cachedPatients;

        public CacheDB()
        {
            connectionString = "DataSource=\"CacheDB.sdf\"; Password=\"MyPassword\"";
            command = new SqlCeCommand();
            if (!File.Exists("CacheDB.sdf"))
            {
                CreateDB();
                connection = new SqlCeConnection(connectionString);
                connection.Open();
                CreateTable();
            }
            else
            {
                connection = new SqlCeConnection(connectionString);
                connection.Open();
            }
        }
        public void CreateDB()  //создает базу данных
        {
            using (SqlCeEngine engine = new SqlCeEngine(connectionString))
                engine.CreateDatabase();
        }
        public void CreateTable()  //создает таблицу кэша пациентов
        {
            command = new SqlCeCommand(@" create table patients (id int primary key identity(1,1), polis nvarchar(20) not null UNIQUE, fio_short nvarchar(3) not null,fio_full nvarchar(50) null)", connection);
            command.ExecuteNonQuery();
        }
        public void Optimize()  //сжимает и перестаривает индексы
        {
            connection.Close();
            using (SqlCeEngine engine = new SqlCeEngine(connectionString))
                engine.Compact(null);
            connection.Open();            
        }
        public void PrepareGetPatients()  //подготавливает массив пациентов для быстрого поиска, искать по БД медленно
        {
            command = new SqlCeCommand(@" select polis, fio_short, fio_full  from patients", connection);
            using (var da = new SqlCeDataAdapter(command))
            using (var dt = new DataTable())
            {
                da.Fill(dt);

                cachedPatients = new Patient[dt.Rows.Count];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cachedPatients[i].polis = dt.Rows[i]["polis"].ToString();
                    cachedPatients[i].fioShort = dt.Rows[i]["fio_short"].ToString();
                    cachedPatients[i].fioFull = dt.Rows[i]["fio_full"].ToString();
                }
            }
        }
        public void AddPatient(Patient patient)  //добавляет пациента в БД
        {
            command = new SqlCeCommand(string.Format(@" insert into patients (polis,fio_short, fio_full) values ('{0}','{1}','{2}') ", patient.polis, patient.fioShort, patient.fioFull), connection);
            command.ExecuteNonQuery();            
        }
        public Patient GetPatient(Patient patient)  //ищет пациента в предварительно загруженном массиве
        {
            for (int i = 0; i < cachedPatients.Count(); i++)
                if (cachedPatients[i].polis == patient.polis)
                    if ((patient.fioShort == cachedPatients[i].fioShort) || (patient.fioShort == patient.fioFull /*выполнимо если оба поля NULL*/) || (patient.fioFull == cachedPatients[i].fioFull))
                        return cachedPatients[i];

            return new Patient();

        }
        public void DeletePatient(string polis)  //удаляет пациента из БД
        {
            command = new SqlCeCommand(string.Format(@" delete from patients where polis like '{0}'",polis), connection);
            command.ExecuteNonQuery();
        }
        public async Task<int> AddPatients(List<Patient> addPatients, bool mode) //Добавляет в БД список пациентов, mode = false добавляются пациенты с новыми полюсами, mode=true как предыдщуй + если не совпало ФИО запись перезаписывается
        {
            return await Task.Run(() =>
            {
                this.PrepareGetPatients();
                Patient patient;
                List<string> addedRecords = new List<string>();

                for (int i = 0; i < addPatients.Count; i++)
                {
                    if (addedRecords.Contains(addPatients[i].polis))  //защита от записи в кэш одинаковых пациентов
                        continue;

                    patient = this.GetPatient(new Patient { polis = addPatients[i].polis });

                    if (patient.polis == null)
                    {
                        this.AddPatient(addPatients[i]);
                        addedRecords.Add(addPatients[i].polis);
                    }
                    else if ((mode) && ((patient.fioFull != addPatients[i].fioFull) || (patient.fioShort != addPatients[i].fioShort)))
                    {
                        this.DeletePatient(addPatients[i].polis);
                        this.AddPatient(addPatients[i]);
                        addedRecords.Add(addPatients[i].polis);
                    }
                }
                return addedRecords.Count;
            });
        }

        public void Dispose()
        {
            connection.Close();
            connection.Dispose();
            command.Dispose();
        }
    }


}
