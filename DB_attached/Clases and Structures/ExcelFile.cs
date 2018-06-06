using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;


namespace DB_attached
{
    public class ExcelFile:IDisposable  //класс для работы с excel файлом
    {
        private string file;
        private ExcelPackage package;
        private ExcelWorksheet ws;
        private int maxRow;
        private int maxCol;
        private ColumnSynonim[] columnSynonims;

        public async Task<bool> Open(string file, ColumnSynonim[] columnSynonims=null)  //открывает xlsx файл
        {
            return await Task.Run(() =>
            {
                this.file = file;
                this.package = new ExcelPackage(new FileInfo(this.file));
                this.ws = package.Workbook.Worksheets[1];
                this.maxRow = ws.Dimension.Rows;
                this.maxCol = ws.Dimension.Columns;
                this.columnSynonims = columnSynonims;
                return false;
            });
        }
        public async Task<bool> Save()
        {
            return await Task.Run(() =>
            {
                package.Save();
                return true;
            });
        }   //сохраняет xlsx файл
        private int FindColumnIndex(string text, int rowIndex)  //ищет в заданной строке текст 
        {
            string str;
            string textAlt = Converter.GetAltName(text, columnSynonims).altName;

            for (int i = 1; i <= maxCol; i++)
            {
                if (ws.Cells[1, i].Value == null)
                    continue;

                str = ws.Cells[1, i].Value.ToString();
                if ((str == text) || (str == textAlt))   //поиск столбца с № полиса
                    return i;
            }
            return -1;
        }
        private int FindColumnIndex(ColumnSynonim text, int rowIndex)  //ищет в заданной строке текст 
        {
            string str;

            for (int i = 1; i <= maxCol; i++)
            {
                if (ws.Cells[1, i].Value == null)
                    continue;

                str = ws.Cells[1, i].Value.ToString();
                if ((str == text.name) || (str == text.altName))   //поиск столбца с № полиса
                    return i;
            }
            return -1;
        }
        public async Task<bool> AutoFilter()    //добавляет фильтр
        {
            return await Task.Run(() =>
            {
                ws.Cells[ws.Dimension.Address].AutoFilter = true;                
                return false;
            });
        }
        public async Task<bool> FitColumnWidth()    //подстраивает ширину столбцов под содержимое
        {
            return await Task.Run(() =>
            {
                ws.Cells.AutoFitColumns();
                return false;
            });
        }
        public async Task<bool> SetColumnsOrder()    //изменяет порядок столбоц
        {
            return await Task.Run(() =>
            {
                int insColPos = 1;
                int foundColPos;

                foreach(var columnSynonim in columnSynonims)
                {
                    foundColPos = FindColumnIndex(columnSynonim, 1);

                    if (foundColPos == insColPos)
                        insColPos++;
                    else if (foundColPos != -1)
                    {
                        ws.InsertColumn(insColPos, 1);
                        foundColPos++;
                        ws.Cells[1, foundColPos, maxRow, foundColPos].Copy(ws.Cells[1, insColPos, maxRow, insColPos]);
                        ws.DeleteColumn(foundColPos);
                        insColPos++;
                    }
                }
                return false;
            });
        }
        public async Task<bool> RenameSex()   //переименовывает цифры с полом в нормальные названия
        {
            return await Task.Run(() =>
            {
                string str;
                int columnSex = FindColumnIndex("SEX", 1);

                if (columnSex != -1)
                    for (int i = 1; i <= maxRow; i++)
                    {
                        if (ws.Cells[i, columnSex].Value == null)
                            continue;

                        str = ws.Cells[i, columnSex].Value.ToString();
                        if (str == "1")
                            ws.Cells[i, columnSex].Value = "Мужской";
                        else if (str == "2")
                            ws.Cells[i, columnSex].Value = "Женский";
                    }
                return false;
            });
        }
        public async Task<bool> ProcessColumns()    //переименовывает названия столбцов в нормальные названия, скрывает и удаляет столбцы
        {
            return await Task.Run(() =>
            {
                string name;
                ColumnSynonim synonim;

                for (int i = 1; i <= maxCol; i++)
                {
                    if (ws.Cells[1, i].Value == null)
                        continue;

                    name = ws.Cells[1,i].Value.ToString();
                    synonim = Converter.GetAltName(name, columnSynonims);
                    if (name != synonim.altName)
                        ws.Cells[1, i].Value = synonim.altName;

                    ws.Column(i).Hidden = synonim.hide;
                    if (synonim.delete)
                    {
                        ws.DeleteColumn(i);
                        maxCol--;
                        i--;
                    }
                }

                return false;
            });
        }
        public async Task<ConcurrentStack<string>> GetPatientsFromCache(CacheDB cacheDB, int stackLimitCount, bool stopIfStackLimitReached=false) //определяет полные ФИО из БД кэша
        {
            return await Task.Run(() =>
            {
                var patients = new ConcurrentStack<string>();

                int polisColumn = FindColumnIndex("ENP", 1);
                int fioColumn = FindColumnIndex("FIO", 1);

                if ((polisColumn==-1) || (fioColumn==-1))
                    return patients;

                //поиск полного ФИО или добавление в лист обработки
                cacheDB.PrepareGetPatients();
                Parallel.For(2, maxRow + 1, (i, state) =>
                {
                    lock(this)
                        if (ws.Cells[i, polisColumn].Value == null)
                            return;

                    Patient patient;
                    lock (this)
                        patient.fioShort = ws.Cells[i, fioColumn].Value.ToString();
                    if (patient.fioShort.Length > 3)  //если ФИО уже полные
                        return;

                    patient.fioFull = "";
                    lock (this)
                        patient.polis = ws.Cells[i, polisColumn].Value.ToString();
                    patient.fioFull = cacheDB.GetPatient(patient).fioFull;

                    if (patient.fioFull != null)  //если фио найдено в кэше
                        lock (this)
                            ws.Cells[i, fioColumn].Value = patient.fioFull;
                    else if (patients.Count < stackLimitCount)   //если фио не найдено в кэше добавляем в лист поиска через сайт
                        patients.Push(patient.polis);
                    else if (stopIfStackLimitReached)
                        state.Break();

                });

                while (patients.Count > stackLimitCount)    //из-за асинхронного выполнения, размер стэка может получиться больше чем надо, выкидываем лишнее
                    patients.TryPop(out string _);
                return patients;
            });
        }
        public async Task<List<Patient>> ToList()   //проеобразует данные excel в список пациентов
        {
            return await Task.Run(() =>
            {

                List<Patient> patients = new List<Patient>();
                Patient patient;

                for (int i = 1; i < maxRow; i++)
                {
                    if (ws.Cells[i, 1].Value == null)
                        continue;

                    patient.polis = ws.Cells[i, 1].Value.ToString();
                    while (patient.polis.Contains(' '))
                        patient.polis = patient.polis.Replace(" ", "");


                    if (patients.Find(Patient => Patient.polis == patient.polis).polis == patient.polis)
                        continue;

                    patient.fioFull = ws.Cells[i, 2].Value.ToString().Trim().ToUpper();
                    while (patient.fioFull.Contains("  "))
                        patient.fioFull = patient.fioFull.Replace("  ", " ");

                    patient.fioShort = Converter.FioFullToShort(patient.fioFull);
                    if (patient.fioShort == "")
                        continue;
                    patients.Add(patient);
                }
                return patients;
            });
        }
        public void Dispose()
        {
            if (ws!=null)
                ws.Dispose();
            if (package != null)
                package.Dispose();
        }
    }
}
