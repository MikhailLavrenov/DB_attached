using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace DB_attached
{
    public partial class Form1 : Form
    {
        public Settings settings = new Settings();    //создаем экземпляр настроек программы
        public bool autoStart = false;  //true если запущено с ключем -start

        public Form1()
        {
            InitializeComponent();
            dateTimePicker1.Value = DateTime.Today.Date;
            checkBox2_CheckedChanged(new object(), new EventArgs());
            tabControl1.SelectedTab = Main;
            LoadSettings();   //загружаем настройки            

            String[] arguments = Environment.GetCommandLineArgs();  //проверяем ключи запуска
            foreach (string argument in arguments)
                if (argument.Substring(1) == "start")
                {
                    autoStart = true;
                    button6_Click(new object(), new EventArgs());
                    break;
                }
        }
      
        
        private void button3_Click(object sender, EventArgs e)  //Сохранить настройки 
        {
            SaveSettings();            
        }

        private void button4_Click(object sender, EventArgs e)   //Загрузить настройки
        {
            LoadSettings();
        }

        private async void button5_Click(object sender, EventArgs e)  //Тестировать настройки
        {
            try
            {
                DisableButtons();
                await TestSetting();
            }
            catch (Exception)
            {
                toolStripStatusLabel1.Text = "Произошла непредвиденная ошибка. Операция завершена некорректно.";
                throw;
            }
            finally
            {
                EnableButtons();
            }
        }

        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e) //Задает путь к папке с DBF файлом
        {
            if (comboBox2.SelectedIndex == 1)
            {
                FolderBrowserDialog folder = new FolderBrowserDialog();

                if (comboBox2.Text == "Текущая папка")
                    folder.SelectedPath = Directory.GetCurrentDirectory();
                else
                    folder.SelectedPath = comboBox2.Text;

                if (folder.ShowDialog() == DialogResult.OK)
                    BeginInvoke(new Action(() => comboBox2.Text = folder.SelectedPath + @"\"));
            }
        }

        private void comboBox1_DropDown(object sender, EventArgs e) //Задает имя DBF файла
        {
            string folder;
            if ((comboBox2.Text == "Текущая папка") || (comboBox2.Text == ""))
                folder = Directory.GetCurrentDirectory();
            else
                folder = comboBox2.Text;

            string[] dbfFiles = Directory.GetFiles(folder, "*.xlsx");
            comboBox1.Items.Clear();
            foreach (string file in dbfFiles)
                comboBox1.Items.Add(Path.GetFileName(file));
        }

        private async void button6_Click(object sender, EventArgs e)  //начать определение ФИО
        {

            try
            {
                DisableButtons();

                if (settings.TestPassed == false)
                    await TestSetting();

                if (settings.TestPassed)
                {
                    int maxReq = 0, madeReq = 0, step = 2;
                    string xlsxStatus;

                    if (settings.DownloadFile)
                    {
                        toolStripStatusLabel1.Text = string.Format("{0}. Выполняется загрузка файла из СРЗ, ждите...", step++);
                        using (WebSiteSRZ site = new WebSiteSRZ(settings.Site, settings.ProxyAddress, settings.ProxyPort))
                        {
                            await site.Authorize(settings.Accounts[0]);
                            await site.DownloadFile(settings.Path, dateTimePicker1.Value);
                            site.Logout();
                        }
                    }
                    toolStripStatusLabel1.Text = string.Format("{0}. Выполняется обработка из кэша, ждите...", step++);
                    foreach (Credential cred in settings.Accounts)
                        maxReq += cred.Requests;
                    using (CacheDB cacheDB = new CacheDB())
                    using (var excel = new ExcelFile())
                    {
                        await excel.Open(settings.Path, settings.ColumnSynonims);
                        ConcurrentStack<string> listJobs = await excel.GetPatientsFromCache(cacheDB, maxReq, true);
                        if (listJobs.Count > 0)
                        {
                            toolStripStatusLabel1.Text = string.Format("{0}. Выполняется поиск пациентов в СРЗ, ждите...", step++);
                            List<Patient> patients = await WebSiteSRZ.GetPatients(listJobs, settings.CopyAccounts(), settings.Threads, settings);
                            madeReq = patients.Count;
                            toolStripStatusLabel1.Text = string.Format("{0}. Выполняется добавление в кэш, ждите...", step++);
                            await cacheDB.AddPatients(patients, true);
                            toolStripStatusLabel1.Text = string.Format("{0}. Выполняется обработка из кэша, ждите...", step++);
                            listJobs = await excel.GetPatientsFromCache(cacheDB, int.MaxValue);
                        }
                        if (listJobs.Count == 0)
                        {
                            if (settings.RenameGender)
                                await excel.RenameSex();
                            if (settings.RenameColumnNames)
                                await excel.ProcessColumns();
                            if (settings.ColumnOrder)
                                await excel.SetColumnsOrder();
                            if (settings.ColumnAutoWidth)
                                await excel.FitColumnWidth();
                            if (settings.AutoFilter)
                                await excel.AutoFilter();

                            xlsxStatus = "Файл готов, найдены все ФИО.";
                        }
                        else xlsxStatus = String.Format("Файл не готов, осталось запросить в СРЗ {0} ФИО.", listJobs.Count);
                        await excel.Save();
                        toolStripStatusLabel1.Text = String.Format("Готово. Запрошено пациентов в СРЗ: {0} из {1} разрешенных. {2}", madeReq, maxReq, xlsxStatus);

                        if (autoStart)
                            Environment.Exit(0);    //выход с кодом 0 если запущено с командной строки
                    }
                }
            }
            catch (Exception)
            {
                toolStripStatusLabel1.Text = "Произошла непредвиденная ошибка. Операция завершена некорректно.";
                if (autoStart)
                    Environment.Exit(1);    //выход с кодом 1 если запущено с командной строки
                throw;
            }
            finally
            {
                EnableButtons();
            }
        }


        private void button7_Click(object sender, EventArgs e)  //Удалить внутреннюю БД (кэш ФИО)
        {
            if (File.Exists("CacheDB.sdf"))
            {
                toolStripStatusLabel1.Text = "Выполняется, ждите...";
                if (DialogResult.Yes == MessageBox.Show("Это действие приведет к потере всех кэшированных ФИО! Продолжить?", "Предупреждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                {
                    File.Delete("CacheDB.sdf");
                    toolStripStatusLabel1.Text = "Готово. БД кэша удалена.";
                }
                else
                    toolStripStatusLabel1.Text = "Отменено.";
            }
            else
                toolStripStatusLabel1.Text = "БД кэша отсутствует. Удаление не возможно.";


        }

        private async void button2_Click(object sender, EventArgs e)    //Добавить в кэш из xlsx
        {
            try
            {
                DisableButtons();

                toolStripStatusLabel1.Text = "Выполняется, ждите...";
                OpenFileDialog fileDialog = new OpenFileDialog();
                fileDialog.InitialDirectory = Path.GetDirectoryName(settings.Path);
                fileDialog.Filter = "xlsx files (*.xslx)|*.xlsx";

                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    bool overrideMode = (DialogResult.Yes == MessageBox.Show("Загружаемые данные новее чем в кэше? \r\nЕсли не уверены, нажмите \"Нет\"", "Вопрос", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2));

                    using (CacheDB cacheDB = new CacheDB())
                    using (ExcelFile excel = new ExcelFile())
                    {
                        await excel.Open(fileDialog.FileName);
                        var patients = await excel.ToList();
                        int n = await cacheDB.AddPatients(patients, overrideMode);
                        toolStripStatusLabel1.Text = string.Format("Готово. В кэш добавлено {0} записи(ей).", n);
                    }
                }
                else
                    toolStripStatusLabel1.Text = "Отменено.";
                EnableButtons();
            }
            catch (Exception)
            {
                toolStripStatusLabel1.Text = "Произошла непредвиденная ошибка. Операция завершена некорректно.";
                throw;
            }
            finally
            {
                EnableButtons();
            }

        }

        private void button9_Click(object sender, EventArgs e)  //Настройки по умолчанию
        {


            textBox1.Text = @"http://11.0.0.1/";
            textBox2.Text = "";
            textBox5.Text = "0";
            comboBox2.SelectedIndex = 0;
            comboBox1.Text = "Прикрепленные пациенты (Выгрузка из СРЗ).XLSX";
            textBox4.Text = "10";
            radioButton1.Checked = true;
            comboBox3.SelectedIndex = 0;
            checkBox2.Checked = true;
            dataGridView1.Rows.Clear();
            dataGridView1.Rows.Add();
            dataGridView1.Rows[0].Cells["Login"].Value = "МойЛогин1";
            dataGridView1.Rows[0].Cells["Password"].Value = "МойПароль1";
            dataGridView1.Rows[0].Cells["Limit"].Value = "400";
            dataGridView1.Rows.Add();
            dataGridView1.Rows[1].Cells["Login"].Value = "МойЛогин2";
            dataGridView1.Rows[1].Cells["Password"].Value = "МойПароль2";
            dataGridView1.Rows[1].Cells["Limit"].Value = "300";
            dataGridView1.Focus();
            toolStripStatusLabel1.Text = "Готово. Настройки по умолчанию добавлены. Не забудьте сохранить.";
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            dataGridView1.CurrentRow.Cells["Password"].Value = textBox3.Text;
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Cells["Password"].Value != null)
                textBox3.Text = dataGridView1.CurrentRow.Cells["Password"].Value.ToString();
            else
                textBox3.Text = string.Empty;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                textBox3.Visible = true;
                label7.Visible = true;
                dataGridView1.Columns["Password"].Visible = false;
            }
            else
            {
                textBox3.Visible = false;
                label7.Visible = false;
                dataGridView1.Columns["Password"].Visible = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)  //Оптимизация БД кэша
        {
            if (File.Exists("CacheDB.sdf"))
            {
                toolStripStatusLabel1.Text = "Выполняется, ждите...";
                using (CacheDB cacheDB = new CacheDB())
                    cacheDB.Optimize();
                toolStripStatusLabel1.Text = "Готово. БД кэша оптимизрована.";

            }
            else
                toolStripStatusLabel1.Text = "БД кэша отсутствует. Оптимизация не требуется.";
        }

        private void button8_Click(object sender, EventArgs e)  //Сохранить настройки
        {
            SaveSettings();
        }

        private void button10_Click(object sender, EventArgs e) //Загрузить настройки
        {
            LoadSettings();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count > 1)
                if ((dataGridView2.CurrentRow.Index > 0) && (dataGridView2.CurrentRow.Index < dataGridView2.Rows.Count - 1))
                {

                    int row = dataGridView2.CurrentCell.RowIndex;
                    int col = dataGridView2.CurrentCell.ColumnIndex;
                    DataGridViewRow curRow = dataGridView2.Rows[row];

                    dataGridView2.Rows.RemoveAt(row);
                    dataGridView2.Rows.Insert(row - 1, curRow);
                    dataGridView2.CurrentCell = dataGridView2[col, row - 1];
                }
        }
        private void button12_Click(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count > 1)
                if (dataGridView2.CurrentRow.Index < dataGridView2.Rows.Count - 2)
                {

                    int row = dataGridView2.CurrentCell.RowIndex;
                    int col = dataGridView2.CurrentCell.ColumnIndex;
                    DataGridViewRow curRow = dataGridView2.Rows[row];

                    dataGridView2.Rows.RemoveAt(row);
                    dataGridView2.Rows.Insert(row + 1, curRow);
                    dataGridView2.CurrentCell = dataGridView2[col, row + 1];
                }
        }
        public async Task<bool> TestSetting()   //Тестирование настроек
        {

            toolStripStatusLabel1.Text = "1. Выполняется: тестирование настроек, ждите...";

            SetDefaultColors();

            ConcurrentDictionary<string, bool> errors = await settings.Test();
            if (errors.Values.Contains(true))
            {
                if (autoStart)
                    Environment.Exit(1);  //выход с кодом 1 если запущено с командной строки

                toolStripStatusLabel1.Text = "Готово. Тест настроек не пройден. Проверьте настройки.";
                tabControl1.SelectedTab = Configuration;

                if (errors["proxy"])
                {
                    textBox2.BackColor = Color.Red;
                    textBox5.BackColor = Color.Red;
                }
                if (errors["site"])
                    textBox1.BackColor = Color.Red;
                if (errors["folder"])
                    comboBox2.BackColor = Color.Red;
                if (errors["file"])
                    comboBox1.BackColor = Color.Red;
                if (errors["threads"])
                    textBox1.BackColor = Color.Red;

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                    if (errors.ContainsKey(dataGridView1.Rows[i].Cells["Login"].Value.ToString()))
                        if (errors[dataGridView1.Rows[i].Cells["Login"].Value.ToString()] == true)
                            dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Red;
            }
            else toolStripStatusLabel1.Text = "Готово. Тест настроек пройден успешно.";
            return true;
        }

        public void SaveSettings() //Сохранение настроек
        {
            toolStripStatusLabel1.Text = "Выполняется, ждите...";

            SetDefaultColors();

            //В не заполненные поля устанавливаем значения по умолчанию
            if (!textBox1.Text.StartsWith("http://"))
                textBox1.Text = @"http://" + textBox1.Text;
            if (textBox1.Text[textBox1.Text.Length - 1] != '/')
                textBox1.Text += @"/";
            if (textBox5.Text == "")
                textBox5.Text = "0";
            if (textBox4.Text == "")
                textBox4.Text = "10";
            if (comboBox2.Text == "")
                comboBox2.SelectedIndex = 0;
            if ((comboBox2.Text != "Текущая папка") && (comboBox2.Text[comboBox2.Text.Length - 1] != '\\'))
                comboBox2.Text += @"\";
            if (comboBox1.Text == "")
                comboBox1.Text = "Прикрепленные пациенты (Выгрузка из СРЗ).XLSX";
            if ((radioButton1.Checked == false) && (radioButton2.Checked == false))
                radioButton1.Checked = true;
            if (comboBox3.SelectedIndex == -1)
                comboBox3.SelectedIndex = 0;

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)    //проверяем заполненность таблицы с учетными записями
                for (int j = 0; j < 3; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value == null)
                    {
                        dataGridView1.Rows.RemoveAt(i);
                        break;
                    }

            for (int i = 0; i < dataGridView2.RowCount - 1; i++)    //проверяем заполненность таблицы альтернативных названий стобцов
            {
                if (dataGridView2.Rows[i].Cells["OldName"].Value == null)
                {
                    dataGridView2.Rows.RemoveAt(i);
                    continue;
                }
                if ((dataGridView2.Rows[i].Cells["NewName"].Value == null) || (dataGridView2.Rows[i].Cells["NewName"].Value.ToString() == ""))
                    dataGridView2.Rows[i].Cells["NewName"].Value = dataGridView2.Rows[i].Cells["OldName"].Value;
                if (dataGridView2.Rows[i].Cells["HideCol"].Value == null)
                    dataGridView2.Rows[i].Cells["HideCol"].Value = false;
                if (dataGridView2.Rows[i].Cells["Delete"].Value == null)
                    dataGridView2.Rows[i].Cells["Delete"].Value = false;
            }

            //Записываем настройки в класс настроек
            settings.Site = textBox1.Text;
            settings.ProxyAddress = textBox2.Text;
            settings.ProxyPort = int.Parse(textBox5.Text);
            settings.Folder = comboBox2.Text;
            settings.File = comboBox1.Text;
            settings.DownloadFile = radioButton1.Checked;
            settings.Threads = int.Parse(textBox4.Text);
            settings.EncryptLevel = comboBox3.SelectedIndex;
            settings.hidePassword = checkBox2.Checked;
            settings.RenameGender = checkBox1.Checked;
            settings.RenameColumnNames = checkBox3.Checked;
            settings.ColumnAutoWidth = checkBox4.Checked;
            settings.AutoFilter = checkBox5.Checked;
            settings.ColumnOrder = checkBox6.Checked;

            Credential[] creds = new Credential[dataGridView1.RowCount - 1];
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                creds[i] = new Credential();
                creds[i].SetLogin(dataGridView1.Rows[i].Cells["Login"].Value.ToString(), settings.EncryptLevel);
                creds[i].SetPassword(dataGridView1.Rows[i].Cells["Password"].Value.ToString(), settings.EncryptLevel);
                creds[i].Requests = int.Parse(dataGridView1.Rows[i].Cells["Limit"].Value.ToString());
            }
            settings.Accounts = creds;

            ColumnSynonim[] colSyn = new ColumnSynonim[dataGridView2.RowCount - 1];
            for (int i = 0; i < dataGridView2.RowCount - 1; i++)
            {
                colSyn[i].name = dataGridView2.Rows[i].Cells["OldName"].Value.ToString();
                colSyn[i].altName = dataGridView2.Rows[i].Cells["NewName"].Value.ToString();
                colSyn[i].hide = (Boolean)dataGridView2.Rows[i].Cells["HideCol"].Value;
                colSyn[i].delete = (Boolean)dataGridView2.Rows[i].Cells["Delete"].Value;
            }
            settings.ColumnSynonims = colSyn;

            settings.SaveSettings();
            toolStripStatusLabel1.Text = "Настройки успешно сохранены.";
        }

        public void LoadSettings()  //Загрузка настроек
        {
            if (File.Exists("Settings.xml"))
            {
                toolStripStatusLabel1.Text = "Выполняется, ждите...";
                settings = Settings.LoadSettings();

                textBox1.Text = settings.Site;
                textBox2.Text = settings.ProxyAddress;
                textBox5.Text = settings.ProxyPort.ToString();
                comboBox2.Text = settings.Folder;
                comboBox1.Text = settings.File;
                radioButton1.Checked = settings.DownloadFile;
                radioButton2.Checked = !radioButton1.Checked;
                textBox4.Text = settings.Threads.ToString();
                comboBox3.SelectedIndex = settings.EncryptLevel;
                checkBox2.Checked = settings.hidePassword;
                checkBox1.Checked = settings.RenameGender;
                checkBox3.Checked = settings.RenameColumnNames;
                checkBox4.Checked = settings.ColumnAutoWidth;
                checkBox5.Checked = settings.AutoFilter;
                checkBox6.Checked = settings.ColumnOrder;
                settings.TestPassed = false;

                dataGridView1.Rows.Clear();
                dataGridView1.RowCount = settings.Accounts.Count() + 1;
                for (int i = 0; i < settings.Accounts.Count(); i++)
                {
                    settings.Accounts[i].GenerateDecryptedCredential(settings.EncryptLevel);
                    dataGridView1.Rows[i].Cells["Login"].Value = settings.Accounts[i].Login;
                    dataGridView1.Rows[i].Cells["Password"].Value = settings.Accounts[i].Password;
                    dataGridView1.Rows[i].Cells["Limit"].Value = settings.Accounts[i].Requests;

                }
                dataGridView1.Focus();

                dataGridView2.Rows.Clear();
                dataGridView2.RowCount = settings.ColumnSynonims.Count() + 1;
                for (int i = 0; i < settings.ColumnSynonims.Count(); i++)
                {
                    dataGridView2.Rows[i].Cells["OldName"].Value = settings.ColumnSynonims[i].name;
                    dataGridView2.Rows[i].Cells["NewName"].Value = settings.ColumnSynonims[i].altName;
                    dataGridView2.Rows[i].Cells["HideCol"].Value = settings.ColumnSynonims[i].hide;
                    dataGridView2.Rows[i].Cells["Delete"].Value = settings.ColumnSynonims[i].delete;
                }
                dataGridView2.Focus();

                toolStripStatusLabel1.Text = "Настройки успешно загружены.";
            }
            else
            {
                toolStripStatusLabel1.Text = "Невозможно отменить. Предыдущие настройки не были сохранены.";
            }

        }
        private void DisableButtons()
        {
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;
            button9.Enabled = false;
            button10.Enabled = false;
            tabControl1.Focus();
            
        }
        private void EnableButtons()
        {
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = true;
            button9.Enabled = true;
            button10.Enabled = true;
        }
        private void SetDefaultColors()
        {
            //Устанавливаем стандартный цвет
            textBox1.BackColor = default(Color);
            textBox2.BackColor = default(Color);
            textBox4.BackColor = default(Color);
            textBox5.BackColor = default(Color);
            comboBox1.BackColor = default(Color);
            comboBox2.BackColor = default(Color);
            dataGridView1.DefaultCellStyle.BackColor = default(Color);
            foreach (DataGridViewRow row in dataGridView1.Rows)
                row.DefaultCellStyle.BackColor = Color.Empty;

        }
    }
}

