using System.Data;
using System.IO;
using System.Text;
using System.Data.OleDb;
using OfficeOpenXml;
using System;

namespace DB_attached
{
    public class Converter   //класс описывает различные конвертеры
    {
        public static void DbfToExcel(string dbfFile, string excelFile)  //преобразует dbf в excel
        {
            //устанавливаем байт с кодовой страницей 866 чтобы драйвер читал правильно
            using (FileStream fs = new FileStream(dbfFile, FileMode.Open))
            {
                fs.Seek(29, SeekOrigin.Begin);
                if (fs.ReadByte() != 101)
                {
                    fs.Seek(-1, SeekOrigin.Current);
                    fs.WriteByte(101);
                }
            }


            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.GetDirectoryName(dbfFile) + @";Extended Properties = dBASE IV; User ID = Admin; Password =; ";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            using (ExcelPackage package = new ExcelPackage())
            {
                var sqlString = "select * from " + Path.GetFileName(dbfFile);
                OleDbCommand command = new OleDbCommand(sqlString, connection);
                connection.Open();

                DataSet ds = new DataSet();
                OleDbDataAdapter da = new OleDbDataAdapter(command);
                da.Fill(ds);
                

                ExcelWorksheet ws = package.Workbook.Worksheets.Add("Лист1");

                for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                    ws.Cells[1, j + 1].Value = ds.Tables[0].Columns[j].ColumnName.ToString();

                DateTime datetime;
                string str;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                    {
                        str = ds.Tables[0].Rows[i][j].ToString();

                        if (ds.Tables[0].Columns[j].DataType.Name == "DateTime")    //отбрасываем время из формата DateTime
                            if (DateTime.TryParse(str, out datetime))
                                str = datetime.ToShortDateString();

                        ws.Cells[i + 2, j + 1].Value = str;
                    }

                package.SaveAs(new FileInfo(excelFile));
            }
        }
        public static string FioFullToShort(string fioFull)  //преобразует полные фио в 2-3 буквы ФИО
        {
            int space1, space2;
            int n = 3;

            space1 = fioFull.IndexOf(' ', 0) + 1;
            if (space1 == 0)
                return "";
            space2 = fioFull.IndexOf(' ', space1) + 1;

            if (space2 <= space1)
                n=2;

            if (n == 3)
                return string.Join("", fioFull[0], fioFull[space1], fioFull[space2]);
            else
                return string.Join("", fioFull[0], fioFull[space1]);
        }
        public static string TextToRef(string text)  //проебразует ответ web сервера в ссылку на файл
        {
            int begin, end;
            begin = text.IndexOf(@"href='");
            begin = text.IndexOf(@"'", begin) + 1;
            end = text.IndexOf(@"' ", begin);

            return text.Substring(begin, end - begin);
        }
        public static string TextToFioFull(string text)  //преобразует ответ web сервера в полные ФИО
        {
            try
            {
                int pos1, pos2;
                StringBuilder res = new StringBuilder("");

                pos1 = 0;
                for (int i = 0; i < 3; i++)
                    pos1 = text.IndexOf("||", pos1) + 2;

                for (int i = 0; i < 3; i++)
                {
                    pos2 = text.IndexOf("||", pos1);
                    if (pos2 != pos1)
                        res.Append(text.Substring(pos1, pos2 - pos1)).Append(' ');
                    pos1 = pos2 + 2;
                }
                res.Remove(res.Length - 1, 1);

                return res.ToString().ToUpper();
            }
            catch (Exception)
            {
                return null;
            }
        }
        public static string TextToFioShort(string text)  //преобразует ответ web сервера в 2-3 буквы ФИО
        {
            int pos1, pos2;
            StringBuilder res = new StringBuilder("");

            pos1 = 0;
            for (int i = 0; i < 3; i++)
                pos1 = text.IndexOf("||", pos1) + 2;

            for (int i = 0; i < 3; i++)
            {
                pos2 = text.IndexOf("||", pos1);
                if (pos2 != pos1)
                    res.Append(text.Substring(pos1, 1));
                pos1 = pos2 + 2;
            }

            return res.ToString().ToUpper();
        }
        public static ColumnSynonim GetAltName(string name, ColumnSynonim[] columnSynonims)  //возвращает альтернативное название столбца, если синонима нет возвращает это же название
        {
            foreach (ColumnSynonim cs in columnSynonims)
            {
                if ((cs.name == name) || (cs.altName == name))
                    return cs;
            }
            
            return new ColumnSynonim {name=name, altName=name,hide=false,delete=false};
        }
    }
}
