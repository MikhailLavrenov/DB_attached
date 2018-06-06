using System;

namespace DB_attached
{
    public struct Patient   //Структура данных пациента
    {
        public string fioShort;
        public string fioFull;
        public string polis;
    }

    [Serializable]
    public struct ColumnSynonim    //Структура альтернативного названия столбца
    {
        public string name;
        public string altName;
        public bool hide;
        public bool delete;
    }
}
