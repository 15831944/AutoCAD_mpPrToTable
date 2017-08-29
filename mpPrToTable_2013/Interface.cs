using mpPInterface;

namespace mpPrToTable
{
    public class Interface : IPluginInterface
    {
        public string Name => "mpPrToTable";
        public string AvailCad => "2013";
        public string LName => "Изделия в таблицу";
        public string Description => "Функция позволяет заполнить таблицу спецификации выбранными изделиями ModPlus";
        public string Author => "Пекшев Александр aka Modis";
        public string Price => "0";
    }
}
