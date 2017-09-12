using System;
using System.Collections.Generic;
using ModPlusAPI.Interfaces;

namespace mpPrToTable
{
    public class Interface : IModPlusFunctionInterface
    {
        public SupportedProduct SupportedProduct => SupportedProduct.AutoCAD;
        public string Name => "mpPrToTable";
        public string AvailProductExternalVersion => "2013";
        public string FullClassName => string.Empty;
        public string AppFullClassName => string.Empty;
        public Guid AddInId => Guid.Empty;
        public string LName => "Изделия в таблицу";
        public string Description => "Функция позволяет заполнить таблицу спецификации выбранными изделиями ModPlus";
        public string Author => "Пекшев Александр aka Modis";
        public string Price => "0";
        public bool CanAddToRibbon => true;
        public string FullDescription => "Функция собирает данные для спецификации из расширенных данных продуктов (блоки, созданные функцией \"Вставить изделие\", или примитивы AutoCAD с расширенными данными, добавленными функцией \"Вставить изделие\"), а также из блоков, имеющих атрибуты для спецификации. Имеется возможность указать строку, с которой начнется заполнение таблицы";
        public string ToolTipHelpImage => string.Empty;
        public List<string> SubFunctionsNames => new List<string>();
        public List<string> SubFunctionsLames => new List<string>();
        public List<string> SubDescriptions => new List<string>();
        public List<string> SubFullDescriptions => new List<string>();
        public List<string> SubHelpImages => new List<string>();
        public List<string> SubClassNames => new List<string>();
    }
}
