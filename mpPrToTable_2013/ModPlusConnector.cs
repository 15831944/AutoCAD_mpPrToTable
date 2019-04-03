using System;
using System.Collections.Generic;
using ModPlusAPI.Interfaces;

namespace mpPrToTable
{
    public class ModPlusConnector : IModPlusFunctionInterface
    {
        public SupportedProduct SupportedProduct => SupportedProduct.AutoCAD;
        public string Name => "mpPrToTable";
#if A2013
        public string AvailProductExternalVersion => "2013";
#elif A2014
        public string AvailProductExternalVersion => "2014";
#elif A2015
        public string AvailProductExternalVersion => "2015";
#elif A2016
        public string AvailProductExternalVersion => "2016";
#elif A2017
        public string AvailProductExternalVersion => "2017";
#elif A2018
        public string AvailProductExternalVersion => "2018";
#elif A2019
        public string AvailProductExternalVersion => "2019";
#elif A2020
        public string AvailProductExternalVersion => "2020";
#endif
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
