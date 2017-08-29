#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using mpMsg;
using mpProductInt;
using ModPlus;
using Visibility = System.Windows.Visibility;

namespace mpPrToTable
{
    public class MpPrToTable : IExtensionApplication
    {
        // ReSharper disable once RedundantDefaultMemberInitializer
        private bool _askRow = false;
        private int _round = 2;
        // Загрузка в автокад
        public void Initialize()
        {
            // Добавляем контекстное меню
            ObjectContextMenu.Attach();
        }

        public void Terminate()
        {

        }

        [CommandMethod("ModPlus", "mpPrToTable", CommandFlags.UsePickSet)]
        public void MpPrToTableFunction()
        {
            try
            {
                bool.TryParse(mpSettings.MpSettings.GetValue("Settings", "mpPrToTable", "AskRow"), out _askRow);
                // Т.к. при нулевом значении строки возвращает ноль, то делаем через if
                if (int.TryParse(mpSettings.MpSettings.GetValue("Settings", "mpPrToTable", "Round"), out int integer))
                    _round = integer;
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                var db = doc.Database;
                //var filList = new[] { new TypedValue((int)DxfCode.Start, "INSERT") };
                //var filter = new SelectionFilter(filList);
                var opts = new PromptSelectionOptions();
                opts.Keywords.Add("СТрока");
                opts.Keywords.Add("ЗНаки");
                var kws = opts.Keywords.GetDisplayString(true);
                opts.MessageForAdding = "\nВыберите объекты, относящиеся к изделиям ModPlus: " + kws;
                opts.KeywordInput += delegate (object sender, SelectionTextInputEventArgs e)
                {
                    if (e.Input.Equals("СТрока"))
                    {
                        var pko = new PromptKeywordOptions("\nСпрашивать строку для начала заполнения [Да/Нет]: ",
                            "Да Нет");
                        pko.Keywords.Default = _askRow ? "Да" : "Нет";
                        var pkor = ed.GetKeywords(pko);
                        if (pkor.Status != PromptStatus.OK) return;
                        _askRow = pkor.StringResult.Equals("Да");
                        mpSettings.MpSettings.SetValue("Settings", "mpPrToTable", "AskRow", _askRow.ToString(), true);
                    }
                    else if (e.Input.Equals("ЗНаки"))
                    {
                        var pio = new PromptIntegerOptions("\nКоличество знаков после запятой: ")
                        {
                            AllowNegative = false,
                            AllowNone = false,
                            AllowZero = true,
                            DefaultValue = _round
                        };
                        var pir = ed.GetInteger(pio);
                        if (pir.Status != PromptStatus.OK) return;
                        _round = pir.Value;
                        mpSettings.MpSettings.SetValue("Settings", "mpPrToTable", "Round", _round.ToString(), true);
                    }
                };
                //var res = ed.GetSelection(opts, filter);
                var res = ed.GetSelection(opts);
                if (res.Status != PromptStatus.OK)
                    return;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var selSet = res.Value;
                    var objectIds = selSet.GetObjectIds();
                    if (objectIds.Length == 0) return;

                    var findProductsWin = new FindProductsProgress(doc, objectIds, tr);
                    if (findProductsWin.ShowDialog() == true)
                    {
                        var peo = new PromptEntityOptions("\nВыберите таблицу: ");
                        peo.SetRejectMessage("\nНеверный выбор! Это не таблица!");
                        peo.AddAllowedClass(typeof(Table), false);
                        var per = ed.GetEntity(peo);
                        if (per.Status != PromptStatus.OK) return;
                        // fill
                        FillTable(tr, per.ObjectId, findProductsWin.SpecificationItems, _askRow, ed, _round);
                    }

                    
                    tr.Commit();
                }
            }
            catch (System.Exception exception)
            {
                MpExWin.Show(exception);
            }
        }
        
        /// <summary>
        /// Проверка, что блок имеет атрибуты для заполнения спецификации
        /// </summary>
        /// <param name="tr"></param>
        /// <param name="objectId"></param>
        /// <returns></returns>
        public static bool HasAttributesForSpecification(Transaction tr, ObjectId objectId)
        {
            var allowAttributesTags = new List<string> { "mp:позиция", "mp:обозначение", "mp:наименование", "mp:масса", "mp:примечание" };
            var blk = tr.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            if (blk != null)
            {
                // Если это блок
                if (blk.AttributeCollection.Count > 0)
                {
                    foreach (ObjectId id in blk.AttributeCollection)
                    {
                        var attr = tr.GetObject(id, OpenMode.ForRead) as AttributeReference;
                        if (allowAttributesTags.Contains(attr?.Tag.ToLower())) return true;
                    }
                }
            }
            return false;
        }
        /// <summary>
        /// Получение "Продукта" из атрибутов блока
        /// </summary>
        /// <returns></returns>
        public static SpecificationItem GetProductFromBlockByAttributes(Transaction tr, ObjectId objectId)
        {
            var blk = tr.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            if (blk != null)
            {
                // Если это блок
                if (blk.AttributeCollection.Count > 0)
                {
                    var mpPosition = string.Empty;
                    var mpDesignation = string.Empty;
                    var mpName = string.Empty;
                    var mpMass = string.Empty;
                    var mpNote = string.Empty;
                    foreach (ObjectId id in blk.AttributeCollection)
                    {
                        var attr = tr.GetObject(id, OpenMode.ForRead) as AttributeReference;
                        if (attr != null)
                        {
                            if (attr.Tag.ToLower().Equals("mp:позиция")) mpPosition = attr.TextString;
                            if (attr.Tag.ToLower().Equals("mp:обозначение")) mpDesignation = attr.TextString;
                            if (attr.Tag.ToLower().Equals("mp:наименование")) mpName = attr.TextString;
                            if (attr.Tag.ToLower().Equals("mp:масса")) mpMass = attr.TextString;
                            if (attr.Tag.ToLower().Equals("mp:примечание")) mpNote = attr.TextString;
                        }
                    }
                    double? mass = double.TryParse(mpMass, out double d) ? d : 0;

                    SpecificationItem specificationItem = new SpecificationItem(
                        null, String.Empty, String.Empty, String.Empty, String.Empty, SpecificationItemInputType.HandInput, String.Empty, String.Empty, String.Empty,
                        mass);
                    //specificationItem.BeforeName = mpName;
                    GetSpecificationItemNameFromAttr(specificationItem, mpName);
                    specificationItem.Designation = mpDesignation;
                    specificationItem.Position = mpPosition;
                    specificationItem.Note = mpNote;

                    return specificationItem;
                }
            }
            return null;
        }
        /// <summary>
        /// Получение наименования из атрибута с установкой значения "Есть сталь"
        /// </summary>
        public static void GetSpecificationItemNameFromAttr(SpecificationItem specificationItem, string attrValue)
        {
            var hasSteel = false;
            if (attrValue.Contains("$") & attrValue.Contains("?"))
            {
                var splitStr = attrValue.Split('$');
                if (splitStr.Count() == 4)
                {
                    try
                    {
                        specificationItem.HasSteel = true;
                        specificationItem.SteelVisibility = Visibility.Visible;
                        specificationItem.BeforeName = splitStr[0];
                        specificationItem.TopName = splitStr[1];
                        specificationItem.AfterName = splitStr[3];
                        specificationItem.SteelDoc = splitStr[2].Split('?')[0];
                        specificationItem.SteelType = splitStr[2].Split('?')[1];
                        hasSteel = true;
                    }
                    catch
                    {
                        hasSteel = false;
                    }
                }
            }
            if(!hasSteel)
            {
                specificationItem.HasSteel = false;
                specificationItem.SteelVisibility = Visibility.Collapsed;
                specificationItem.SteelType = string.Empty;
                specificationItem.SteelDoc = string.Empty;
                specificationItem.BeforeName = attrValue;
                specificationItem.AfterName = string.Empty;
                specificationItem.TopName = string.Empty;
            }
        }

        private static void FillTable(Transaction tr, ObjectId tblId, List<SpecificationItem> sItems, bool askRow, Editor ed, int round)
        {
            // Получаем список элементов для спецификации
            //List<SpecificationItem> sItems = data.Select((t, i) => t.GetSpecificationItem(count[i])).ToList();
            if (sItems.Count == 0) return;
            var table = (Table)tr.GetObject(tblId, OpenMode.ForWrite);
            var startRow = 2;
            if (!askRow)
            {
                if (table.TableStyleName.Equals("Mp_GOST_P_21.1101_F8"))
                {
                    startRow = 3;
                }
                int firstEmptyRow;
                CheckAndAddRowCount(table, startRow, sItems.Count, out firstEmptyRow);
                FillTableRows(table, firstEmptyRow, sItems, round);
            }
            else
            {
                var ppo = new PromptPointOptions("\nВыберите строку: ");
                var end = false;
                var vector = new Vector3d(0.0, 0.0, 1.0);
                while (end == false)
                {
                    var ppr = ed.GetPoint(ppo);
                    if (ppr.Status != PromptStatus.OK) return;
                    try
                    {
                        var tblhittestinfo = table.HitTest(ppr.Value, vector);
                        if (tblhittestinfo.Type == TableHitTestType.Cell)
                        {
                            startRow = tblhittestinfo.Row;
                            end = true;
                        }
                    } // try
                    catch
                    {
                        MpMsgWin.Show("Не попали в ячейку!");
                    }
                } // while
                int firstEmptyRow;
                CheckAndAddRowCount(table, startRow, sItems.Count, out firstEmptyRow);
                FillTableRows(table, startRow, sItems, round);
            }
        }

        /// <summary>
        /// Заполнение ячеек таблицы
        /// </summary>
        /// <param name="table">Таблица</param>
        /// <param name="firstRow">Номер строки для начала заполнения</param>
        /// <param name="sItems">Заполняемые данные</param>
        /// <param name="round">Количество знаков после запятой</param>
        private static void FillTableRows(Table table, int firstRow, IList<SpecificationItem> sItems, int round)
        {
            // Делаем итерацию по кол-ву элементов
            for (var i = 0; i < sItems.Count; i++)
            {
                
                var mass = string.Empty;
                if(sItems[i].Mass != null)
                    mass = Math.Round(sItems[i].Mass.Value, round).ToString(CultureInfo.InvariantCulture);
                // В зависимости от Наименования и стали создаем строку наименования
                string name;
                if (sItems[i].HasSteel)
                {
                    name = "\\A1;{\\C0;" + sItems[i].BeforeName + " \\H0.9x;\\S" + sItems[i].TopName + "/" +
                           sItems[i].SteelDoc + " " + sItems[i].SteelType + ";\\H1.1111x; " + sItems[i].AfterName;
                }
                else name = sItems[i].BeforeName + " " + sItems[i].TopName + " " + sItems[i].AfterName;

                // Если это таблица ModPlus
                if (table.TableStyleName.Contains("Mp_"))
                {
                    if (table.TableStyleName.Equals("Mp_GOST_P_21.1101_F7") |
                        table.TableStyleName.Equals("Mp_DSTU_B_A.2.4-4_F7") |
                        table.TableStyleName.Equals("Mp_STB_2255_Z1"))
                    {
                        if (CheckColumnsCount(table.Columns.Count, 6))
                        {
                            // Позиция
                            table.Cells[firstRow + i, 0].TextString = sItems[i].Position.Trim();
                            // Обозначение
                            table.Cells[firstRow + i, 1].TextString = sItems[i].Designation.Trim();
                            // Наименование
                            table.Cells[firstRow + i, 2].TextString = name.Trim();
                            // Количество
                            table.Cells[firstRow + i, 3].TextString = sItems[i].Count;
                            // Масса
                            table.Cells[firstRow + i, table.Columns.Count - 2].TextString = mass.Trim();
                        }
                    }
                    if (table.TableStyleName.Equals("Mp_GOST_P_21.1101_F8"))
                    {
                        // Позиция
                        table.Cells[firstRow + i, 0].TextString = sItems[i].Position.Trim();
                        // Обозначение
                        table.Cells[firstRow + i, 1].TextString = sItems[i].Designation.Trim();
                        // Наименование
                        table.Cells[firstRow + i, 2].TextString = name.Trim();
                        // Количество
                        table.Cells[firstRow + i, 3].TextString = sItems[i].Count;
                        // Масса
                        table.Cells[firstRow + i, table.Columns.Count - 2].TextString = mass.Trim();
                    }
                    if (table.TableStyleName.Equals("Mp_GOST_21.501_F7"))
                    {
                        if (CheckColumnsCount(table.Columns.Count, 4))
                        {
                            // Позиция
                            table.Cells[firstRow + i, 0].TextString = sItems[i].Position.Trim();
                            // Наименование
                            table.Cells[firstRow + i, 1].TextString = name.Trim();
                            // Количество
                            table.Cells[firstRow + i, 2].TextString = sItems[i].Count;
                            // Масса
                            table.Cells[firstRow + i, table.Columns.Count - 1].TextString = mass.Trim();
                        }
                    }
                    if (table.TableStyleName.Equals("Mp_GOST_21.501_F8"))
                    {
                        if (CheckColumnsCount(table.Columns.Count, 6))
                        {
                            // Позиция
                            table.Cells[firstRow + i, 1].TextString = sItems[i].Position.Trim();
                            // Наименование
                            table.Cells[firstRow + i, 2].TextString = name.Trim();
                            // Количество
                            table.Cells[firstRow + i, 3].TextString = sItems[i].Count;
                            // Масса
                            table.Cells[firstRow + i, table.Columns.Count - 2].TextString = mass.Trim();
                        }
                    }
                    if (table.TableStyleName.Equals("Mp_GOST_2.106_F1"))
                    {
                        if (CheckColumnsCount(table.Columns.Count, 7))
                        {
                            // Позиция
                            table.Cells[firstRow + i, 2].TextString = sItems[i].Position.Trim();
                            // Обозначение
                            table.Cells[firstRow + i, 3].TextString = sItems[i].Designation.Trim();
                            // Наименование
                            table.Cells[firstRow + i, 4].TextString = name.Trim();
                            // Количество
                            table.Cells[firstRow + i, 5].TextString = sItems[i].Count;
                        }
                    }
                    if (table.TableStyleName.Equals("Mp_GOST_2.106_F1a"))
                    {
                        if (CheckColumnsCount(table.Columns.Count, 5))
                        {
                            // Позиция
                            table.Cells[firstRow + i, 0].TextString = sItems[i].Position.Trim();
                            // Обозначение
                            table.Cells[firstRow + i, 1].TextString = sItems[i].Designation.Trim();
                            // Наименование
                            table.Cells[firstRow + i, 2].TextString = name.Trim();
                            // Количество
                            table.Cells[firstRow + i, 3].TextString = sItems[i].Count;
                        }
                    }
                }
                else
                // Если таблица не из плагина
                {
                    if (MpQstWin.Show("Таблица не является таблицей ModPlus. Данные могут заполнится не верно!" +
                                      Environment.NewLine + "Продолжить?"))
                    {
                        if (table.Columns.Count == 4)
                        {
                            // Позиция
                            table.Cells[firstRow + i, 0].TextString = sItems[i].Position.Trim();
                            // Наименование
                            table.Cells[firstRow + i, 1].TextString = name.Trim();
                            // Количество
                            table.Cells[firstRow + i, 2].TextString = sItems[i].Count;
                            // Масса
                            table.Cells[firstRow + i, table.Columns.Count - 1].TextString = mass.Trim();
                        }
                        if (table.Columns.Count == 5)
                        {
                            // Позиция
                            table.Cells[firstRow + i, 0].TextString = sItems[i].Position.Trim();
                            // Обозначение
                            table.Cells[firstRow + i, 1].TextString = sItems[i].Designation.Trim();
                            // Наименование
                            table.Cells[firstRow + i, 2].TextString = name.Trim();
                            // Количество
                            table.Cells[firstRow + i, 3].TextString = sItems[i].Count;
                        }
                        if (table.Columns.Count >= 6)
                        {
                            // Позиция
                            table.Cells[firstRow + i, 0].TextString = sItems[i].Position.Trim();
                            // Обозначение
                            table.Cells[firstRow + i, 1].TextString = sItems[i].Designation.Trim();
                            // Наименование
                            table.Cells[firstRow + i, 2].TextString = name.Trim();
                            // Количество
                            table.Cells[firstRow + i, 3].TextString = sItems[i].Count;
                            // Масса
                            table.Cells[firstRow + i, table.Columns.Count - 2].TextString = mass.Trim();
                        }
                    }
                }
            }
        }
        private static bool CheckColumnsCount(int columns, int need)
        {
            return columns == need || MpQstWin.Show("В таблице неверное количество столбцов!" + Environment.NewLine + "Продолжить?");
        }
        private static void CheckAndAddRowCount(Table table, int startRow, int sItemsCount, out int firstEmptyRow)
        {
            var rows = table.Rows.Count;
            var firstRow = startRow;
            firstEmptyRow = startRow; // Первая пустая строка
            // Пробегаем по всем ячейкам и проверяем "чистоту" таблицы
            var empty = true;
            var stopLoop = false;
            for (var i = startRow; i <= table.Rows.Count - 1; i++)
            {
                for (var j = 0; j < table.Columns.Count; j++)
                {
                    if (!table.Cells[i, j].TextString.Equals(string.Empty))
                    {
                        empty = false;
                        stopLoop = true;
                        break;
                    }
                }
                if (stopLoop) break;
            }
            // Если не пустая
            if (!empty)
            {
                if (!MpQstWin.Show("Таблица не пуста! Переписать?" + Environment.NewLine + "Да - переписать, Нет - дополнить"))
                {
                    // Если "Нет", тогда ищем последуюю пустую строку
                    // Если последняя строка не пуста, то добавляем 
                    // еще строчку, иначе...
                    var findEmpty = true;
                    for (var j = 0; j < table.Columns.Count; j++)
                    {
                        if (!string.IsNullOrEmpty(table.Cells[rows - 1, j].TextString))
                        {
                            //table.InsertRows(rows, 8, 1);
                            table.InsertRowsAndInherit(rows, rows - 1, 1);
                            rows++;
                            firstRow = rows - 1; // Так как таблица не обновляется
                            findEmpty = false; // чтобы не искать последнюю пустую
                            break;
                        }
                    }
                    if (findEmpty)
                    {
                        // идем по таблице в обратном порядке.
                        stopLoop = false;
                        for (var i = rows - 1; i >= 2; i--)
                        {
                            // Сделаем счетчик k
                            // Если ячейка пустая - будем увеличивать, а иначе - обнулять
                            var k = 1;
                            for (var j = 0; j < table.Columns.Count; j++)
                            {
                                if (table.Cells[i, j].TextString.Equals(string.Empty))
                                {
                                    firstRow = i;
                                    k++;
                                    // Если счетчик k равен количеству колонок
                                    // значит вся строка пустая и можно тормозить цикл
                                    if (k == table.Columns.Count)
                                    {
                                        stopLoop = true;
                                        break;
                                    }
                                }
                                else
                                {
                                    stopLoop = true;
                                    break;
                                }
                            }
                            if (stopLoop) break;
                        }
                        // Разбиваем ячейки
                        ////////////////////////////////////////
                    }
                }
                // Если "да", то очищаем таблицу
                else
                {
                    for (var i = startRow; i <= rows - 1; i++)
                    {
                        for (var j = 0; j < table.Columns.Count; j++)
                        {
                            table.Cells[i, j].TextString = string.Empty;
                            table.Cells[i, j].IsMergeAllEnabled = false;
                        }
                    }
                    // Разбиваем ячейки
                    //table.UnmergeCells(
                }
            }
            // Если в таблице мало строк
            if (sItemsCount > rows - firstRow)
                table.InsertRowsAndInherit(firstRow, firstRow, (sItemsCount - (rows - firstRow) + 1));
            // После всех манипуляций ищем первую пустую строчку
            for (var j = 0; j < rows; j++)
            {
                var isEmpty = table.Rows[j].IsEmpty;
                if (isEmpty != null && isEmpty.Value)
                {
                    firstEmptyRow = j;
                    break;
                }
            }
        }
    }

    public class ObjectContextMenu
    {
        public static ContextMenuExtension MpPrToTableCme;
        public static void Attach()
        {
            if (MpPrToTableCme == null)
            {
                MpPrToTableCme = new ContextMenuExtension();
                var miEnt = new MenuItem("MP:Добавить в спецификацию");
                miEnt.Click += SendCommand;
                MpPrToTableCme.MenuItems.Add(miEnt);
                //ADD the popup item
                MpPrToTableCme.Popup += contextMenu_Popup;

                var rxcEnt = RXObject.GetClass(typeof(Entity));
                Autodesk.AutoCAD.ApplicationServices.Application.AddObjectContextMenuExtension(rxcEnt, MpPrToTableCme);
            }
        }

        private static void SendCommand(object sender, EventArgs e)
        {
            AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("_.MPPRTOTABLE ", true, false, false);
        }

        public static void Detach()
        {
            if (MpPrToTableCme != null)
            {
                var rxcEnt = RXObject.GetClass(typeof(Entity));
                Autodesk.AutoCAD.ApplicationServices.Application.RemoveObjectContextMenuExtension(rxcEnt, MpPrToTableCme);
            }
        }
        // Обработка выпадающего меню
        static void contextMenu_Popup(object sender, EventArgs e)
        {
            try
            {
                var contextMenu = sender as ContextMenuExtension;

                if (contextMenu != null)
                {
                    var doc = AcApp.DocumentManager.MdiActiveDocument;
                    var ed = doc.Editor;
                    var db = doc.Database;
                    // This is the "Root context menu" item
                    var rootItem = contextMenu.MenuItems[0];
                    var acSsPrompt = ed.SelectImplied();
                    var mVisible = true;

                    if (acSsPrompt.Status == PromptStatus.OK)
                    {
                        var set = acSsPrompt.Value;
                        var ids = set.GetObjectIds();
                        // проходим по всем выбранным блокам и если хоть один не мой - отключаем меню
                        foreach (var objectId in ids)
                        {
                            using (var tr = doc.TransactionManager.StartTransaction())
                            {
                                var entity = tr.GetObject(objectId, OpenMode.ForRead) as Entity;
                                var mpProductToSave = XDataHelpersForProducts.NewFromEntity(entity) as MpProductToSave;
                                if (mpProductToSave == null)
                                {
                                    mVisible = false;
                                    break;
                                }
                            }
                        }
                    }
                    rootItem.Visible = mVisible;
                }
            }
            catch
            {
                // ignored
            }
        }
    }
}
