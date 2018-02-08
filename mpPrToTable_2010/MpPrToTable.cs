#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using mpProductInt;
using ModPlus;
using ModPlus.Helpers;
using ModPlusAPI;
using ModPlusAPI.Windows;
using Visibility = System.Windows.Visibility;

namespace mpPrToTable
{
    public class MpPrToTable : IExtensionApplication
    {
        private const string LangItem = "mpPrToTable";
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
            Statistic.SendCommandStarting(new Interface());

            try
            {
                bool.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "mpPrToTable", "AskRow"), out _askRow);
                // Т.к. при нулевом значении строки возвращает ноль, то делаем через if
                if (int.TryParse(UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "mpPrToTable", "Round"), out int integer))
                    _round = integer;
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                var db = doc.Database;
                //var filList = new[] { new TypedValue((int)DxfCode.Start, "INSERT") };
                //var filter = new SelectionFilter(filList);
                var opts = new PromptSelectionOptions();
                opts.Keywords.Add(Language.GetItem(LangItem, "h2"));
                opts.Keywords.Add(Language.GetItem(LangItem, "h3"));
                var kws = opts.Keywords.GetDisplayString(true);
                opts.MessageForAdding = "\n" + Language.GetItem(LangItem, "h1") + ": " + kws;
                opts.KeywordInput += delegate (object sender, SelectionTextInputEventArgs e)
                {
                    if (e.Input.Equals(Language.GetItem(LangItem, "h2")))
                    {
                        var pko = new PromptKeywordOptions("\n" + Language.GetItem(LangItem, "h4") +
                            " [" + Language.GetItem(LangItem, "yes") + "/" + Language.GetItem(LangItem, "no") + "]: ",
                            Language.GetItem(LangItem, "yes") + " " + Language.GetItem(LangItem, "no"));
                        pko.Keywords.Default = _askRow ? Language.GetItem(LangItem, "yes") : Language.GetItem(LangItem, "no");
                        var pkor = ed.GetKeywords(pko);
                        if (pkor.Status != PromptStatus.OK) return;
                        _askRow = pkor.StringResult.Equals(Language.GetItem(LangItem, "yes"));
                        UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "mpPrToTable", "AskRow", _askRow.ToString(), true);
                    }
                    else if (e.Input.Equals(Language.GetItem(LangItem, "h3")))
                    {
                        var pio = new PromptIntegerOptions("\n" + Language.GetItem(LangItem, "h5") + ": ")
                        {
                            AllowNegative = false,
                            AllowNone = false,
                            AllowZero = true,
                            DefaultValue = _round
                        };
                        var pir = ed.GetInteger(pio);
                        if (pir.Status != PromptStatus.OK) return;
                        _round = pir.Value;
                        UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "mpPrToTable", "Round", _round.ToString(), true);
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

                    var findProductsWin = new FindProductsProgress(objectIds, tr);
                    if (findProductsWin.ShowDialog() == true)
                    {
                        var peo = new PromptEntityOptions("\n" + Language.GetItem(LangItem, "h6") + ": ");
                        peo.SetRejectMessage("\n" + Language.GetItem(LangItem, "h7"));
                        peo.AddAllowedClass(typeof(Table), false);
                        var per = ed.GetEntity(peo);
                        if (per.Status != PromptStatus.OK) return;
                        // fill
                        FillTable(findProductsWin.SpecificationItems, _askRow, _round);
                    }


                    tr.Commit();
                }
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
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
            var allowAttributesTags = new List<string> { "mp:position", "mp:designation", "mp:name", "mp:mass", "mp:note" };
            if (Language.RusWebLanguages.Contains(Language.CurrentLanguageName))
                allowAttributesTags = new List<string> { "mp:позиция", "mp:обозначение", "mp:наименование", "mp:масса", "mp:примечание" };

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
                            if (attr.Tag.ToLower().Equals("mp:позиция") ||
                                attr.Tag.ToLower().Equals("mp:position"))
                                mpPosition = attr.TextString;
                            if (attr.Tag.ToLower().Equals("mp:обозначение") ||
                                attr.Tag.ToLower().Equals("mp:designation"))
                                mpDesignation = attr.TextString;
                            if (attr.Tag.ToLower().Equals("mp:наименование") ||
                                attr.Tag.ToLower().Equals("mp:name"))
                                mpName = attr.TextString;
                            if (attr.Tag.ToLower().Equals("mp:масса") ||
                                attr.Tag.ToLower().Equals("mp:mass"))
                                mpMass = attr.TextString;
                            if (attr.Tag.ToLower().Equals("mp:примечание") ||
                                attr.Tag.ToLower().Equals("mp:note"))
                                mpNote = attr.TextString;

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
                if (splitStr.Length == 4)
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
            if (!hasSteel)
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

        private static void FillTable(ICollection<SpecificationItem> sItems, bool askRow, int round)
        {
            if (sItems.Count == 0) return;
            var specificationItems = new List<InsertToAutoCad.SpecificationItemForTable>();
            foreach (var selectedSpecItem in sItems)
            {
                var mass = string.Empty;
                if (selectedSpecItem.Mass != null)
                    mass = Math.Round(selectedSpecItem.Mass.Value, round).ToString(CultureInfo.InvariantCulture);
                // В зависимости от Наименования и стали создаем строку наименования
                string name;
                if (selectedSpecItem.HasSteel)
                {
                    name = "\\A1;{\\C0;" + selectedSpecItem.BeforeName + " \\H0.9x;\\S" + selectedSpecItem.TopName + "/" +
                           selectedSpecItem.SteelDoc + " " + selectedSpecItem.SteelType + ";\\H1.1111x; " + selectedSpecItem.AfterName;
                }
                else name = selectedSpecItem.BeforeName + " " + selectedSpecItem.TopName + " " + selectedSpecItem.AfterName;
                specificationItems.Add(new InsertToAutoCad.SpecificationItemForTable(
                    selectedSpecItem.Position,
                    selectedSpecItem.Designation,
                    name,
                    mass,
                    selectedSpecItem.Count,
                    selectedSpecItem.Note
                ));
            }
            InsertToAutoCad.AddSpecificationItemsToTable(specificationItems, askRow);
        }
    }

    public class ObjectContextMenu
    {
        private const string LangItem = "mpPrToTable";
        public static ContextMenuExtension MpPrToTableCme;
        public static void Attach()
        {
            if (MpPrToTableCme == null)
            {
                MpPrToTableCme = new ContextMenuExtension();
                var miEnt = new MenuItem(Language.GetItem(LangItem, "h8"));
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
                if (sender is ContextMenuExtension contextMenu)
                {
                    var doc = AcApp.DocumentManager.MdiActiveDocument;
                    var ed = doc.Editor;
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
                                if (!(XDataHelpersForProducts.NewFromEntity(entity) is MpProductToSave mpProductToSave))
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
