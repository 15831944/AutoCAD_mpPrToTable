namespace mpPrToTable
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Runtime;
    using ModPlus.Helpers;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using mpProductInt;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
    using Visibility = System.Windows.Visibility;

    public class ProductsToTable : IExtensionApplication
    {
        private const string LangItem = "mpPrToTable";

        // ReSharper disable once RedundantDefaultMemberInitializer
        private bool _askRow = false;
        private int _round = 2;

        /// <inheritdoc />
        public void Initialize()
        {
            // Добавляем контекстное меню
            AcApp.DocumentManager.DocumentCreated += Documents_DocumentCreated;
            AcApp.DocumentManager.DocumentActivated += Documents_DocumentActivated;

            foreach (Document document in AcApp.DocumentManager)
            {
                document.ImpliedSelectionChanged += Document_ImpliedSelectionChanged;
            }
        }

        /// <inheritdoc />
        public void Terminate()
        {
        }

        private static void Documents_DocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            if (e.Document != null)
            {
                e.Document.ImpliedSelectionChanged -= Document_ImpliedSelectionChanged;
                e.Document.ImpliedSelectionChanged += Document_ImpliedSelectionChanged;
            }
        }

        private static void Documents_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            if (e.Document != null)
            {
                e.Document.ImpliedSelectionChanged -= Document_ImpliedSelectionChanged;
                e.Document.ImpliedSelectionChanged += Document_ImpliedSelectionChanged;
            }
        }

        private static void Document_ImpliedSelectionChanged(object sender, EventArgs e)
        {
            var psr = AcApp.DocumentManager.MdiActiveDocument.Editor.SelectImplied();
            var detach = true;
            if (psr.Value != null)
            {
                using (AcApp.DocumentManager.MdiActiveDocument.LockDocument())
                {
                    using (var tr = new OpenCloseTransaction())
                    {
                        foreach (SelectedObject selectedObject in psr.Value)
                        {
                            if (selectedObject.ObjectId == ObjectId.Null)
                                continue;
                            var obj = tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead);
                            if (obj is Entity entity)
                            {
                                var xData = entity.GetXDataForApplication("ModPlusProduct");
                                if (xData != null)
                                {
                                    detach = false;
                                    break;
                                }
                            }
                        }

                        tr.Commit();
                    }
                }
            }

            if (detach)
                ObjectContextMenu.Detach();
            else
                ObjectContextMenu.Attach();
        }

        [CommandMethod("ModPlus", "mpPrToTable", CommandFlags.UsePickSet)]
        public void MpPrToTableFunction()
        {
            Statistic.SendCommandStarting(new ModPlusConnector());

            try
            {
                bool.TryParse(UserConfigFile.GetValue(LangItem, "AskRow"), out _askRow);

                // Т.к. при нулевом значении строки возвращает ноль, то делаем через if
                if (int.TryParse(UserConfigFile.GetValue(LangItem, "Round"), out var integer))
                    _round = integer;
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                var db = doc.Database;

                var opts = new PromptSelectionOptions();
                opts.Keywords.Add(Language.GetItem(LangItem, "h2"));
                opts.Keywords.Add(Language.GetItem(LangItem, "h3"));
                var kws = opts.Keywords.GetDisplayString(true);
                opts.MessageForAdding = "\n" + Language.GetItem(LangItem, "h1") + ": " + kws;
                opts.KeywordInput += (sender, e) =>
                {
                    if (e.Input.Equals(Language.GetItem(LangItem, "h2")))
                    {
                        var pko = new PromptKeywordOptions(
                            "\n" + Language.GetItem(LangItem, "h4") +
                            " [" + Language.GetItem(LangItem, "yes") + "/" + Language.GetItem(LangItem, "no") + "]: ",
                            Language.GetItem(LangItem, "yes") + " " + Language.GetItem(LangItem, "no"));
                        pko.Keywords.Default =
                            _askRow ? Language.GetItem(LangItem, "yes") : Language.GetItem(LangItem, "no");
                        var promptResult = ed.GetKeywords(pko);
                        if (promptResult.Status != PromptStatus.OK)
                            return;
                        _askRow = promptResult.StringResult.Equals(Language.GetItem(LangItem, "yes"));
                        UserConfigFile.SetValue(LangItem, "AskRow", _askRow.ToString(), true);
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
                        if (pir.Status != PromptStatus.OK)
                            return;
                        _round = pir.Value;
                        UserConfigFile.SetValue("mpPrToTable", "Round",
                            _round.ToString(), true);
                    }
                };

                var res = ed.GetSelection(opts);
                if (res.Status != PromptStatus.OK)
                    return;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var selSet = res.Value;
                    var objectIds = selSet.GetObjectIds();
                    if (objectIds.Length == 0)
                        return;

                    var findProductsWin = new FindProductsProgress(objectIds, tr);
                    if (findProductsWin.ShowDialog() == true)
                    {
                        var peo = new PromptEntityOptions("\n" + Language.GetItem(LangItem, "h6") + ": ");
                        peo.SetRejectMessage("\n" + Language.GetItem(LangItem, "h7"));
                        peo.AddAllowedClass(typeof(Table), false);
                        var per = ed.GetEntity(peo);
                        if (per.Status != PromptStatus.OK)
                            return;

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
        /// <param name="tr">Transaction</param>
        /// <param name="objectId">Block id</param>
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
                        if (allowAttributesTags.Contains(attr?.Tag.ToLower()))
                            return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Получение "Продукта" из атрибутов блока
        /// </summary>
        /// <param name="tr">Transaction</param>
        /// <param name="objectId">Block id</param>
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

                    double? mass = double.TryParse(mpMass, out var d) ? d : 0;

                    var specificationItem = new SpecificationItem(
                        null, 
                        string.Empty,
                        string.Empty, 
                        string.Empty,
                        string.Empty, 
                        SpecificationItemInputType.HandInput, 
                        string.Empty, 
                        string.Empty, 
                        string.Empty,
                        mass);

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
            if (sItems.Count == 0)
                return;
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
                else
                {
                    name = selectedSpecItem.BeforeName + " " + selectedSpecItem.TopName + " " + selectedSpecItem.AfterName;
                }

                specificationItems.Add(new InsertToAutoCad.SpecificationItemForTable(
                    selectedSpecItem.Position,
                    selectedSpecItem.Designation,
                    name,
                    mass,
                    selectedSpecItem.Count,
                    selectedSpecItem.Note));
            }

            InsertToAutoCad.AddSpecificationItemsToTable(specificationItems, askRow);
        }
    }
}
