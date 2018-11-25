namespace mpPrToTable
{
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Autodesk.AutoCAD.DatabaseServices;
    using mpProductInt;
    using ModPlus;
    using ModPlusAPI.Windows;

    public partial class FindProductsProgress
    {
        //private delegate void UpdateProgressBarDelegate(DependencyProperty dp, object value);
        //private delegate void UpdateProgressTextDelegate(DependencyProperty dp, object value);

        private readonly ObjectId[] _objectIds;
        private readonly Transaction _tr;

        public List<SpecificationItem> SpecificationItems;

        public FindProductsProgress(ObjectId[] objectIds, Transaction tr)
        {
            InitializeComponent();
            _objectIds = objectIds;
            _tr = tr;
            ProgressBar.Minimum = 0;
            ProgressBar.Maximum = objectIds.Length;
            ProgressBar.Value = 0;
            ContentRendered += FindProductsProgress_ContentRendered;
        }

        SynchronizationContext _context;
        private async void FindProductsProgress_ContentRendered(object sender, EventArgs e)
        {
            _context = SynchronizationContext.Current;

            Task task = new Task(FindProducts);
            task.Start();
            await task;
            DialogResult = true;
            Close();
        }

        private void FindProducts()
        {
            try
            {
                SpecificationItems = new List<SpecificationItem>();

                var products = new List<MpProduct>();
                var productsByAttr = new List<SpecificationItem>();

                var counts = new List<int>();
                var countsByAttr = new List<int>();

                for (var i = 0; i < _objectIds.Length; i++)
                {
                    // Проверяем - если это блок и он имеет атрибуты для спецификации
                    if (MpPrToTable.HasAttributesForSpecification(_tr, _objectIds[i]))
                    {
                        // post progress
                        _context.Post(_ =>
                        {
                            ProgressBar.Value = i;
                            ProgressText.Text = i + "/" + _objectIds.Length;
                        }, null);

                        var specificationItemByBlockAttributes =
                            MpPrToTable.GetProductFromBlockByAttributes(_tr, _objectIds[i]);
                        if (specificationItemByBlockAttributes != null)
                        {
                            if (
                                !productsByAttr.Contains(specificationItemByBlockAttributes,
                                    new SpecificationItemHelpers.EqualSpecificationItem()))
                            {
                                productsByAttr.Add(specificationItemByBlockAttributes);
                                countsByAttr.Add(1);
                            }
                            else
                            {
                                for (int j = 0; j < productsByAttr.Count; j++)
                                {
                                    if (productsByAttr[j].Position == specificationItemByBlockAttributes.Position &
                                        productsByAttr[j].BeforeName ==
                                        specificationItemByBlockAttributes.BeforeName &
                                        productsByAttr[j].Designation ==
                                        specificationItemByBlockAttributes.Designation &
                                        productsByAttr[j].Mass == specificationItemByBlockAttributes.Mass &
                                        productsByAttr[j].Note == specificationItemByBlockAttributes.Note)
                                        countsByAttr[j]++;
                                }
                            }
                        }
                    }
                    else // Иначе пробуем читать из расширенных данных
                    {
                        var entity = (Entity) _tr.GetObject(_objectIds[i], OpenMode.ForRead);
                        if (!entity.IsModPlusProduct()) continue;
                        // post progress
                        _context.Post(_ =>
                        {
                            ProgressBar.Value = i;
                            ProgressText.Text = i + "/" + _objectIds.Length;
                        }, null);

                        if (XDataHelpersForProducts.NewFromEntity(entity) is MpProductToSave mpProductToSave)
                        {
                            var productFromSaved = MpProduct.GetProductFromSaved(mpProductToSave);
                            if (productFromSaved != null)
                            {
                                if (!products.Contains(productFromSaved))
                                {
                                    products.Add(productFromSaved);
                                    counts.Add(1);
                                }
                                else
                                {
                                    var index = products.IndexOf(productFromSaved);
                                    counts[index]++;
                                }
                            }
                        }
                    }
                }

                if (!products.Any() & !productsByAttr.Any())
                {
                    DialogResult = false;
                    Close();
                }
                // Для продуктов собранных из атрибутов вставляем количество
                for (var i = 0; i < productsByAttr.Count; i++)
                {
                    productsByAttr[i].Count = countsByAttr[i].ToString();
                }
                // Добавляем продукты собранные из расширенных данных
                for (var j = 0; j < products.Count; j++)
                {
                    var specificationItem = products[j].GetSpecificationItem(counts[j]);
                    SpecificationItems.Add(specificationItem);
                }
                // Добавляем продукты, собранные из атрибутов
                foreach (var specificationItem in productsByAttr)
                {
                    SpecificationItems.Add(specificationItem);
                }
                // Сортировка по значению Позиции
                SpecificationItems.Sort(new SpecificationItemHelpers.AlphanumComparatorFastToSortByPosition());
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        //private void FindProductsProgress_ContentRendered(object sender, EventArgs e)
        //{
        //    FindProducts(_objectIds, _tr);
        //}
        //private void FindProducts(IList<ObjectId> objectIds, Transaction tr)
        //{
        //    try
        //    {
        //        //Create a new instance of our ProgressBar Delegate that points
        //        //  to the ProgressBar's SetValue method.
        //        var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
        //        var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);

        //        SpecificationItems = new List<SpecificationItem>();

        //        // progress start settings
        //        ProgressBar.Minimum = 0;
        //        ProgressBar.Maximum = objectIds.Count;
        //        ProgressBar.Value = 0;

        //        //using (var tr = doc.TransactionManager.StartTransaction())
        //        //{
        //        var products = new List<MpProduct>();
        //        var productsByAttr = new List<SpecificationItem>();

        //        var counts = new List<int>();
        //        var countsByAttr = new List<int>();

        //        //var objectIds = psr.Value.GetObjectIds();
        //        for (var i = 0; i < objectIds.Count; i++)
        //        {
        //            //Dispatcher.Invoke(updatePbDelegate, DispatcherPriority.Background, System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)i);
        //            //Dispatcher.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty, i + "/" + objectIds.Count);
        //            // Проверяем - если это блок и он имеет атрибуты для спецификации
        //            if (MpPrToTable.HasAttributesForSpecification(tr, objectIds[i]))
        //            {
        //                Dispatcher.Invoke(updatePbDelegate, DispatcherPriority.Background, System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)i);
        //                Dispatcher.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty, i + "/" + objectIds.Count);

        //                var specificationItemByBlockAttributes =
        //                    MpPrToTable.GetProductFromBlockByAttributes(tr, objectIds[i]);
        //                if (specificationItemByBlockAttributes != null)
        //                {
        //                    if (
        //                        !productsByAttr.Contains(specificationItemByBlockAttributes,
        //                            new SpecificationItemHelpers.EqualSpecificationItem()))
        //                    {
        //                        productsByAttr.Add(specificationItemByBlockAttributes);
        //                        countsByAttr.Add(1);
        //                    }
        //                    else
        //                    {
        //                        for (int j = 0; j < productsByAttr.Count; j++)
        //                        {
        //                            if (productsByAttr[j].Position == specificationItemByBlockAttributes.Position &
        //                                productsByAttr[j].BeforeName ==
        //                                specificationItemByBlockAttributes.BeforeName &
        //                                productsByAttr[j].Designation ==
        //                                specificationItemByBlockAttributes.Designation &
        //                                productsByAttr[j].Mass == specificationItemByBlockAttributes.Mass &
        //                                productsByAttr[j].Note == specificationItemByBlockAttributes.Note)
        //                                countsByAttr[j]++;
        //                        }
        //                    }
        //                }
        //            }
        //            else // Иначе пробуем читать из расширенных данных
        //            {
        //                var entity = (Entity)tr.GetObject(objectIds[i], OpenMode.ForRead);
        //                if (!entity.IsModPlusProduct()) continue;
        //                Dispatcher.Invoke(updatePbDelegate, DispatcherPriority.Background, System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)i);
        //                Dispatcher.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty, i + "/" + objectIds.Count);
        //                var mpProductToSave = XDataHelpersForProducts.NewFromEntity(entity) as MpProductToSave;
        //                if (mpProductToSave != null)
        //                {
        //                    var productFromSaved = MpProduct.GetProductFromSaved(mpProductToSave);
        //                    if (productFromSaved != null)
        //                    {
        //                        if (!products.Contains(productFromSaved))
        //                        {
        //                            products.Add(productFromSaved);
        //                            counts.Add(1);
        //                        }
        //                        else
        //                        {
        //                            var index = products.IndexOf(productFromSaved);
        //                            counts[index]++;
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        if (!products.Any() & !productsByAttr.Any())
        //        {
        //            DialogResult = false;
        //            Close();
        //        }
        //        // Для продуктов собранных из атрибутов вставляем количество
        //        for (var i = 0; i < productsByAttr.Count; i++)
        //        {
        //            productsByAttr[i].Count = countsByAttr[i].ToString();
        //        }
        //        // Добавляем продукты собранные из расширенных данных
        //        for (var j = 0; j < products.Count; j++)
        //        {
        //            var specificationItem = products[j].GetSpecificationItem(counts[j]);
        //            SpecificationItems.Add(specificationItem);
        //        }
        //        // Добавляем продукты, собранные из атрибутов
        //        foreach (var specificationItem in productsByAttr)
        //        {
        //            SpecificationItems.Add(specificationItem);
        //        }
        //        // Сортировка по значению Позиции
        //        SpecificationItems.Sort(new SpecificationItemHelpers.AlphanumComparatorFastToSortByPosition());

        //        //    tr.Commit();
        //        //}
        //        DialogResult = true;
        //        Close();
        //    }
        //    catch (Exception exception)
        //    {
        //        ExceptionBox.Show(exception);
        //        DialogResult = false;
        //        Close();
        //    }
        //}
    }
}
