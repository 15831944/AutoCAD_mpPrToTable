#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using mpMsg;
using mpProductInt;
using mpSettings;
using MahApps.Metro.Controls;
using ModPlus;

namespace mpPrToTable
{
    public partial class FindProductsProgress
    {
        private delegate void UpdateProgressBarDelegate(DependencyProperty dp, object value);
        private delegate void UpdateProgressTextDelegate(DependencyProperty dp, object value);

        private readonly Document doc;
        private readonly ObjectId[] objectIds;
        private Transaction tr;

        public List<SpecificationItem> SpecificationItems;

        public FindProductsProgress(Document _doc, ObjectId[] _objectIds, Transaction _tr)
        {
            InitializeComponent();
            MpWindowHelpers.OnWindowStartUp(
                this,
                MpSettings.GetValue("Settings", "MainSet", "Theme"),
                MpSettings.GetValue("Settings", "MainSet", "AccentColor"),
                MpSettings.GetValue("Settings", "MainSet", "BordersType")
                );
            doc = _doc;
            objectIds = _objectIds;
            tr = _tr;
            ContentRendered += FindProductsProgress_ContentRendered;
        }

        private void FindProductsProgress_ContentRendered(object sender, EventArgs e)
        {
            FindProducts(doc, objectIds, tr);
        }
        public void FindProducts(Document doc, ObjectId[] objectIds, Transaction tr)
        {
            try
            {
                //Create a new instance of our ProgressBar Delegate that points
                //  to the ProgressBar's SetValue method.
                var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
                var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);

                SpecificationItems = new List<SpecificationItem>();

                // progress start settings
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = objectIds.Count();
                ProgressBar.Value = 0;

                //using (var tr = doc.TransactionManager.StartTransaction())
                //{
                    var products = new List<MpProduct>();
                    var productsByAttr = new List<SpecificationItem>();

                    var counts = new List<int>();
                    var countsByAttr = new List<int>();

                    //var objectIds = psr.Value.GetObjectIds();
                    for (var i = 0; i < objectIds.Length; i++)
                    {
                        Dispatcher.Invoke(updatePbDelegate, DispatcherPriority.Background, System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)i);
                        Dispatcher.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty, i + "/" + objectIds.Length);
                        // Проверяем - если это блок и он имеет атрибуты для спецификации
                        if (mpPrToTable.MpPrToTable.HasAttributesForSpecification(tr, objectIds[i]))
                        {
                            var specificationItemByBlockAttributes =
                                mpPrToTable.MpPrToTable.GetProductFromBlockByAttributes(tr, objectIds[i]);
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
                            var entity = (Entity) tr.GetObject(objectIds[i], OpenMode.ForRead);
                            var mpProductToSave = XDataHelpersForProducts.NewFromEntity(entity) as MpProductToSave;
                            if (mpProductToSave != null)
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

                //    tr.Commit();
                //}
                DialogResult = true;
                Close();
            }
            catch (Exception exception)
            {
                MpExWin.Show(exception);
                DialogResult = false;
                Close();
            }
        }
        
    }
}
