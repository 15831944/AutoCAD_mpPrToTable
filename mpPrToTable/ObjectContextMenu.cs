namespace mpPrToTable
{
    using System;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Runtime;
    using Autodesk.AutoCAD.Windows;
    using ModPlusAPI;

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
            }

            var rxcEnt = RXObject.GetClass(typeof(Entity));
            Application.AddObjectContextMenuExtension(rxcEnt, MpPrToTableCme);
        }

        private static void SendCommand(object sender, EventArgs e)
        {
            Autodesk.AutoCAD.ApplicationServices.Core.Application
                .DocumentManager.MdiActiveDocument.SendStringToExecute("_.MPPRTOTABLE ", true, false, false);
        }

        public static void Detach()
        {
            if (MpPrToTableCme != null)
            {
                var rxcEnt = RXObject.GetClass(typeof(Entity));
                Application.RemoveObjectContextMenuExtension(rxcEnt, MpPrToTableCme);
            }
        }
    }
}