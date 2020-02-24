using System;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using System.Collections.Generic;
using System.Linq;

namespace RevitAdd_in
{
    public class AddPanel : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            // Add a new ribbon panel
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("Тестовая панель");

            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData buttonData = new PushButtonData("cmdSignUp",
               "Подписать", thisAssemblyPath, "RevitAdd_in.SignUp");

            PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;

            // Optionally, other properties may be assigned to the button
            // a) tool-tip
            pushButton.ToolTip = "Скопировать фамилии подписантов в экземпляры штампа";

            // b) large bitmap
            //Uri uriImage = new Uri("Sign_Icon.png");
            //BitmapImage largeImage = new BitmapImage(uriImage);
            //pushButton.LargeImage = largeImage;

         return Result.Succeeded;
      }

      public Result OnShutdown(UIControlledApplication application)
      {
         // nothing to clean up in this simple case
         return Result.Succeeded;
      }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    //[RegenerationAttribute(RegenerationOption.Manual)]
    class SignUp : IExternalCommand
    {
        public Result Execute(ExternalCommandData revit,
          ref string message, ElementSet elements)
        {
            var sign = new Dictionary<string, string>();
            sign["ADSK_Штамп Строка 1 фамилия"] = "";
            sign["ADSK_Штамп Строка 2 фамилия"] = "";
            sign["ADSK_Штамп Строка 3 фамилия"] = "";
            sign["ADSK_Штамп Строка 4 фамилия"] = "";
            sign["ADSK_Штамп Строка 5 фамилия"] = "";
            sign["ADSK_Штамп Строка 6 фамилия"] = "";
            
            var nameOfForm = new[] { "Форма 3", "Форма 5", "Форма 6" };

            var dataSheet = new Dictionary<string, Dictionary<string, string>>();

            //Получение объектов приложения и документа
            UIApplication uiApp = revit.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            //создаем фильтр и получаем экземпляры титульников      	
            FilteredElementCollector coll = new FilteredElementCollector(doc);

            var vSheets = coll.OfClass(typeof(ViewSheet)).ToElements();
            string sheetID;

            foreach (var sheet in vSheets)
            {
                sheetID = sheet.Id.ToString();
                dataSheet[sheetID] = new Dictionary<string, string>();
                foreach (var parametr in sheet.GetOrderedParameters())
                    if (sign.ContainsKey(parametr.Definition.Name))
                        dataSheet[sheetID][parametr.Definition.Name] = parametr.AsString();
            }

            coll = new FilteredElementCollector(doc);
            var titles = coll.OfCategory(BuiltInCategory.OST_TitleBlocks).OfClass(typeof(FamilyInstance)).ToElements();
            var titleFiltrated = titles.Where(t => nameOfForm.Contains(t.Name));

            
            using (Transaction transaction = new Transaction(doc))
            {
                transaction.Start("Copy sign text");
                foreach (var title in titleFiltrated)
                {
                    sheetID = title.OwnerViewId.ToString();
                    foreach (var parametr in title.GetOrderedParameters())
                        if (sign.ContainsKey(parametr.Definition.Name))
                            parametr.Set(dataSheet[sheetID][parametr.Definition.Name]);
                }
                transaction.Commit();
                return Result.Succeeded;
            }

        }
    }

}
