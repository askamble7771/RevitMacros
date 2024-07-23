/*
 * Created by SharpDevelop.
 * User: Ashish Kamble
 * Date: 16-07-2024
 * Time: 18:15
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace RevitViewSet
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("2057C9D5-89DF-436E-847B-4A51DB50CE9E")]
	public partial class ThisApplication
	{
		private void Module_Startup(object sender, EventArgs e)
		{

		}

		private void Module_Shutdown(object sender, EventArgs e)
		{

		}

		#region Revit Macros generated code
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Module_Startup);
			this.Shutdown += new System.EventHandler(Module_Shutdown);
		}
		#endregion
		
		public void AddSet()
	    {
	      Document doc = this.ActiveUIDocument.Document;
	      DeleteSets(doc);
	
	      List<Element> views = new FilteredElementCollector(doc).OfClass(typeof(View)).WhereElementIsNotElementType().Where(v => v.Name.Contains("Analytical ")).ToList();
	
	      ViewSet myViewSet = new ViewSet();
	      foreach (View v in views) {
	        myViewSet.Insert(v);
	      }
	
	
	      PrintManager printMgr = doc.PrintManager;
	      printMgr.PrintRange = Autodesk.Revit.DB.PrintRange.Select;
	      ViewSheetSetting viewSheetSetting = printMgr.ViewSheetSetting;
	      viewSheetSetting.CurrentViewSheetSet.Views = myViewSet;
	      IViewSheetSet viewSheetSet = viewSheetSetting.CurrentViewSheetSet;
	
	
	      using (Transaction transaction = new Transaction(doc, "NewSheetSet"))
	      {
	        transaction.Start();
	        try
	        {
	           if (viewSheetSet is ViewSheetSet)
	                  {  // if the CurrentViewSheetSet is one view sheet set of Print Setup, such as "set 1"
	                     viewSheetSet.Views = myViewSet;
	                     // make sure save the changes for the current view sheet set.
	                     viewSheetSetting.Save();
	                  }
	                  else if (viewSheetSet is InSessionViewSheetSet)
	                  {
	                     // if the CurrentViewSheetSet is in-session view sheet set of Print Setup
	                     viewSheetSet.Views = myViewSet;
	                     // For in-session view sheet set:
	                     // Cannot use Save() method, one InvalidOperationException thrown; Please use SaveAs(string newName) to save the the changes for the current view sheet set.
	                     // also can not invoke SaveAs, then the CurrentViewSheetSet changes to the new selected views, and keeps the in-session.
	          viewSheetSetting.SaveAs("NewSetNameA");
	//          printMgr.Apply();
	
	                  }
	        }
	        catch (Exception ex)
	        {
	          TaskDialog.Show("", ex.Message);
	        }
	        printMgr.Apply();
	        transaction.Commit();
      }

//      using (Transaction trans = new Transaction(doc)) {
//        trans.Start("Including Sets");
//        
//        ViewSheetSet selected = null;
//         ViewSheetSetting newviewSheetSetting = printMgr.ViewSheetSetting;
//        IViewSheetSet vsheetset = newviewSheetSetting.CurrentViewSheetSet;
//
//          FilteredElementCollector viewCollector = new FilteredElementCollector(doc);
//          viewCollector.OfClass(typeof(ViewSheetSet));
//          
//          //Find the sheet set that you just created
//          foreach (ViewSheetSet set in viewCollector.ToElements())
//          {
//            if (String.Compare(set.Name, "NewSetNameA") == 0)
//            {
//              selected = set;
//              break;
//            }
//          }
//          
//          //Set the current view sheet set to the one that you just created
//          viewSheetSetting.CurrentViewSheetSet = selected;
////          printMgr.PrintSetup.CurrentPrintSetting = printMgr.PrintSetup.InSession;
////          printMgr.PrintSetup.Save();
//          
//          //Set the views to which ever set you would like to print
////          viewSheetSetting.CurrentViewSheetSet.Views = myViewSet;
//          vsheetset.Views = myViewSet;
//          viewSheetSetting.Save();
////          printMgr.PrintSetup.Save();
////          printMgr.Apply();
//        trans.Commit();
//        
//      }
//      
      //IncludeSet(doc);
    }

 		
	
		public void DeleteSets(Document doc)
		{
//			Document doc = this.ActiveUIDocument.Document;

            List<ViewSheetSet> viewSet = new FilteredElementCollector(doc).OfClass(typeof(ViewSheetSet)).Cast<ViewSheetSet>().ToList();
            PrintManager printMgr = doc.PrintManager;
			printMgr.PrintRange = Autodesk.Revit.DB.PrintRange.Select;
            ViewSheetSetting setting = printMgr.ViewSheetSetting;

            using (Transaction deleteViewSheetSet = new Transaction(doc,"Delete View Sheet Set"))
            {
                deleteViewSheetSet.Start();

                foreach (var set in viewSet)
                {
//                	if(set.Name.Contains("Set"))
//                    {
                        setting.CurrentViewSheetSet = set;
                        setting.Delete();
//                    }
                }

                deleteViewSheetSet.Commit();
            }
		}
		
		public void IncludeSet(Document doc)
		{
			
			  ViewSheetSet existingViewSet = new FilteredElementCollector(doc)
			    .OfClass(typeof(ViewSheetSet))
			    .Cast<ViewSheetSet>()
			    .FirstOrDefault(vs => vs.Name == "NewSetNameA");
			
			  var schemaId = new Guid("57c66e83-4651-496b-aebb-69d085752c1b");
			
			  var schema =
			    Schema.ListSchemas().FirstOrDefault(schemaVS => schemaVS.GUID == schemaId);
			    //?? throw new InvalidOperationException("Schema ExportViewSheetSetListSchema not found");
			
			  var field =
			    schema.GetField("ExportViewViewSheetSetIdList");
			    //?? throw new InvalidOperationException("Field ExportViewViewSheetSetIdList not found");
			
			  var entity = doc.ProjectInformation.GetEntity(schema);
			
			  var viewSheetSetIds = entity.Get<IList<int>>(field);
			  var viewSheetSets = viewSheetSetIds.Select(id => doc.GetElement(new ElementId(id))).Cast<ViewSheetSet>();
			  var views = viewSheetSets.SelectMany(viewSheetSet => viewSheetSet.Views.Cast<View>());
			
			  // Add the additional ViewSheetSet
			
			  viewSheetSetIds.Add(existingViewSet.Id.IntegerValue);
			  entity.Set(field, viewSheetSetIds);
			  doc.ProjectInformation.SetEntity(entity);
			
		}
	}
}