/*
 * Created by SharpDevelop.
 * User: Ashish Kamble
 * Date: 09-05-2024
 * Time: 11:00
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Analysis;
using System.IO;
using System.Text;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;


namespace BIM
{
	[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
	[Autodesk.Revit.DB.Macros.AddInId("B3BD364B-C884-4493-86C3-013A42178032")]
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
		
		public void ReadJSON()
		{
			Document doc = this.ActiveUIDocument.Document;

			using (Transaction trans = new Transaction(doc, "wallCreation"))
			{
				trans.Start("Creating wall in Revit");
				
				string fileName = "E:\\building.json";
				StreamReader streamReader = new StreamReader(fileName);
				string json = streamReader.ReadToEnd();
				streamReader.Close();
				
				JObject builderRoot = JsonConvert.DeserializeObject<JObject>(json);
				Dictionary<string, Wall> wallDictionary = new Dictionary<string, Wall>(); // Dictionary to store wall names and walls
				

				for (int i = 1; i <= builderRoot["layers"].Count(); i++)
				{
					string layerName = "layer-" + i;
					double layerAltitude = builderRoot["layers"][layerName]["altitude"].ToObject<double>();

					CreateWalls(doc, builderRoot, layerName, layerAltitude, wallDictionary);
					CreateHoles(doc, builderRoot, layerName, layerAltitude, wallDictionary);
					CreateFloorAndCeiling(doc, builderRoot, layerName, layerAltitude, i);
				}
				
				trans.Commit();
			}
		}

		private void CreateWalls(Document doc, JObject builderRoot, string layerName, double layerAltitude, Dictionary<string, Wall> wallDictionary)
		{
			var lines = builderRoot["layers"][layerName]["lines"].ToObject<JObject>();
			var wallType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault();

			foreach (var item in lines)
			{
				var data = item.Value;

				string wallName = data["id"].ToObject<string>();
				string vertexId1 = data["vertices"][0].ToObject<string>();
				string vertexId2 = data["vertices"][1].ToObject<string>();

				double vertex1_x = builderRoot["layers"][layerName]["vertices"][vertexId1]["x"].ToObject<double>();
				double vertex1_y = builderRoot["layers"][layerName]["vertices"][vertexId1]["y"].ToObject<double>();
				double vertex2_x = builderRoot["layers"][layerName]["vertices"][vertexId2]["x"].ToObject<double>();
				double vertex2_y = builderRoot["layers"][layerName]["vertices"][vertexId2]["y"].ToObject<double>();

				double wallThickness = data["properties"]["thickness"].ToObject<double>();

				XYZ vertex1 = new XYZ(vertex1_x, vertex1_y, 0.0);
				XYZ vertex2 = new XYZ(vertex2_x, vertex2_y, 0.0);

				Wall wall = Wall.Create(doc, Line.CreateBound(vertex1, vertex2), wallType.Id, Level.GetNearestLevelId(doc, layerAltitude), 10, 0, true, false);
				wallDictionary.Add(wallName, wall);
			}
		}

		private void CreateHoles(Document doc, JObject builderRoot, string layerName, double layerAltitude, Dictionary<string, Wall> wallDictionary)
		{
			var holes = builderRoot["layers"][layerName]["holes"].ToObject<JObject>();
			
			ElementId level_Id = Level.GetNearestLevelId(doc, layerAltitude);
			Level level = doc.GetElement(level_Id) as Level;

			foreach (var item in holes)
			{
				var data = item.Value;

				string holeName = data["id"].ToObject<string>();
				var layer = builderRoot["layers"][layerName].ToObject<JObject>();
				
				double offset = layer["holes"][holeName]["offset"].ToObject<double>();
				string lineName = layer["holes"][holeName]["line"].ToObject<string>();
				
				double lineLength = layer["lines"][lineName]["properties"]["length"].ToObject<double>();
				string vertex1a = layer["lines"][lineName]["vertices"][0].ToObject<string>();
				string vertex2a = layer["lines"][lineName]["vertices"][1].ToObject<string>();
				
				double x1 =  layer["vertices"][vertex1a]["x"].ToObject<double>();
				double y1 = layer["vertices"][vertex1a]["y"].ToObject<double>();

				double x2 = layer["vertices"][vertex2a]["x"].ToObject<double>();
				double y2 = layer["vertices"][vertex2a]["y"].ToObject<double>();

				double z = data["properties"]["altitude"].ToObject<double>();
				XYZ location = new XYZ(0, 0, 0);
				
				//if condition is for wall is horizontal and else if for wall is vertical
				if ((x1 < x2 && y1 == y2) || (x2 < x1 && y1 == y2))//(x1-x2 < 0.1) (y1 - y2 > 1)
				{
					location = new XYZ(x1 + offset * (lineLength), y1, z + layerAltitude);
				}
				else if ((y1 < y2 && x1 == x2) || (y2 < y1 && x1 == x2))
				{
					location = new XYZ(x1, y1 + offset * (lineLength), z + layerAltitude);
				}

				//Check whether the hole is Window or Door
				if (data["type"].ToString() == "window") {
					FilteredElementCollector collectorWindows = new FilteredElementCollector(doc);
					FamilySymbol symbolWindow = collectorWindows.OfCategory(BuiltInCategory.OST_Windows).OfClass(typeof(FamilySymbol))
						.Cast<FamilySymbol>().FirstOrDefault();
					
					Wall specificWall = wallDictionary[lineName];
					FamilyInstance windowInstance = doc.Create.NewFamilyInstance(location, symbolWindow,specificWall, level, StructuralType.NonStructural);
					
					//changing width of the window
					Parameter widthParam = windowInstance.LookupParameter("width"); // Replace "Width" with the actual parameter name
					
					// Check if the parameter is valid
					if (widthParam != null)
					{
						// Set the new width value
						widthParam.Set(4.0);
					}
				}
				else if (data["type"].ToString() == "double door") {
					// Filter for double doors
					FilteredElementCollector collectorDoor = new FilteredElementCollector(doc);
					FamilySymbol symbolDoor = collectorDoor.OfCategory(BuiltInCategory.OST_Doors).OfClass(typeof(FamilySymbol))
						.Cast<FamilySymbol>().FirstOrDefault(sym => sym.FamilyName.Contains("double "));

					Wall specificWall = wallDictionary[lineName];
					FamilyInstance doorInstance = doc.Create.NewFamilyInstance(location, symbolDoor, specificWall, level, StructuralType.NonStructural);

					// Set width of Door
					Parameter doorWidthParameter = doorInstance.LookupParameter("Width");
					if (doorWidthParameter != null)
					{
						TaskDialog.Show("Result","Width set successfully!!");
						doorWidthParameter.Set(10.0); // Set door width to 6 feet
					}
				}
			}
		}
		
		private void CreateFloorAndCeiling(Document doc, JObject builderRoot, string layerName, double layerAltitude, int i)
		{
			var floorType = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>().FirstOrDefault();
			var ceilingType = new FilteredElementCollector(doc).OfClass(typeof(CeilingType)).Cast<CeilingType>().FirstOrDefault();
			ElementId levelId = Level.GetNearestLevelId(doc, layerAltitude);
			
			var areaCount = builderRoot["layers"][layerName]["areas"].Count();
			var areas = builderRoot["layers"][layerName]["areas"].ToObject<JObject>();
			
			foreach (var area in areas){
				
				var profile = new CurveLoop();
				List<XYZ> vertexList = new List<XYZ>();
				
				var areaValue = area.Value;
				var vertices = areaValue["vertices"].ToObject<JArray>();
				var fluidPointX = areaValue["fluidPoint"]["x"].ToObject<double>();
				var fluidPointY = areaValue["fluidPoint"]["y"].ToObject<double>();
				
				UV fluidPoint = new UV(fluidPointX, fluidPointY);
				
				for (int j = 0; j < vertices.Count; j++) {
					string vertexId1 = vertices[j].ToString();
					
					double vertex1_x = builderRoot["layers"][layerName]["vertices"][vertexId1]["x"].ToObject<double>();
					double vertex1_y = builderRoot["layers"][layerName]["vertices"][vertexId1]["y"].ToObject<double>();
					
					XYZ vertex1 = new XYZ(vertex1_x, vertex1_y, 0.0);
					vertexList.Add(vertex1);
				}
				
				for (int k = 0; k < vertexList.Count; k++) {
					profile.Append(Line.CreateBound(vertexList[k], vertexList[(k + 1) % vertexList.Count]));
				}
				Floor.Create(doc, new List<CurveLoop> { profile }, floorType.Id, levelId);
				
				if(i == builderRoot["layers"].Count()){
					//To create ceiling
					var ceiling = Ceiling.Create(doc, new List<CurveLoop> { profile }, ceilingType.Id, levelId);
					Parameter param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
					param.Set(10.00);
				}
				addRoomTags(doc, vertexList, levelId, i, fluidPoint);
			}
		}
		
		private void addRoomTags(Document doc, List<XYZ> vertexList, ElementId levelId, int i, UV fluidPoint)
		{
			Element element = doc.GetElement(levelId);
			Level level = element as Level;
			
			ViewPlan view = GetViewForLevel(doc, level, i);

			Room room = doc.Create.NewRoom(level, fluidPoint);
			RoomTag roomTag = doc.Create.NewRoomTag(new LinkElementId(room.Id), fluidPoint, view.Id);
		}
		
		private ViewPlan GetViewForLevel(Document document, Level level, int i)
		{
			FilteredElementCollector collector = new FilteredElementCollector(document);
			ICollection<Element> views = collector.OfClass(typeof(ViewPlan)).ToElements();
			foreach (Element elem in views)
			{
				ViewPlan view = elem as ViewPlan;
				if (view != null && view.Name == "Level "+i && view.ViewType.ToString() == "FloorPlan")
				{
					return view;
				}
			}
			
			return null;
		}
	}
}