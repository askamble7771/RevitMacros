/*
 * Created by SharpDevelop.
 * User: Ashish Kamble
 * Date: 24-04-2024
 * Time: 11:12
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Autodesk.Revit.Exceptions;
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

namespace BIMConfigurator
{
	public class HelloWorld
	{
		public HelloWorld()
		{
			TaskDialog.Show("Result", "Hello World in HelloWorld class");
		}
	}
	
	[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
	[Autodesk.Revit.DB.Macros.AddInId("9E781097-37B3-420F-A06B-59F9CAB2A55B")]
	

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
			
			using(Transaction trans = new Transaction(doc, "wallCreation"))
			{
				trans.Start("Creating wall in Revit");
				string fileName = "E:\\multiFloorBuilding.json";
				StreamReader streamReader = new StreamReader(fileName);
				string json = streamReader.ReadToEnd();
				streamReader.Close();
				
				JObject builderRoot = JsonConvert.DeserializeObject<JObject>(json);
				List<XYZ> vertexList = new List<XYZ>();
				var layers = builderRoot["layers"].ToObject<JObject>();
				FloorType floorType = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>().FirstOrDefault();
				CeilingType ceilingType = new FilteredElementCollector(doc).OfClass(typeof(CeilingType)).Cast<CeilingType>().FirstOrDefault();
				
				ElementId materialId = Material.Create(doc, "My Material17a1b");
				Material material = doc.GetElement(materialId) as Material;
				
				//Create a new property set that can be used by this material
				StructuralAsset strucAsset = new StructuralAsset("My Property Set17a1b", StructuralAssetClass.Concrete);
				strucAsset.Behavior = StructuralBehavior.Isotropic;
				strucAsset.Density = 600.0;
				strucAsset.ConcreteBendingReinforcement = 2.0;
				strucAsset.Lightweight = true;
				
				double inch0 = 0.0;
					double inch1 = 0.1/12.0;
					double inch2 = 2.0/12.0;
					double inch625 = ((5.0/8.0)/12.0);

				for (int i = 1; i <= layers.Count; i++) {
					string layerName = "layer-"+i;
					var vertices = builderRoot["layers"][layerName]["vertices"].ToObject<JObject>();
					var lines = builderRoot["layers"][layerName]["lines"].ToObject<JObject>();
					var holes = builderRoot["layers"][layerName]["holes"].ToObject<JObject>();
					double layerAltitude = builderRoot["layers"][layerName]["altitude"].ToObject<double>();
					
					// Create a wall
					WallType wallType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault();
					ElementId level_Id = Level.GetNearestLevelId(doc, layerAltitude);
					Level level = doc.GetElement(level_Id) as Level;
					
					Dictionary<string, Wall> wallDictionary = new Dictionary<string, Wall>(); // Dictionary to store wall names and walls
					int wallCount = 1;
					
					//Creating wall and store information with their names
					foreach (var item in lines) {
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
						
						WallType newWallType = wallType.Duplicate("NewWall"+ wallCount+"Floor"+i) as WallType;
					
					ElementId oldWallLayerMaterialId = wallType.GetCompoundStructure().GetLayers()[0].MaterialId;
					
					//exterior
					CompoundStructureLayer extLayerFinish1a = new CompoundStructureLayer(inch1, MaterialFunctionAssignment.Finish1, oldWallLayerMaterialId);
//					CompoundStructureLayer extLayerMembrane1a = new CompoundStructureLayer(inch0, MaterialFunctionAssignment.Membrane, oldWallLayerMaterialId);
					CompoundStructureLayer extLayerSubstrate1a = new CompoundStructureLayer(inch625, MaterialFunctionAssignment.Substrate, oldWallLayerMaterialId);
					
					//interior
//					CompoundStructureLayer intLayerMembrane2a = new CompoundStructureLayer(inch0, MaterialFunctionAssignment.Membrane, oldWallLayerMaterialId);
					CompoundStructureLayer intLayerFinish2a = new CompoundStructureLayer(inch2, MaterialFunctionAssignment.Finish2, oldWallLayerMaterialId);
					
					//Wall compound structure
					CompoundStructure compoundStructure = newWallType.GetCompoundStructure();
					
					//add all individual layers to wall compound structure layer
					IList<CompoundStructureLayer> layers1 = compoundStructure.GetLayers();
					
					//Set compund structure layers
					layers1.Insert(0, extLayerFinish1a);
//					layers1.Insert(1, extLayerMembrane1a);
					layers1.Insert(1, extLayerSubstrate1a);
//					layers1.Insert(3, intLayerMembrane2a);
					layers1.Insert(3, intLayerFinish2a);
					
					compoundStructure.SetLayers(layers1);
						
						Wall wall= Wall.Create(doc, Line.CreateBound(vertex1, vertex2), newWallType.Id, level_Id, 10, 0, true, false);
						
//Code to create Termal Properties of the wall - changing
						ThermalProperties prop = wallType.ThermalProperties;
						prop.Roughness = 6;
						prop.Absorptance = 0.9;
						
//Set wall thickness - changing
						compoundStructure.SetLayerWidth(2, wallThickness*0.0833333);//inches to feet
						compoundStructure.StructuralMaterialIndex = 2;
						
						newWallType.SetCompoundStructure(compoundStructure);
						wallDictionary.Add(wallName, wall);
						
						wallCount++;
					}
					
					foreach (var item in holes) {
						var data = item.Value;
						
						string holeName = data["id"].ToObject<string>();
						double offset = builderRoot["layers"][layerName]["holes"][holeName]["offset"].ToObject<double>();
						
						string lineName = builderRoot["layers"][layerName]["holes"][holeName]["line"].ToObject<string>();
						double lineLength = builderRoot["layers"][layerName]["lines"][lineName]["properties"]["length"].ToObject<double>();

						string vertex1a = builderRoot["layers"][layerName]["lines"][lineName]["vertices"][0].ToObject<string>();
						string vertex2a = builderRoot["layers"][layerName]["lines"][lineName]["vertices"][1].ToObject<string>();
						
						double x1 = builderRoot["layers"][layerName]["vertices"][vertex1a]["x"].ToObject<double>();
						double y1 = builderRoot["layers"][layerName]["vertices"][vertex1a]["y"].ToObject<double>();

						double x2 = builderRoot["layers"][layerName]["vertices"][vertex2a]["x"].ToObject<double>();
						double y2 = builderRoot["layers"][layerName]["vertices"][vertex2a]["y"].ToObject<double>();
						
						double z = builderRoot["layers"][layerName]["holes"][holeName]["properties"]["altitude"].ToObject<double>();
						XYZ location = new XYZ(0,0,0);
						
						if ((x1<x2 && y1==y2) || (x2<x1 && y1==y2)) {
							location = new XYZ(x1+offset*(lineLength/20), y1, z+layerAltitude);//We divide length by 20 as in json factor due to AHC reader
						}
						else if ((y1<y2 && x1==x2 ) || (y2<y1 && x1==x2)) {
							location = new XYZ(x1, y1+offset*(lineLength/20), z+layerAltitude);
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
					
					ElementId floorTypeId = Floor.GetDefaultFloorType(doc, false);
					CurveLoop profile = new CurveLoop();
					
					//Collecting vertices to create Floor or Ceiling
					foreach (var item in vertices)
					{
						var data = item.Value;

						string id = item.Key;
						double x = data["x"].ToObject<double>() ;
						double y = data["y"].ToObject<double>() ;
						double z = 0.0;
						XYZ vertex = new XYZ(x, y, z);
						vertexList.Add(vertex);
					}
					
					profile.Append(Line.CreateBound(vertexList.First(), vertexList[1]));
					profile.Append(Line.CreateBound(vertexList[1], vertexList[2]));
					profile.Append(Line.CreateBound(vertexList[2], vertexList[3]));
					profile.Append(Line.CreateBound(vertexList[3], vertexList.First()));
					
//Starting to changing material
					FloorType newFloorType =  floorType.Duplicate("New Wall17a"+i) as FloorType;
					CeilingType newCeilingType = ceilingType.Duplicate("New Wall17a"+i) as CeilingType;
					
					ElementId oldfloorLayerMaterialId = floorType.GetCompoundStructure().GetLayers()[0].MaterialId;
					
					
					//exterior
					CompoundStructureLayer extLayerFinish1 = new CompoundStructureLayer(inch1, MaterialFunctionAssignment.Finish1, oldfloorLayerMaterialId);
					CompoundStructureLayer extLayerMembrane1 = new CompoundStructureLayer(inch0, MaterialFunctionAssignment.Membrane, oldfloorLayerMaterialId);
					CompoundStructureLayer extLayerSubstrate1 = new CompoundStructureLayer(inch625, MaterialFunctionAssignment.Substrate, oldfloorLayerMaterialId);
					
					//interior
					CompoundStructureLayer intLayerMembrane2 = new CompoundStructureLayer(inch0, MaterialFunctionAssignment.Membrane, oldfloorLayerMaterialId);
					CompoundStructureLayer intLayerFinish2 = new CompoundStructureLayer(inch2, MaterialFunctionAssignment.Finish2, oldfloorLayerMaterialId);
					
					//Wall compound structure
					CompoundStructure compoundStructure1 = newFloorType.GetCompoundStructure();
					
					//add all individual layers to wall compound structure layer
					IList<CompoundStructureLayer> layers2 = compoundStructure1.GetLayers();
					
					//Set compund structure layers
					layers2.Insert(0, extLayerFinish1);
					layers2.Insert(1, extLayerMembrane1);
					layers2.Insert(2, extLayerSubstrate1);
					layers2.Insert(3, intLayerMembrane2);
					layers2.Insert(4, intLayerFinish2);
					
					compoundStructure1.SetLayers(layers2);
					newFloorType.SetCompoundStructure(compoundStructure1);
					newCeilingType.SetCompoundStructure(compoundStructure1);
					
					//To create floors
					Floor floor1 = Floor.Create(doc, new List<CurveLoop> { profile }, newFloorType.Id, level_Id);
					
					
					if(i==layers.Count){
						//To create ceiling
						var ceiling = Ceiling.Create(doc, new List<CurveLoop> { profile }, newCeilingType.Id, level_Id);
						Parameter param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
						param.Set(10.00);
					}
				}
				
				trans.Commit();
			}
		}
		
		
		public void CreateMultiFloorBuilding()
		{
			
			Document doc = this.ActiveUIDocument.Document;
			
			using (Transaction transaction = new Transaction(doc, "WallCreation"))
			{
				transaction.Start("Create Wall");
				List<XYZ> pointList = new List<XYZ>();
				List<XYZ> pointList1 = new List<XYZ>();
				
				// Getting the points for vertices of the wall
				pointList.Add(new XYZ(0, 0, 0));
				pointList.Add(new XYZ(30, 0, 0));
				pointList.Add(new XYZ(30, 30, 0));
				pointList.Add(new XYZ(0, 30, 0));
				
				pointList.Add(new XYZ(15, 0, 0));
				pointList.Add(new XYZ(15, 30, 0));
				
				pointList.Add(new XYZ(0, 15, 0));
				pointList.Add(new XYZ(30, 15, 0));
				
				// Create a wall
				WallType wallType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault();
				ElementId level1_Id = Level.GetNearestLevelId(doc, 5.00);
				
				Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
				
				Wall wall1 = Wall.Create(doc, Line.CreateBound(pointList[0], pointList[1]), wallType.Id, level1_Id, 10, 0, true, false);
				Wall wall2 = Wall.Create(doc, Line.CreateBound(pointList[1], pointList[2]), wallType.Id, level1_Id, 10, 0, true, false);
				Wall wall3 = Wall.Create(doc, Line.CreateBound(pointList[2], pointList[3]), wallType.Id, level1_Id, 10, 0, true, false);
				Wall wall4 = Wall.Create(doc, Line.CreateBound(pointList[3], pointList[0]), wallType.Id, level1_Id, 10, 0, true, false);

				
				ElementId level2_Id = Level.GetNearestLevelId(doc, 14.00);
				
				Wall wall5 = Wall.Create(doc, Line.CreateBound(pointList[0], pointList[1]), wallType.Id, level2_Id, 10, 0, true, false);
				Wall wall6 = Wall.Create(doc, Line.CreateBound(pointList[1], pointList[2]), wallType.Id, level2_Id, 10, 0, true, false);
				Wall wall7 = Wall.Create(doc, Line.CreateBound(pointList[2], pointList[3]), wallType.Id, level2_Id, 10, 0, true, false);
				Wall wall8 = Wall.Create(doc, Line.CreateBound(pointList[3], pointList[0]), wallType.Id, level2_Id, 10, 0, true, false);
				
				Wall wall9 = Wall.Create(doc, Line.CreateBound(pointList[4], pointList[5]), wallType.Id, level1_Id, 10, 0, true, false);
				Wall wall10 = Wall.Create(doc, Line.CreateBound(pointList[4], pointList[5]), wallType.Id, level2_Id, 10, 0, true, false);
				
				Wall wall11 = Wall.Create(doc, Line.CreateBound(pointList[6], pointList[7]), wallType.Id, level1_Id, 10, 0, true, false);
				Wall wall12 = Wall.Create(doc, Line.CreateBound(pointList[6], pointList[7]), wallType.Id, level2_Id, 10, 0, true, false);
				
				
				//Creating second house
				pointList1.Add(new XYZ(40, 0, 0));
				pointList1.Add(new XYZ(70, 0, 0));
				pointList1.Add(new XYZ(70, 30, 0));
				pointList1.Add(new XYZ(40, 30, 0));
				
				pointList1.Add(new XYZ(55, 0, 0));
				pointList1.Add(new XYZ(55, 30, 0));
				
				pointList1.Add(new XYZ(40, 15, 0));
				pointList1.Add(new XYZ(70, 15, 0));
				
				Wall wall1a = Wall.Create(doc, Line.CreateBound(pointList1[0], pointList1[1]), wallType.Id, level1_Id, 10, 0, true, false);
				Wall wall2a = Wall.Create(doc, Line.CreateBound(pointList1[1], pointList1[2]), wallType.Id, level1_Id, 10, 0, true, false);
				Wall wall3a = Wall.Create(doc, Line.CreateBound(pointList1[2], pointList1[3]), wallType.Id, level1_Id, 10, 0, true, false);
				Wall wall4a = Wall.Create(doc, Line.CreateBound(pointList1[3], pointList1[0]), wallType.Id, level1_Id, 10, 0, true, false);

				Wall wall5a = Wall.Create(doc, Line.CreateBound(pointList1[0], pointList1[1]), wallType.Id, level2_Id, 10, 0, true, false);
				Wall wall6a = Wall.Create(doc, Line.CreateBound(pointList1[1], pointList1[2]), wallType.Id, level2_Id, 10, 0, true, false);
				Wall wall7a = Wall.Create(doc, Line.CreateBound(pointList1[2], pointList1[3]), wallType.Id, level2_Id, 10, 0, true, false);
				Wall wall8a = Wall.Create(doc, Line.CreateBound(pointList1[3], pointList1[0]), wallType.Id, level2_Id, 10, 0, true, false);
				
				Wall wall9a = Wall.Create(doc, Line.CreateBound(pointList1[4], pointList1[5]), wallType.Id, level1_Id, 10, 0, true, false);
				Wall wall10a = Wall.Create(doc, Line.CreateBound(pointList1[4], pointList1[5]), wallType.Id, level2_Id, 10, 0, true, false);
				
				Wall wall11a = Wall.Create(doc, Line.CreateBound(pointList1[6], pointList1[7]), wallType.Id, level1_Id, 10, 0, true, false);
				Wall wall12a = Wall.Create(doc, Line.CreateBound(pointList1[6], pointList1[7]), wallType.Id, level2_Id, 10, 0, true, false);
				
				
				// Get a floor type for floor creation
				ElementId floorTypeId = Floor.GetDefaultFloorType(doc, false);
				CurveLoop profile = new CurveLoop();
				CurveLoop profile1 = new CurveLoop();
				
				profile.Append(Line.CreateBound(pointList.First(), pointList[1]));
				profile.Append(Line.CreateBound(pointList[1], pointList[2]));
				profile.Append(Line.CreateBound(pointList[2], pointList[3]));
				profile.Append(Line.CreateBound(pointList[3], pointList.First()));
				
				profile1.Append(Line.CreateBound(pointList1.First(), pointList1[1]));
				profile1.Append(Line.CreateBound(pointList1[1], pointList1[2]));
				profile1.Append(Line.CreateBound(pointList1[2], pointList1[3]));
				profile1.Append(Line.CreateBound(pointList1[3], pointList1.First()));
				
				
				//To craete floors
				Floor.Create(doc, new List<CurveLoop> { profile }, floorTypeId, level1_Id);
				Floor.Create(doc, new List<CurveLoop> { profile }, floorTypeId, level2_Id);
				
				//To craete floors
				Floor.Create(doc, new List<CurveLoop> { profile1 }, floorTypeId, level1_Id);
				Floor.Create(doc, new List<CurveLoop> { profile1 }, floorTypeId, level2_Id);
				
				//To create ceiling
				var ceiling = Ceiling.Create(doc, new List<CurveLoop> { profile }, ElementId.InvalidElementId, level2_Id);
				Parameter param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
				param.Set(10.00);
				
				//To create ceiling
				ceiling = Ceiling.Create(doc, new List<CurveLoop> { profile1 }, ElementId.InvalidElementId, level2_Id);
				param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
				param.Set(10.00);
				
				//Creating doors and windows for house1
				FilteredElementCollector collectorDoor = new FilteredElementCollector(doc);
				FilteredElementCollector collectorWindows = new FilteredElementCollector(doc);
				
				FamilySymbol symbolDoor = collectorDoor.OfCategory(BuiltInCategory.OST_Doors).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().FirstOrDefault();
				FamilySymbol symbolWindow = collectorWindows.OfCategory(BuiltInCategory.OST_Windows).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().FirstOrDefault();
				
				XYZ locationDoor1 = new XYZ(5, 0, 0);
				XYZ locationWindow1 = new XYZ(10, 0, 4);
				FamilyInstance instanceDoor = doc.Create.NewFamilyInstance(locationDoor1, symbolDoor, wall1, level, StructuralType.NonStructural);
				FamilyInstance instanceWindow = doc.Create.NewFamilyInstance(locationWindow1, symbolWindow, wall1, level, StructuralType.NonStructural);
				
				XYZ locationDoor2 = new XYZ(25, 0, 0);
				XYZ locationWindow2 = new XYZ(20, 0, 4);
				instanceDoor = doc.Create.NewFamilyInstance(locationDoor2, symbolDoor, wall1, level, StructuralType.NonStructural);
				instanceWindow = doc.Create.NewFamilyInstance(locationWindow2, symbolWindow, wall1, level, StructuralType.NonStructural);
				
				locationWindow1 = new XYZ(10, 0, 13);
				locationWindow2 = new XYZ(20, 0, 13);
				instanceWindow = doc.Create.NewFamilyInstance(locationWindow1, symbolWindow, wall5, level, StructuralType.NonStructural);
				instanceWindow = doc.Create.NewFamilyInstance(locationWindow2, symbolWindow, wall5, level, StructuralType.NonStructural);
				
				instanceWindow = doc.Create.NewFamilyInstance(new XYZ(0, 10, 5), symbolWindow, wall2, level, StructuralType.NonStructural);
				instanceWindow = doc.Create.NewFamilyInstance(new XYZ(0, 20, 5), symbolWindow, wall2, level, StructuralType.NonStructural);
				instanceWindow = doc.Create.NewFamilyInstance(new XYZ(0, 10, 13), symbolWindow, wall6, level, StructuralType.NonStructural);
				instanceWindow = doc.Create.NewFamilyInstance(new XYZ(0, 20, 13), symbolWindow, wall6, level, StructuralType.NonStructural);
				
				//Creating doors and windows for house2
				locationDoor1 = new XYZ(45, 0, 0);
				locationWindow1 = new XYZ(50, 0, 4);
				instanceDoor = doc.Create.NewFamilyInstance(locationDoor1, symbolDoor, wall1a, level, StructuralType.NonStructural);
				instanceWindow = doc.Create.NewFamilyInstance(locationWindow1, symbolWindow, wall1a, level, StructuralType.NonStructural);
				
				locationDoor2 = new XYZ(65, 0, 0);
				locationWindow2 = new XYZ(60, 0, 4);
				instanceDoor = doc.Create.NewFamilyInstance(locationDoor2, symbolDoor, wall1a, level, StructuralType.NonStructural);
				instanceWindow = doc.Create.NewFamilyInstance(locationWindow2, symbolWindow, wall1a, level, StructuralType.NonStructural);
				
				locationWindow1 = new XYZ(50, 0, 13);
				locationWindow2 = new XYZ(60, 0, 13);
				instanceWindow = doc.Create.NewFamilyInstance(locationWindow1, symbolWindow, wall5a, level, StructuralType.NonStructural);
				instanceWindow = doc.Create.NewFamilyInstance(locationWindow2, symbolWindow, wall5a, level, StructuralType.NonStructural);
				
				instanceWindow = doc.Create.NewFamilyInstance(new XYZ(0, 10, 5), symbolWindow, wall2a, level, StructuralType.NonStructural);
				instanceWindow = doc.Create.NewFamilyInstance(new XYZ(0, 20, 5), symbolWindow, wall2a, level, StructuralType.NonStructural);
				instanceWindow = doc.Create.NewFamilyInstance(new XYZ(0, 10, 13), symbolWindow, wall6a, level, StructuralType.NonStructural);
				instanceWindow = doc.Create.NewFamilyInstance(new XYZ(0, 20, 13), symbolWindow, wall6a, level, StructuralType.NonStructural);
				
				transaction.Commit();
			}
		}
		
		public void ChangeWallMaterial()
		{
			Document doc = this.ActiveUIDocument.Document;
			
//			WallType newWallType = null;
			
			using (Transaction transaction = new Transaction(doc, "WallCreation"))
			{
				transaction.Start("Create Wall");
				
				// Create a wall
				WallType wallType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault();
				var materialList = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Materials).OfClass(typeof(FamilySymbol));
				
				List<string> materialType = new List<string>();
				var countMaterial = 0; 
				foreach (var element in materialList) {
					FamilySymbol famSym = element as FamilySymbol;
					if(element != null){
						Parameter mType = element.get_Parameter(BuiltInParameter.MATERIAL_NAME);
						string mTypeString = mType.ToString();
						if(countMaterial == 0){
							TaskDialog.Show("Result" , mTypeString);
						}
						
						materialType.Add(mTypeString);
					}
					
				}
				
				ElementId level_Id = Level.GetNearestLevelId(doc, 5.0);
				Level level = doc.GetElement(level_Id) as Level;
				
				XYZ startPoint = new XYZ(0.0, 0.0, 0.0);
				XYZ endPoint = new XYZ(25.0, 0.0, 0.0);
				
				//Code to create Termal Properties of the wall
				ThermalProperties prop = wallType.ThermalProperties;
				prop.Roughness = 5;
				prop.Absorptance = 0.62;
				
//				ICollection<ElementId> wallIDs = wallType.GetMaterialIds(true);
				
				//Create the material
				ElementId materialId = Material.Create(doc, "My Material15");
				Material material = doc.GetElement(materialId) as Material;
				
				
				string materialName = "Fabric";
//				ElementId materialId = new ElementId(BuiltInParameter.MATERIAL_AREA);

				FilteredElementCollector collector = new FilteredElementCollector(doc);
				ICollection<Element> materials = collector.OfClass(typeof(Material)).ToElements();


				foreach (Element element in materials)
				{
					material = element as Material;
					if (material != null && material.Name == materialName)
					{
						// Found the material with the specified name
						materialId = material.Id;
//						material = doc.GetElement(materialId) as Material;
						// Now 'materialId' holds the ID of the material with the specified name
						break; // Exit the loop once the material is found
					}
				}

				Material wallMaterial = doc.GetElement(materialId) as Material;
				
				//Create a new property set that can be used by this material
				StructuralAsset strucAsset = new StructuralAsset(materialName, StructuralAssetClass.Concrete);
				strucAsset.Behavior = StructuralBehavior.Isotropic;
				strucAsset.Density = 232.0;
				strucAsset.ConcreteBendingReinforcement = 2.0;
				strucAsset.Lightweight = true;

				
				//Starting to aterial change
				WallType newWallType =  wallType.Duplicate("New Wall17a") as WallType;
				
//				ElementId oldLayerMaterialId = wallType.GetCompoundStructure().GetLayers()[0].MaterialId;
				ElementId oldLayerMaterialId = wallType.GetCompoundStructure().GetLayers()[0].MaterialId;
				double inch0 = 0.0;
				double inch1 = 1.0/12.0;
				double inch2 = 2.0/12.0;
				double inch625 = ((5.0/8.0)/12.0);
				
				//exterior
				CompoundStructureLayer extLayerFinish1 = new CompoundStructureLayer(inch1, MaterialFunctionAssignment.Finish1, materialId);
				CompoundStructureLayer extLayerMembrane1 = new CompoundStructureLayer(inch0, MaterialFunctionAssignment.Membrane, materialId);
				CompoundStructureLayer extLayerSubstrate1 = new CompoundStructureLayer(inch625, MaterialFunctionAssignment.Substrate, materialId);
				
				//interior
				CompoundStructureLayer intLayerMembrane2 = new CompoundStructureLayer(inch0, MaterialFunctionAssignment.Membrane, materialId);
				CompoundStructureLayer intLayerFinish2 = new CompoundStructureLayer(inch2, MaterialFunctionAssignment.Finish2, materialId);
				
				//Wall compound structure
				CompoundStructure compoundStructure = newWallType.GetCompoundStructure();
				
				//add all individual layers to wall compound structure layer
				IList<CompoundStructureLayer> layers = compoundStructure.GetLayers();
				
				//Set compund structure layers
				layers.Insert(0, extLayerFinish1);
				layers.Insert(1, extLayerMembrane1);
				layers.Insert(2, extLayerSubstrate1);
				layers.Insert(3, intLayerMembrane2);
				layers.Insert(5, intLayerFinish2);
				
				compoundStructure.SetLayers(layers);
				
				
				
				newWallType.SetCompoundStructure(compoundStructure);
				Wall.Create(doc, Line.CreateBound(startPoint, endPoint), newWallType.Id, level_Id, 10, 0,false, true);
				
				transaction.Commit();
			}
			
		}
		
		public void ChangeWallMaterialNew()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			using (Transaction transaction = new Transaction(doc, "WallCreation"))
			{
				transaction.Start("Create Wall");
				
				// Create a wall
				WallType wallType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault();

				ElementId level_Id = Level.GetNearestLevelId(doc, 5.0);
				Level level = doc.GetElement(level_Id) as Level;
				
				XYZ startPoint = new XYZ(0.0, 0.0, 0.0);
				XYZ endPoint = new XYZ(25.0, 0.0, 0.0);
				
				//Code to create Termal Properties of the wall
				ThermalProperties prop = wallType.ThermalProperties;
				prop.Roughness = 5;
				prop.Absorptance = 0.62;
				
				//Create the material
				ElementId materialId = Material.Create(doc, "My Material15");
				Material material = doc.GetElement(materialId) as Material;
				
				
				string materialName = "Fabric";
//				ElementId materialId = new ElementId(BuiltInParameter.MATERIAL_AREA);

				FilteredElementCollector collector = new FilteredElementCollector(doc);
				ICollection<Element> materials = collector.OfClass(typeof(Material)).ToElements();


				foreach (Element element in materials)
				{
					material = element as Material;
					if (material != null && material.Name == materialName)
					{
						// Found the material with the specified name
						materialId = material.Id;
//						material = doc.GetElement(materialId) as Material;
						// Now 'materialId' holds the ID of the material with the specified name
						break; // Exit the loop once the material is found
					}
				}

				Material wallMaterial = doc.GetElement(materialId) as Material;
				
				//Create a new property set that can be used by this material
				StructuralAsset strucAsset = new StructuralAsset(materialName, StructuralAssetClass.Concrete);
				strucAsset.Behavior = StructuralBehavior.Isotropic;
				strucAsset.Density = 232.0;
				strucAsset.ConcreteBendingReinforcement = 2.0;
				strucAsset.Lightweight = true;

				
				//Starting to material change
				WallType newWallType =  wallType.Duplicate("New Wall17a") as WallType;
				ElementId oldLayerMaterialId = wallType.GetCompoundStructure().GetLayers()[0].MaterialId;
				
				
				double inch0 = 0.0;
				double inch1 = 1.0/12.0;
				double inch2 = 2.0/12.0;
				double inch625 = ((5.0/8.0)/12.0);
								
				//exterior
				CompoundStructureLayer extLayerFinish1 = new CompoundStructureLayer(inch1, MaterialFunctionAssignment.Finish1, materialId);
				CompoundStructureLayer extLayerMembrane1 = new CompoundStructureLayer(inch0, MaterialFunctionAssignment.Membrane, materialId);
				CompoundStructureLayer extLayerSubstrate1 = new CompoundStructureLayer(inch625, MaterialFunctionAssignment.Substrate, materialId);
				CompoundStructureLayer extLayerStructure1 = new CompoundStructureLayer(10.0, MaterialFunctionAssignment.Structure, materialId);
				
				//interior
				CompoundStructureLayer intLayerMembrane2 = new CompoundStructureLayer(inch0, MaterialFunctionAssignment.Membrane, materialId);
				CompoundStructureLayer intLayerFinish2 = new CompoundStructureLayer(inch2, MaterialFunctionAssignment.Finish2, materialId);
				
				
				//Wall compound structure
				CompoundStructure compoundStructure = newWallType.GetCompoundStructure();
				
				//add all individual layers to wall compound structure layer
				IList<CompoundStructureLayer> layers = compoundStructure.GetLayers();
				
				
				//Set compund structure layers
				layers.Insert(0, extLayerFinish1);
				layers.Insert(1, extLayerMembrane1);
				layers.Insert(2, extLayerSubstrate1);
				layers.Insert(3, intLayerMembrane2);
				layers.Insert(4, extLayerStructure1);
				layers.Insert(5, intLayerFinish2);
				
				compoundStructure.SetLayers(layers);
				newWallType.SetCompoundStructure(compoundStructure);
				
				compoundStructure.DeleteLayer(4);
				Wall.Create(doc, Line.CreateBound(startPoint, endPoint), newWallType.Id, level_Id, 10, 0,false, false);
				compoundStructure.SetMaterialId(0, materialId);
				compoundStructure.StructuralMaterialIndex = 2;
				compoundStructure.DeleteLayer(3);
				
				
				transaction.Commit();
			}
			
		}
		
		public void ChangeWindowDoorMaterials()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			using (Transaction transaction = new Transaction(doc, "WallCreation"))
			{
				transaction.Start("Create Wall");
				
				WallType wallType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault();
				ElementId level_Id = Level.GetNearestLevelId(doc, 5.0);
				Level level = doc.GetElement(level_Id) as Level;
				
				XYZ startPoint = new XYZ(0.0, 0.0, 0.0);
				XYZ endPoint = new XYZ(25.0, 0.0, 0.0);
				
				Wall wall = Wall.Create(doc, Line.CreateBound(startPoint, endPoint), wallType.Id, level_Id, 10, 0,false, true);
				
				FilteredElementCollector collectorWindows = new FilteredElementCollector(doc);
				FamilySymbol symbolWindow = collectorWindows.OfCategory(BuiltInCategory.OST_Windows).OfClass(typeof(FamilySymbol))
					.Cast<FamilySymbol>().FirstOrDefault();
				
				XYZ location = new XYZ(5,0,2);
				
				FamilyInstance windowInstance = doc.Create.NewFamilyInstance(location, symbolWindow,wall, level, StructuralType.NonStructural);

				
				transaction.Commit();
			}
		}
		
		public void ChangeFloorCeilingMaterials()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			using (Transaction transaction = new Transaction(doc, "MatrialChanging"))
			{
				transaction.Start("Change Floor Ceiling Materials");
				
				TaskDialog.Show("Result", printHello());
				
				HelloWorld hello = new HelloWorld();
				
//				Class1 class1 = new Class1();
				
				
				transaction.Commit();
			}
		}
		
		public void getAnalyticConstructionValues()
		{
			Document doc = this.ActiveUIDocument.Document;
			using (Transaction transaction = new Transaction(doc, "GetAnalyticConstructionValues"))
			{
				transaction.Start("Started");
				
				WallType wallType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault();
				ElementId level_Id = Level.GetNearestLevelId(doc, 5.0);
				Level level1 = doc.GetElement(level_Id) as Level;
				
				XYZ point1 = new XYZ(0.0, 0.0, 0.0);
				XYZ point2 = new XYZ(25.0, 0.0, 0.0);
				
				Wall wall = Wall.Create(doc, Line.CreateBound(point1, point2), wallType.Id, level1.Id, 10.0, 0.0, false, true);
				
				XYZ location = new XYZ(10.0, 0.0, 2.0);
				
				FamilySymbol windowSymbol = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().FirstOrDefault();
				List<string> analyticConstructionList = new List<string>();
				
				FamilyThermalProperties fam = new FamilyThermalProperties();
				string constructionName = fam.AnalyticConstructionName;
				
				FamilyInstance window = doc.Create.NewFamilyInstance(location, windowSymbol, wall, StructuralType.NonStructural);
					
				transaction.Commit();
			}
		}
		
		public string printHello()
		{
			return "Hello World!!";
		}
		
		
		public void addLevel()
		{
			Document doc = this.ActiveUIDocument.Document;
			using(Transaction trans = new Transaction(doc))
			{
				trans.Start("Creating level");
				
				// The elevation to apply to the new level
			    double elevation = 30.0; 
			
			    // Begin to create a level
			    Level level = Level.Create(doc, elevation);
			    if (null == level)
			    {
			        throw new Exception("Create a new level failed.");
			    }
			
			    // Change the level name
			    level.Name = "New level";
				
				
				trans.Commit();
			}
			
		
		}
	}
}

//				string materialName = "Fabric";
//				ElementId materialId = new ElementId(BuiltInParameter.MATERIAL_AREA);
//
//				FilteredElementCollector collector = new FilteredElementCollector(doc);
//				ICollection<Element> materials = collector.OfClass(typeof(Material)).ToElements();
//
//
//				foreach (Element element in materials)
//				{
//					Material material = element as Material;
//					if (material != null && material.Name == materialName)
//					{
//						// Found the material with the specified name
//						materialId = material.Id;
//						// Now 'materialId' holds the ID of the material with the specified name
//						break; // Exit the loop once the material is found
//					}
//				}
//
//				Material wallMaterial = doc.GetElement(materialId) as Material;
//
//				string newName = "new" + wallMaterial.Name;
//				Material myMaterial = wallMaterial.Duplicate(newName);
//
//				//   					ElementId materialId = wallMaterial.Id;
////
////		            // Set the material parameter of the wall
////		            Parameter materialParameter = wall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_MATERIAL_PARAM);
//
//
//
//				//						Parameter parameter = wall.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM);
////						ForgeTypeId forge = new ForgeTypeId("");
//
//				Parameter parameter = wall.get_Parameter(BuiltInParameter.WALL_STRUCTURE_ID_PARAM);
//
//				if (parameter != null && wallMaterial != null)
//				{
//					parameter.Set(wallMaterial.Id);
//					TaskDialog.Show("Success", "Material successfully assigned to the wall.");
//				}
//
////						TaskDialog.Show("result" , materialName);




//				//Assign the property set to the material.
//				PropertySetElement pse = PropertySetElement.Create(doc, strucAsset);
//				material.SetMaterialAspectByPropertySet(MaterialAspect.Structural, pse.Id);

//				ElementId elemTypeId = wall.GetTypeId();
//				ElementType elemType = (ElementType)doc.GetElement(elemTypeId);
//	            elemType.LookupParameter("Material").Set(materialId);
