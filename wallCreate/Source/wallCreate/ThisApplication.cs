/*
 * Created by SharpDevelop.
 * User: Ashish Kamble
 * Date: 09-02-2024
 * Time: 17:50
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

namespace wallCreate
{
	[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
	[Autodesk.Revit.DB.Macros.AddInId("3A791480-5B32-47E6-B568-60FF3737A5B9")]
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
		public void createWall()
		{
			var doc = ActiveUIDocument.Document;
			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("Create Wall");
				
				// Create a line for the wall's location
				XYZ startPoint = XYZ.Zero; // Assuming the start point is at the origin
				XYZ endPoint = new XYZ(10, 0, 0); // Assuming the end point is (10, 0, 0)
				Line line = Line.CreateBound(startPoint, endPoint);
				
				// Create a wall
				WallType wallType = new FilteredElementCollector(doc)
					.OfClass(typeof(WallType))
					.Cast<WallType>()
					.FirstOrDefault();
				
				Level level = new FilteredElementCollector(doc)
					.OfClass(typeof(Level))
					.Cast<Level>()
					.FirstOrDefault();
				
				if (wallType != null && level != null)
				{
					Wall wall = Wall.Create(doc, line, wallType.Id, level.Id, 10, 0, true, false);
					if (wall != null)
					{
						
						TaskDialog.Show("Success", "Wall created successfully.");
					}
				}
				
//				transaction.RollBack();
				transaction.Commit();
//				TaskDialog.Show("Error", "Failed to create wall.");
			}
		}

		public void addCeiling()
		{
			Document doc = this.ActiveUIDocument.Document;

			using (Transaction trans = new Transaction(doc))
			{
				trans.Start("Creating ceiling");
				try
				{
					
					
					IEnumerable<Space> spaces = new FilteredElementCollector(doc, doc.ActiveView.Id)
						.WhereElementIsNotElementType()
						.OfCategory(BuiltInCategory.OST_MEPSpaces)
						.Cast<Space>();
					
//					IEnumerable<Room> rooms = new FilteredElementCollector(doc, doc.ActiveView.Id)
//						.WhereElementIsNotElementType()
//						.OfCategory(BuiltInCategory.OST_Rooms)
//						.Cast<Room>();
					
					ElementId levelId = Level.GetNearestLevelId(doc, 3.0);

					foreach (Space space in spaces)
					{
						using (GeometryElement geometryElement = space.get_Geometry(new Options())) // Use using statement
						{
							List<XYZ> points = new List<XYZ>();
							CurveLoop profile = new CurveLoop();

							foreach (GeometryObject geometryObject in geometryElement)
							{
								int count = 0;
								Solid solid = geometryObject as Solid;
								double height = 0;

								if (solid != null)
								{
									foreach (Face face in solid.Faces)//It takes each face of enclosed space like roof, ceiling and walls
									{
										IList<XYZ> edgePoints = face.Triangulate().Vertices;

										if (count >= 2) // To avoid the roofs and floor surface
										{
											if (height == 0)
											{
												height = edgePoints[1].Z - edgePoints[2].Z;
											}
											XYZ point = new XYZ(edgePoints[0].X, edgePoints[0].Y, 0);
											points.Add(point);
										}
										count++;
									}
								}

								for (int i = 0; i < points.Count() - 1; i++)
								{
									profile.Append(Line.CreateBound(points[i], points[i + 1]));

									if (i == points.Count() - 2)
									{
										profile.Append(Line.CreateBound(points[i + 1], points[0]));
									}
								}
//								HostObjAttributes host = new HostObjAttributes();
//								RoofBase rf = null;
//								FootPrintRoof.
//								SlabShapeEditor slab = null;
//								slab.AddPoints(points);
								

								var ceiling = Ceiling.Create(doc, new List<CurveLoop> { profile }, ElementId.InvalidElementId, levelId);
								Parameter param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
								param.Set(height);//To set height of the ceiling according to wall height
							}

						}
					}
				}
				catch (Exception ex)
				{
					// Handle exceptions
					TaskDialog.Show("Error", ex.Message);
					trans.RollBack();
				}

				trans.Commit();
			}
			
		}
		
		public void SetRoomLimitOffset()//Code of add ceiling
		{
			
			Document doc = this.ActiveUIDocument.Document;

			using (Transaction transaction = new Transaction(doc))
			{
				try
				{
					transaction.Start("Creating Ceiling");
					
					IEnumerable<Room> rooms = new FilteredElementCollector(doc, doc.ActiveView.Id)
						.WhereElementIsNotElementType()
						.OfCategory(BuiltInCategory.OST_Rooms)
						.Cast<Room>();
					
					foreach (Room room in rooms)
					{
						using (GeometryElement geometryElement = room.get_Geometry(new Options())) // Use using statement
						{
							foreach (GeometryObject geometryObject in geometryElement)
							{
								int count = 0;
								Solid solid = geometryObject as Solid;
								double height = 0;

								if (solid != null)
								{
									foreach (Face face in solid.Faces)//It takes each face of enclosed space like roof, ceiling and walls
									{
										IList<XYZ> edgePoints = face.Triangulate().Vertices;

										if (count >= 2) // To avoid the roofs and floor surface
										{
											if (height == 0)
											{
												height = edgePoints[1].Z - edgePoints[2].Z;
											}
											
										}
										count++;
									}
								}
								if (height != room.UnboundedHeight) {
									room.LimitOffset = height;
								}
							}
						}
						
					}
				}
				catch (Exception ex)
				{
					// Handle exceptions
					TaskDialog.Show("Error", ex.Message);
					transaction.RollBack();
				}

				transaction.Commit();
			}
			
		}
		public void addRoofs()
		{
			Document doc = this.ActiveUIDocument.Document;

			using (Transaction trans = new Transaction(doc))
			{
				trans.Start("Creating ceiling");
				try
				{
					
					
					IEnumerable<Space> spaces = new FilteredElementCollector(doc, doc.ActiveView.Id)
						.WhereElementIsNotElementType()
						.OfCategory(BuiltInCategory.OST_MEPSpaces)
						.Cast<Space>();
					
//					IEnumerable<Room> rooms = new FilteredElementCollector(doc, doc.ActiveView.Id)
//						.WhereElementIsNotElementType()
//						.OfCategory(BuiltInCategory.OST_Rooms)
//						.Cast<Room>();
					
					ElementId levelId = Level.GetNearestLevelId(doc, 3.0);

					foreach (Space space in spaces)
					{
						using (GeometryElement geometryElement = space.get_Geometry(new Options())) // Use using statement
						{
							List<XYZ> points = new List<XYZ>();
							CurveLoop profile = new CurveLoop();

							foreach (GeometryObject geometryObject in geometryElement)
							{
								int count = 0;
								Solid solid = geometryObject as Solid;
								double height = 0;

								if (solid != null)
								{
									foreach (Face face in solid.Faces)//It takes each face of enclosed space like roof, ceiling and walls
									{
										IList<XYZ> edgePoints = face.Triangulate().Vertices;

										if (count >= 2) // To avoid the roofs and floor surface
										{
											if (height == 0)
											{
												height = edgePoints[1].Z - edgePoints[2].Z;
											}
											XYZ point = new XYZ(edgePoints[0].X, edgePoints[0].Y, 0);
											points.Add(point);
										}
										count++;
									}
								}

								for (int i = 0; i < points.Count() - 1; i++)
								{
									profile.Append(Line.CreateBound(points[i], points[i + 1]));

									if (i == points.Count() - 2)
									{
										profile.Append(Line.CreateBound(points[i + 1], points[0]));
									}
								}
//								HostObjAttributes host = new HostObjAttributes();
								//RoofBase rf = null;
//								FootPrintRoof.
//								rf.RoofType = RoofType.
								
								SlabShapeEditor slab = null;
								slab.AddPoints(points);
							}

						}
					}
				}
				catch (Exception ex)
				{
					// Handle exceptions
					TaskDialog.Show("Error", ex.Message);
					trans.RollBack();
				}

				trans.Commit();
			}
		}
		public void CreateMultiFloor()
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
		
		public void AddDoors()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			using (Transaction transaction = new Transaction(doc, "WallCreation"))
			{
				transaction.Start("Create Wall");
				
				// Get a floor type for floor creation
				ElementId floorTypeId = Floor.GetDefaultFloorType(doc, false);
				CurveLoop profile = new CurveLoop();
				
				// Create a wall
				WallType wallType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault();
				//Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
				ElementId levelId1 = Level.GetNearestLevelId(doc, 5.00);
				
				FilteredElementCollector doorFilter = new FilteredElementCollector(doc);
				ICollection<Element> collection = doorFilter.OfClass(typeof(FamilySymbol))
					.OfCategory(BuiltInCategory.OST_Doors)
					.ToElements();
				
				IEnumerator<Element> iterator =  collection.GetEnumerator();
				
				Wall wall1 = Wall.Create(doc, Line.CreateBound(new XYZ(0.0, 0.0, 0.0),new XYZ(0.0, 50.0, 0.0)), wallType.Id, levelId1, 10, 0, true, false);
				
				// get wall's level for door creation
				Level level = doc.GetElement(wall1.LevelId) as Level;
				
				FilteredElementCollector collector = new FilteredElementCollector(doc);

				double x = 0, y = 5.0, z = 0;
				
				FamilySymbol symbol = collector.OfCategory(BuiltInCategory.OST_Doors).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().FirstOrDefault();
				
				XYZ location = new XYZ(x, y, z);
				FamilyInstance instance = doc.Create.NewFamilyInstance(location, symbol, wall1, level, StructuralType.NonStructural);
				y += 5;

				
				
				transaction.Commit();
			}
		}
		
		public void ReadJsonFile()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			using(Transaction trans = new Transaction(doc, "wallCreation")){
				trans.Start("Creating wall in Revit");
				string fileName = "E:\\multiFloorBuilding.json";
				StreamReader streamReader = new StreamReader(fileName);
				string json = streamReader.ReadToEnd();
				streamReader.Close();
				
				JObject builderRoot = JsonConvert.DeserializeObject<JObject>(json);
				List<XYZ> vertexList = new List<XYZ>();

				for (int i = 1; i <= 2; i++) {
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
					
					//Creating wall and store information with their names
					foreach (var item in lines) {
						var data = item.Value;
						
						string wallName = data["id"].ToObject<string>();
						string vertexId1 = data["vertices"][0].ToObject<string>();
						string vertexId2 = data["vertices"][1].ToObject<string>();
						
						double vertex1_x = builderRoot["layers"][layerName]["vertices"][vertexId1]["x"].ToObject<double>();
//						double vertex1_x = GetData.getVertex(vertexId1, builderRoot, layerName);
						double vertex1_y = builderRoot["layers"][layerName]["vertices"][vertexId1]["y"].ToObject<double>();
						
						double vertex2_x = builderRoot["layers"][layerName]["vertices"][vertexId2]["x"].ToObject<double>();
						double vertex2_y = builderRoot["layers"][layerName]["vertices"][vertexId2]["y"].ToObject<double>();
						
						XYZ vertex1 = new XYZ(vertex1_x, vertex1_y, 0.0);
						XYZ vertex2 = new XYZ(vertex2_x, vertex2_y, 0.0);
						
						Wall wall= Wall.Create(doc, Line.CreateBound(vertex1, vertex2), wallType.Id, level_Id, 10, 0, true, false);
						
						//Set wall thickness
						var wallCompoundStructure = wallType.GetCompoundStructure();
						double wallThickness = data["properties"]["thickness"].ToObject<double>();
						wallCompoundStructure.SetLayerWidth(0, wallThickness*0.0833333);//inches to feet
						wallType.SetCompoundStructure(wallCompoundStructure);
						wallDictionary.Add(wallName, wall);
					}
					
					foreach (var item in holes) {
						var data = item.Value;
						string holeName = data["id"].ToObject<string>();
						string lineName = builderRoot["layers"][layerName]["holes"][holeName]["line"].ToObject<string>();
						double z = builderRoot["layers"][layerName]["holes"][holeName]["properties"]["altitude"].ToObject<double>();
						XYZ location = new XYZ(0,0,0);
						
						double offset = builderRoot["layers"][layerName]["holes"][holeName]["offset"].ToObject<double>();
						double lineLength = builderRoot["layers"][layerName]["lines"][lineName]["properties"]["length"].ToObject<double>();

						string vertex1a = builderRoot["layers"][layerName]["lines"][lineName]["vertices"][0].ToObject<string>();
						string vertex2a = builderRoot["layers"][layerName]["lines"][lineName]["vertices"][1].ToObject<string>();
						
						double x1 = builderRoot["layers"][layerName]["vertices"][vertex1a]["x"].ToObject<double>();
						double y1 = builderRoot["layers"][layerName]["vertices"][vertex1a]["y"].ToObject<double>();

						double x2 = builderRoot["layers"][layerName]["vertices"][vertex2a]["x"].ToObject<double>();
						double y2 = builderRoot["layers"][layerName]["vertices"][vertex2a]["y"].ToObject<double>();
						
						if ((x1<x2 && y1==y2) || (x2<x1 && y1==y2)) {
							location = new XYZ(x1+offset*(lineLength/20), y1, z);//We divide length by 20 as in json factor due to AHC reader
						}
						else if ((y1<y2 && x1==x2 ) || (y2<y1 && x1==x2)) {
							location = new XYZ(x1, y1+offset*(lineLength/20), z);
						}
						
						//Check whether the hole is Window or Door
						if (data["type"].ToString() == "window") {
							FilteredElementCollector collectorWindows = new FilteredElementCollector(doc);
							FamilySymbol symbolWindow = collectorWindows
								.OfCategory(BuiltInCategory.OST_Windows)
								.OfClass(typeof(FamilySymbol))
								.Cast<FamilySymbol>().FirstOrDefault();
							
							Wall specificWall = wallDictionary[lineName];
							FamilyInstance windowInstance = doc.Create.NewFamilyInstance(location, symbolWindow,specificWall, level, StructuralType.NonStructural);
							
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
							FamilySymbol symbolDoor = collectorDoor
								.OfCategory(BuiltInCategory.OST_Doors)
								.OfClass(typeof(FamilySymbol))
								.Cast<FamilySymbol>()
								.FirstOrDefault(sym => sym.FamilyName.Contains("double ")); // Assuming the family name of the double door contains "Double"
							string name = symbolDoor.FamilyName;
							
							TaskDialog.Show("Result", name);

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
					
					//To create floors
					Floor.Create(doc, new List<CurveLoop> { profile }, floorTypeId, level_Id);
					
					if(i==2){
						//To create ceiling
						var ceiling = Ceiling.Create(doc, new List<CurveLoop> { profile }, ElementId.InvalidElementId, level_Id);
						Parameter param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
						param.Set(10.00);
					}
				}
				
				trans.Commit();
			}
		}
//		public void BIMConfigurator()
//		{
//		}
	}
}