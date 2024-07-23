using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Text;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using System.Diagnostics;

namespace Kunal2
{
	[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
	[Autodesk.Revit.DB.Macros.AddInId("6E3669CC-110D-462B-BC78-D1E6951E3C4D")]
	public partial class ThisApplication
	{
		FailureProcessor failureProcessor = new FailureProcessor();
		
		Dictionary<string, FamilySymbol> mDocumentFamilySymbols = new Dictionary<string, FamilySymbol>();
		
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
		
		public void ReadBuildingJson()
		{
			Stopwatch stopwatch = Stopwatch.StartNew();
			Document doc = this.ActiveUIDocument.Document;
			string filePath = "E:\\JSON\\site 99.json";			
			JObject builderRoot = ReadJson(filePath);			
			var levels = builderRoot["siteProperties"]["Level"].ToObject<JObject>();
			editLevels(doc, levels);
			
			try
			{
				int buildingCount = builderRoot["site"]["Buildings"].Count();
				for (int i = 1; i <= buildingCount; ++i)
				{
					Building building = new Building();
					string buildingName = "" + i;
					ProcessBuildingLevels(doc, builderRoot, building, buildingName);
				}
			}
			catch (Exception ex)
			{
				TaskDialog.Show("Exception ReadBuildingJson()", ex.StackTrace);
			}
			stopwatch.Stop();//Time counting stopped
			double endTime = (double)stopwatch.ElapsedMilliseconds/1000;
			TaskDialog.Show("TimeTaken", "Time taken to execute InitiateDataCollection:\n"+ endTime + " Seconds");
		}
		
		private JObject ReadJson(string filePath)
		{
			string json = null;
			using (StreamReader streamReader = new StreamReader(filePath))
			{
				json = streamReader.ReadToEnd();
			}
			return JsonConvert.DeserializeObject<JObject>(json);
		}
		
		public void editLevels(Document doc, JObject levels)
		{
			DeleteLevels(doc);//Deleting the existing levels
			using(Transaction trans = new Transaction(doc))
			{
				trans.Start("Creating new levels as per Json");
				try{
					foreach (var level in levels) {
						var data = level.Value;
						var item = level.Key;
						double elevation = data["Elevation"].ToObject<double>();
						string name = item.ToString();
						
						// Begin to create a level
						Level newLevel = Level.Create(doc, elevation);
						if (null == newLevel)
						{
							throw new Exception("Create a new level failed.");
						}
						
						// Change the level name
						newLevel.Name = "Level "+name;
						
						// Get a ViewFamilyType for creating a floor plan
						ViewFamilyType viewFamilyType = null;
						FilteredElementCollector collector = new FilteredElementCollector(doc);
						foreach (Autodesk.Revit.DB.Element e in collector.OfClass(typeof(ViewFamilyType)))
						{
							ViewFamilyType vft = e as ViewFamilyType;
							if (vft.ViewFamily == ViewFamily.FloorPlan)
							{
								viewFamilyType = vft;
								break;
							}
						}
						
						// Create a new view plan associated with the level
						if (viewFamilyType != null)
						{
							ViewPlan viewPlan = ViewPlan.Create(doc, viewFamilyType.Id, newLevel.Id);
							if (viewPlan == null)
							{
								throw new Exception("Failed to create view for the new level.");
							}
						}
						else
						{
							throw new Exception("No ViewFamilyType found for creating a floor plan.");
						}
					}
				}
				catch(Exception ex){
					TaskDialog.Show("Exception addLevel()", ex.StackTrace);
				}
				trans.Commit();
			}
		}
		
		public void DeleteLevels(Document doc)
		{
			using(Transaction tx = new Transaction(doc, "Delete Levels"))
			{
				tx.Start();
				
				int deleted = 0;
				FilteredElementCollector collector = new FilteredElementCollector(doc);
				ICollection<Autodesk.Revit.DB.Element> levels = collector.OfClass(typeof(Level)).ToElements();
				List<ElementId> elementsToBeDeleted = new List<ElementId>();
				
				foreach(Autodesk.Revit.DB.Element element in levels)
				{
					elementsToBeDeleted.Add(element.Id);
					deleted++;
				}
				doc.Delete(elementsToBeDeleted);
				
				tx.Commit();
			}
		}

		private void ProcessBuildingLevels(Document doc, JObject builderRoot, Building building, string buildingName)
		{
			int levelsCount = builderRoot["site"]["Buildings"][buildingName]["Levels"].Count();
			var levelsData = builderRoot["site"]["Buildings"][buildingName]["Levels"].ToObject<JObject>();

			foreach(var levelData in levelsData){
				string levelName = levelData.Key.ToString();
				double elevation = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["elevation"].ToObject<double>();
				var vertices = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["vertices"].ToObject<JObject>();
				var lines = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["lines"].ToObject<JObject>();
				var areas = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["areas"].ToObject<JObject>();
				var holes = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["holes"].ToObject<JObject>();
				var slabs = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["slabs"].ToObject<JObject>();
				var roofSlabs = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["roofSlabs"].ToObject<JObject>();
				
				//creating windows and door types from JSON and use it as per requirement
				List<float> WL1 = new List<float>();
				List<float> WL2 = new List<float>();
				List<float> DL1 = new List<float>();
				List<float> DL2 = new List<float>();
				
				getWindowDoorDimensions(holes, out WL1, out WL2, out DL1, out DL2);
				CreateWindowTypes(doc, WL1, WL2, "double ");
				CreateDoorTypes(doc, DL1, DL2, "double ");
				
				Dictionary<String, Vertex> vertexMap;
				addVertices(vertices, out vertexMap, elevation);

				Dictionary<String, WallComponent> wallComponentMap;
				addHoles(holes, out wallComponentMap);

				List<Wall> wallList = new List<Wall>();
				addWalls(lines, out wallList, vertexMap, wallComponentMap);

				building.SlabVertexList = new List<List<UV>>();
				List<UV> slabVertexList;
				List<XYZ> slabHolePointsList;
				addSlabs(doc, slabs, out slabVertexList, out slabHolePointsList, elevation);
				addSlabs(doc, roofSlabs, out slabVertexList, out slabHolePointsList, elevation);
				
				List<Room> roomList = new List<Room>();
				List<XYZ> vertexList1 = new List<XYZ>();
				
				building.VertexList = vertexMap.Values.ToList();
				building.WallComponent = wallComponentMap.Values.ToList();
				building.WallList = wallList;
				building.RoomList = roomList;
				
				createBuildingRevitFile(building, doc, elevation, elevation, levelsCount, Convert.ToInt32(levelName));
				createRooms(doc, areas, out roomList, out vertexList1, elevation);//Creates Room and adds RoomTags to rooms
			}
		}
		
		public static void getWindowDoorDimensions(JObject holes, out List<float> WL1, out List<float> WL2, out List<float> DL1, out List<float> DL2)
		{
			WL1 = new List<float>();
			WL2 = new List<float>();
			DL1 = new List<float>();
			DL2 = new List<float>();
			
			foreach (var hole in holes)
			{
				var data = hole.Value;
				string holeType = data["type"].ToString();
				string style = data["properties"]["style"].ToString();
				
				if (holeType.Contains("Door/Window Assembly"))
				{
					if (style.Contains("Window"))
					{
						var windowWidth = data["properties"]["width"].ToObject<float>();
						WL1.Add((float)Math.Round(windowWidth, 2));
						
						var windowHeight = data["properties"]["height"].ToObject<float>();
						WL2.Add((float)Math.Round(windowHeight, 2));
					}
				}
				else if (holeType.Contains("Window"))
				{
					var windowWidth = data["properties"]["width"].ToObject<float>();
					WL1.Add((float)Math.Round(windowWidth, 2));
					
					var windowHeight = data["properties"]["height"].ToObject<float>();
					WL2.Add((float)Math.Round(windowHeight, 2));
				}
				else if (holeType.Contains("Door"))
				{
					var doorWidth = data["properties"]["width"].ToObject<float>();
					DL1.Add((float)Math.Round(doorWidth, 2));
					
					var doorHeight = data["properties"]["height"].ToObject<float>();
					DL2.Add((float)Math.Round(doorHeight, 2));
				}
			}
		}
		
		private Result CreateWindowTypes(Document doc, List<float> L1, List<float> L2, string familyName)
		{
			return CreateTypes<WindowType>(doc, L1, L2, BuiltInCategory.OST_Windows, familyName);
		}

		private Result CreateDoorTypes(Document doc, List<float> L1, List<float> L2, string familyName)
		{
			return CreateTypes<DoorType>(doc, L1, L2, BuiltInCategory.OST_Doors, familyName);
		}
		
		private Result CreateTypes<T>(Document doc, List<float> L1, List<float> L2, BuiltInCategory category, string familyName) where T : HoleType
		{
			if (L1.Count != L2.Count)
				throw new ArgumentException("The L1 and L2 lists must have the same number of elements.");

			List<T> all = new List<T>();
			
			try
			{
				for (int i = 0; i < L1.Count(); ++i)
				{
					float l1 = L1[i];
					float l2 = L2[i];
					
					all.Add((T)Activator.CreateInstance(typeof(T), l1, l2));
				}
				
			}
			catch(Exception e)
			{
				TaskDialog.Show("Exception Result CreateTypes<T>", e.Message);
			}
			
			// unique elements in the list
			all = all.Distinct().ToList();
			FilteredElementCollector symbols = new FilteredElementCollector(doc).WhereElementIsElementType().OfCategory(category);
			IEnumerable<FamilySymbol> existing = symbols.OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>();
			List<FamilySymbol> existingList = existing.ToList();
			FamilySymbol defaultSymbol = existing.FirstOrDefault();

			if (existingList.Count == 0)
			{
				return Result.Cancelled;
			}

			List<T> AlreadyExists = new List<T>();
			List<T> ToBeMade = new List<T>();

			for (int i = 0; i < all.Count; ++i)
			{
				string typeName = (string)typeof(T).GetProperty("Name").GetValue(all[i]);
				double w = all[i].W;
				double h = all[i].H;

				FamilySymbol fs = existing.FirstOrDefault(x => x.Name == typeName);
				var fsm = ToBeMade.FirstOrDefault(x => ((x.H == h)&&(x.W == w)));
				
				if (fs == null && fsm == null)
				{
					ToBeMade.Add(all[i]);
				}
				else
				{
					AlreadyExists.Add(all[i]);
				}
			}

			if (ToBeMade.Count == 0)
			{
				return Result.Cancelled;
			}
			
			List<ElementType> sys = new List<ElementType>();

			using (Transaction tx = new Transaction(doc))
			{
				try
				{
					if (tx.Start("Make types") == TransactionStatus.Started)
					{
						tx.SetFailureHandlingOptions(tx.GetFailureHandlingOptions().SetFailuresPreprocessor(failureProcessor));
						FamilySymbol first = existing.First().IsActive ? existing.First() : defaultSymbol;

						foreach (T ct in ToBeMade)
						{
							string typeName = (string)typeof(T).GetProperty("Name").GetValue(ct);
							ElementType et = first.Duplicate(typeName);

							double height = ct.H;
							double width = ct.W;

							Parameter hParam = et.LookupParameter("Height");
							Parameter wParam = et.LookupParameter("Width");

							if (!hParam.IsReadOnly)
							{
								hParam.Set(height);
							}

							if (!wParam.IsReadOnly)
							{
								wParam.Set(width);
							}
							sys.Add(et);
						}
					}
				}
				catch (Exception ex)
				{
					TaskDialog.Show("err", ex.Message);
				}
				tx.Commit();
			}
			return Result.Succeeded;
		}

		private void addVertices(JObject vertices, out Dictionary<String, Vertex> vertexMap, double levelAltitude)
		{
			vertexMap = new Dictionary<string, Vertex>();
			foreach (var item in vertices)
			{
				var data = item.Value;
				string id = item.Key;
				double x = data["x"].ToObject<double>();
				double y = data["y"].ToObject<double>();
				double z = levelAltitude;
				Vertex vertex = new Vertex(x, y, z, id, id);
				vertexMap.Add(id, vertex);
				vertex.WallList = new List<Wall>();
			}
		}
		
		private void addHoles(JObject holes, out Dictionary<String, WallComponent> wallComponentMap)
		{
			wallComponentMap = new Dictionary<String, WallComponent>();
			foreach (var item in holes)
			{
				var data = item.Value.ToObject<JObject>();
				string id = item.Key;
				string name = data["type"].ToString();
				string type = data["type"].ToString();
				Double offset = data["offset"].ToObject<Double>(); // This value is Unit less so no need to apply unit conversion factor
				double startPoint_x = Double.Parse(data["startPoint"]["X"].ToString());
				double startPoint_y = (double)data["startPoint"]["Y"];
				double startPoint_z = (double)data["startPoint"]["Z"];
				
				XYZ startPoint = new XYZ(startPoint_x,startPoint_y, startPoint_z);
				double endPoint_x = data["startPoint"]["X"].ToObject<double>();
				double endPoint_y = data["startPoint"]["Y"].ToObject<double>();
				double endPoint_z = data["startPoint"]["Z"].ToObject<double>();
				XYZ endPoint = new XYZ(endPoint_x,endPoint_y,endPoint_z);
				XYZ midPoint = (endPoint - startPoint)/2;
				double normalX = 0.0;
				double normalY = 0.0;
				double normalZ = 0.0;
				
				//Consider Normal Only when hole is other than Door or window
				if (false == (type.ToLower().Contains("door") || type.ToLower().Contains("window")))
				{
					normalX = data["normal"]["x"].ToObject<Double>();
					normalY = data["normal"]["y"].ToObject<Double>();
					normalZ = data["normal"]["z"].ToObject<Double>();
				}
				
				WallComponentType componentType = WallComponentType.Door;
				componentType = getComponentType(type, out componentType);
				
				Dictionary<String, Object> properties = new Dictionary<String, Object>();
				var itemsprop = data["properties"].ToObject<JObject>();
				foreach (var prop in itemsprop)
				{
					Object value = prop.Value;
					if (prop.Key == "altitude")
					{
						value = (prop.Value.ToObject<double>()).ToString();
					}
					properties.Add(prop.Key, value);
				}
				WallComponent wallComponent = new WallComponent(componentType, name, id, properties, offset, normalX, normalY, normalZ);
				wallComponentMap.Add(id, wallComponent);
			}
		}
		
		private WallComponentType getComponentType(string type, out WallComponentType componentType)
		{
			componentType = WallComponentType.Door;
			switch (type)
			{
				case "Door":
					componentType = WallComponentType.Door;
					break;
				case "Window":
					componentType = WallComponentType.Window;
					break;
				case "Door/Window Assembly":
					componentType = WallComponentType.Window;
					break;
			}
			return componentType;
		}
		
		private void addWalls(JObject lines, out List<Wall> wallList, Dictionary<String, Vertex> vertexMap, Dictionary<String, WallComponent> wallComponentMap)
		{
			wallList = new List<Wall>();
			try{
				foreach (var item in lines)
				{
					var data = item.Value;
					string id = item.Key;
					
					string name = data["type"].ToString();
					if(data["vertices"] == null)
					{
						TaskDialog.Show("Vertices Id", id);
						continue;
					}
					
					Wall wall = CreateWall(data, vertexMap, name, id);
					List<WallComponent> wallComponentList = new List<WallComponent>();
					if(data["holes"].Count() != 0)
					{
						var wallComponentIDList = data["holes"] as JArray;
						foreach (var wallComponentIDItem in wallComponentIDList)
						{
							string wallComponentID = wallComponentIDItem.ToObject<String>();
							if (wallComponentMap.ContainsKey(wallComponentID))
							{
								// It's Component on Wall such as Door or Window
								WallComponent wallComponent = wallComponentMap[wallComponentID];
								wallComponent.Wall = wall;
								wallComponentList.Add(wallComponent);
							}
						}
						//Add Wall Diffuser List
						wall.ItemsOnWallIDList = wallComponentList;
					}
					wallList.Add(wall);
				}
			}
			catch(Exception ex)
			{
				string e =  "Error Messaage : " + ex.StackTrace;
				TaskDialog.Show("Error", e);
			}
		}
		
		Wall CreateWall(JToken data, Dictionary<string, Vertex> vertexMap, string name, string id)
		{
			// Extract Wall Type
			WallType wallType = data["type"].ToObject<string>().Contains("Curtain") ? WallType.Curtain : WallType.Exposed_Wall;
			
			// Extract Wall Properties
			var properties = data["properties"].ToObject<JObject>();
			double height = properties["height"].ToObject<double>();
			double thickness = properties["thickness"].ToObject<double>() != 0 ? properties["thickness"].ToObject<double>() : 0;
			double baseOffset = properties["baseOffset"].ToObject<double>();
			
			// Extract Vertices
			var vertexIDList = data["vertices"] as JArray;
			string vertexID1 = vertexIDList[0].ToObject<string>();
			string vertexID2 = vertexIDList[1].ToObject<string>();
			
			Vertex v1 = vertexMap[vertexID1];
			Vertex v2 = vertexMap[vertexID2];
			
			List<Vertex> vertexList = new List<Vertex> {v1, v2};
			
			Wall wall;
			Arc arcWall = null;
			
			// Check for Curtain Arc Wall
			if (data["curatainArcWall"] != null && data["curatainArcWall"].HasValues)
			{
				var curtainArcWallPoints = data["curatainArcWall"].ToObject<JObject>();
				XYZ center = new XYZ(
					(double)curtainArcWallPoints["Center"]["X"],
					(double)curtainArcWallPoints["Center"]["Y"],
					(double)curtainArcWallPoints["Center"]["Z"]);
				
				double startAngle = (double)curtainArcWallPoints["StartAngle"];
				double endAngle = (double)curtainArcWallPoints["EndAngle"];
				double radius = (double)curtainArcWallPoints["Radius"];
				
				XYZ xAxis = new XYZ(
					(double)curtainArcWallPoints["Xaxis"]["X"],
					(double)curtainArcWallPoints["Xaxis"]["Y"],
					(double)curtainArcWallPoints["Xaxis"]["Z"]);
				
				XYZ yAxis = new XYZ(
					(double)curtainArcWallPoints["Yaxis"]["X"],
					(double)curtainArcWallPoints["Yaxis"]["Y"],
					(double)curtainArcWallPoints["Yaxis"]["Z"]);
				
				arcWall = Arc.Create(center, radius, startAngle, endAngle, xAxis, yAxis);
			}
			// Check for Standard Arc Wall
			else if (data["ArcWall"] != null && data["ArcWall"].HasValues)
			{
				var standardArcWallPoints = data["ArcWall"].ToObject<JObject>();
				XYZ startPoint = new XYZ(
					(double)standardArcWallPoints["StartPoint"]["X"],
					(double)standardArcWallPoints["StartPoint"]["Y"],
					(double)standardArcWallPoints["StartPoint"]["Z"]);
				
				XYZ pointOnArc = new XYZ(
					(double)standardArcWallPoints["PointOnArc"]["X"],
					(double)standardArcWallPoints["PointOnArc"]["Y"],
					(double)standardArcWallPoints["PointOnArc"]["Z"]);
				
				XYZ endPoint = new XYZ(
					(double)standardArcWallPoints["EndPoint"]["X"],
					(double)standardArcWallPoints["EndPoint"]["Y"],
					(double)standardArcWallPoints["EndPoint"]["Z"]);
				
				arcWall = Arc.Create(startPoint, endPoint, pointOnArc);
			}
			
			// Create Wall
			wall = new Wall(wallType, name, id, vertexList, height, baseOffset, arcWall, thickness);
			wall.PropertyDictionary = properties.ToObject<Dictionary<string, object>>();
			v1.WallList.Add(wall);
			v2.WallList.Add(wall);
			
			return wall;
		}

		private void createRooms(Document doc, JObject areas, out List<Room> roomList, out List<XYZ> vertexList1, double elevation)
		{
			vertexList1 = new List<XYZ>();
			roomList = new List<Room>();
			
			using (Transaction trans = new Transaction(doc)) {
				
				trans.Start("Start");
				trans.SetFailureHandlingOptions(trans.GetFailureHandlingOptions().SetFailuresPreprocessor(failureProcessor));
				
				foreach (var item in areas)
				{
					var data = item.Value;
					string id = item.Key;
					
					string name = data["name"].ToString();
					double fluidPoint_x = data["fluidPoint"]["x"].ToObject<double>();
					double fluidPoint_y = data["fluidPoint"]["y"].ToObject<double>();
					
					UV fluidPoint= new UV(fluidPoint_x, fluidPoint_y);
					
					XYZ vertex1 = new XYZ(fluidPoint_x, fluidPoint_y, 0.0);
					vertexList1.Add(vertex1);
					
					ElementId levelId = Level.GetNearestLevelId(doc, elevation);
					
					Autodesk.Revit.DB.Element element = doc.GetElement(levelId);
					Level level = element as Level;
					
					ViewPlan view = GetViewForLevel(doc, level);
					
					Autodesk.Revit.DB.Architecture.Room room = doc.Create.NewRoom(level, fluidPoint);
					RoomTag roomTag = doc.Create.NewRoomTag(new LinkElementId(room.Id), fluidPoint, view.Id);
				}
				trans.Commit();
			}
		}
		
		private void addSlabs(Document doc, JObject slabs, out List<UV> slabVertexList, out List<XYZ> slabHolePointsList, double elevation)
		{
			slabVertexList = new List<UV>();
			slabHolePointsList = new List<XYZ>();
			
			FloorType floorType = createFloorType(doc);
			CeilingType ceilingType = createCeilingType(doc);
			
			ElementId levelId1 = Level.GetNearestLevelId(doc, elevation);
			
			foreach (var item in slabs)
			{
				var data = item.Value;
				string id = item.Key;
				var vertexList = data["slabLoop"] as JArray;
				var floorHoles = data["holes"] as JArray;
				var lowPoint =  data["lowPoint"].ToObject<double>();
				var highPoint = data["highPoint"].ToObject<double>();
				
				foreach (var v in vertexList)
				{
					double x = v[0].ToObject<double>();
					double y = v[1].ToObject<double>();
					
					UV slabPoint = new UV(x, y);
					slabVertexList.Add(slabPoint);
				}
				CurveLoop profile = new CurveLoop();
				CurveArray holePoints = new CurveArray();
				Floor floor = null;
				
				try  
				{
					using (Transaction trans = new Transaction(doc))
					{
						trans.Start("started");
						trans.SetFailureHandlingOptions(trans.GetFailureHandlingOptions().SetFailuresPreprocessor(failureProcessor));
						
						for (int k = 0; k < slabVertexList.Count; k++)
						{
							var vertex1 = slabVertexList[k];
							XYZ p1 = new XYZ(vertex1.U, vertex1.V, 0.0);
							int index = ((k + 1) < slabVertexList.Count) ? k + 1 : 0;
							var vertex2 = slabVertexList[index];
							XYZ p2 = new XYZ(vertex2.U, vertex2.V, 0.0);
							var c = Line.CreateBound(p1, p2) as Curve;
							holePoints.Append(c);
							profile.Append(c);
						}
						
						if (highPoint <= 0.0 || lowPoint <= 0.0) {
							floor = Floor.Create(doc, new List<CurveLoop> { profile }, floorType.Id, levelId1);
						}
						else{
							Ceiling ceiling = Ceiling.Create(doc, new List<CurveLoop> { profile }, ceilingType.Id, levelId1);
							Parameter param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
							param.Set(highPoint);
						}
						slabVertexList.Clear();
						trans.Commit();
					}
				} 
				catch (Exception ex) 
				{
					TaskDialog.Show("Floor Exception", ex.Message);
				}
				
				
				using (Transaction trans = new Transaction(doc))
				{
					trans.Start("started");
					trans.SetFailureHandlingOptions(trans.GetFailureHandlingOptions().SetFailuresPreprocessor(failureProcessor));
					FilteredElementCollector collector = new FilteredElementCollector(doc);
					collector.OfCategory(BuiltInCategory.OST_Levels);
					collector.OfClass(typeof(Level));
					
					RoofType roofType = new FilteredElementCollector(doc).OfClass(typeof(RoofType)).FirstOrDefault() as RoofType;
					ElementId levelId = Level.GetNearestLevelId(doc, elevation);
					
					Autodesk.Revit.DB.Element element = doc.GetElement(levelId);
					Level level = element as Level;
					
					//Creating floor holes
					if (floorHoles.Count() != 0) {
						foreach (var v in floorHoles) {
							foreach (var e in v) {
								double x = e[0].ToObject<double>();
								double y = e[1].ToObject<double>();
								
								XYZ floorHolePoint = new XYZ(x, y, 0);
								slabHolePointsList.Add(floorHolePoint);
							}
							List<Line> openingLines = new List<Line>();
							for (int k = 0; k < slabHolePointsList.Count(); k++) {
								openingLines.Add(Line.CreateBound(slabHolePointsList[k], slabHolePointsList[(k+1) % slabHolePointsList.Count()]));
							}
							
							try
							{
								CurveArray openingProfile = doc.Application.Create.NewCurveArray();
								foreach (var line in openingLines){
									openingProfile.Append(line);
								}
								var op3 = doc.Create.NewOpening(floor, openingProfile, true);
								openingLines.Clear();
								slabHolePointsList.Clear();
							}
							catch (Exception ex)
							{
								TaskDialog.Show("Error", ex.Message);
							}							
						}
					}
					trans.Commit();
				}
			}
		}
		
		private ViewPlan GetViewForLevel(Document document, Level level)
		{
			FilteredElementCollector collector = new FilteredElementCollector(document);
			ICollection<Autodesk.Revit.DB.Element> views = collector.OfClass(typeof(ViewPlan)).ToElements();
			foreach (Autodesk.Revit.DB.Element elem in views)
			{
				ViewPlan view = elem as ViewPlan;
				if (view != null && view.Name == level.Name && view.ViewType.ToString() == "FloorPlan")
				{
					return view;
				}
			}
			return null;
		}
		
		private void createBuildingRevitFile(Building building, Document document, double levelAltitude, double altitude, int levelsCount, int j)
		{
			ElementId wallTypeId = document.GetDefaultElementTypeId(ElementTypeGroup.WallType);// replace var with Element Id (statically-typed)
			Dictionary<XYZ, string> centroidRoomNameMap = new Dictionary<XYZ, string>();
			createBuilding(building, document, ref centroidRoomNameMap, levelAltitude, levelsCount, j);
			createComponentsInBuilding(building, document, centroidRoomNameMap);
		}
		
		private void createComponentsInBuilding(Building building, Document document, Dictionary<XYZ, string> centroidRoomNameMap)
		{
			Transaction component = new Transaction(document);// no need to to use "Using", Autodesk disposes of all the methods and data in Execute method.
			{
				component.Start("create componets in building spaces");
				component.SetFailureHandlingOptions(component.GetFailureHandlingOptions().SetFailuresPreprocessor(failureProcessor));
				assignRoomName(document, centroidRoomNameMap);
				
				try{
					PlaceWallComponents(document, building.WallComponent);
				}
				catch(Exception ex)
				{
					Console.WriteLine("Exception occurs at createComponentsInBuilding: " + ex.Message);
				}
				
				component.Commit();
			}
		}
		
		private void assignRoomName(Document document, Dictionary<XYZ, string> centroidRoomNameMap)
		{
			FilteredElementCollector collector = new FilteredElementCollector(document);
			ICollection<Autodesk.Revit.DB.Element> collection = collector.OfClass(typeof(SpatialElement)).ToElements();

			foreach (Autodesk.Revit.DB.Element e in collection)
			{
				Autodesk.Revit.DB.Architecture.Room room = e as Autodesk.Revit.DB.Architecture.Room;
				foreach (var item in centroidRoomNameMap)
				{
					if (room.IsPointInRoom(item.Key))
					{
						centroidRoomNameMap.Remove(item.Key);
						room.Name = item.Value;
						break;
					}
				}
			}
		}
		
		private void createBuilding(Building building, Document document, ref Dictionary<XYZ, string> centroidRoomNameMap, 
		                            							double levelAltitude, int levelsCount, int j)
		{
			// Get a floor type for floor creation
			FloorType floorType = createFloorType(document);
			CeilingType ceilingType = createCeilingType(document);

			XYZ normal = XYZ.BasisZ;

			Transaction transaction = new Transaction(document);
			{
				transaction.Start("create building spaces");
				transaction.SetFailureHandlingOptions(transaction.GetFailureHandlingOptions().SetFailuresPreprocessor(failureProcessor));
				document.Regenerate();
				
				Dictionary<WallType, ElementId> wallTypeElemIdMap = GetWallTypeElementIDMap(document);
				createWalls(document, building.WallList, levelAltitude, wallTypeElemIdMap);
				
				transaction.Commit();
			}
		}
		
		private FloorType createFloorType(Document document)
		{
			FilteredElementCollector collector1 = new FilteredElementCollector(document);
			collector1.OfClass(typeof(FloorType));
			FloorType floorType = collector1.FirstElement() as FloorType;

			return floorType;
		}
		
		private CeilingType createCeilingType(Document document)
		{
			FilteredElementCollector collector1 = new FilteredElementCollector(document);
			collector1.OfClass(typeof(CeilingType));
			CeilingType ceilingType = collector1.FirstElement() as CeilingType;

			return ceilingType;
		}
		
		private Dictionary<WallType, ElementId> GetWallTypeElementIDMap(Document document)
		{
			FilteredElementCollector collector = new FilteredElementCollector(document);
			var rvtWallTypes = collector.OfClass(typeof(Autodesk.Revit.DB.WallType)).ToElements();

			var plannerWallTypes = new List<WallType>()
			{
				WallType.Exposed_Wall,
				WallType.Curtain
			};

			var wallTypeElemIdMap = new Dictionary<WallType, ElementId>();

			foreach (var plannerWallType in plannerWallTypes)
			{
				foreach (var rvtWallType in rvtWallTypes)
				{
					if (rvtWallType.Name.Contains(plannerWallType.ToString()))
					{
						wallTypeElemIdMap.Add(plannerWallType, rvtWallType.Id);
						break;
					}
				}
			}
			return wallTypeElemIdMap;
		}
		
		
		private void createWalls(Document doc,List<Wall> wallList, double levelAltitude, Dictionary<WallType, ElementId> wallTypeIdMap)
		{
			foreach (var wallItem in wallList)
			{
				Autodesk.Revit.DB.Wall wall;
				ChangeWallThickness(doc, wallItem, wallTypeIdMap);
				ElementId levelId = Level.GetNearestLevelId(doc, levelAltitude);
				if(wallItem.WallArc == null)
				{
					XYZ p1 = new XYZ(wallItem.VertextIDList[0].X, wallItem.VertextIDList[0].Y, wallItem.VertextIDList[0].Z);
					XYZ p2 = new XYZ(wallItem.VertextIDList[1].X, wallItem.VertextIDList[1].Y, wallItem.VertextIDList[1].Z);
										
					Curve curve = Line.CreateBound(p1, p2);
					wall = Autodesk.Revit.DB.Wall.Create(doc, curve, wallTypeIdMap[wallItem.Type], 
					                                     levelId, wallItem.WallHeight, wallItem.WallBaseOffset, true, false);
				}
				else
				{
					wall = Autodesk.Revit.DB.Wall.Create(doc, wallItem.WallArc, wallTypeIdMap[wallItem.Type], 
					                                     levelId, wallItem.WallHeight, wallItem.WallBaseOffset, true, false);
				}
				wallItem.RevitID = wall.Id.Value;
			}
		}
		
		private void ChangeWallThickness(Document doc, Wall wall, Dictionary<WallType, ElementId> dict)
		{
			Autodesk.Revit.DB.WallType e = doc.GetElement(dict[wall.Type]) as Autodesk.Revit.DB.WallType ;
			if(wall.WallThickness != e.Width)
			{
				List<Autodesk.Revit.DB.WallType> typeList = new FilteredElementCollector(doc)
					.OfClass(typeof(Autodesk.Revit.DB.WallType)).Cast<Autodesk.Revit.DB.WallType>().ToList();
				
				string name = "Exposed_Wall_" + Math.Round(wall.WallThickness, 2);
				string wallName = wall.Type.ToString();
				
				if (wall.Type.ToString().Contains("Curtain")){
					 name = "Curtain_Wall_" + Math.Round(wall.WallThickness, 2);
				}
				
				
				foreach(var wType in typeList)
				{
					if(wType.Name == name)
					{
						dict[wall.Type] = wType.Id;
						return;
					}
				}
				Autodesk.Revit.DB.WallType newWallType;
				
				try{
					if (!wall.Type.ToString().Contains("Curtain") && (wall.WallThickness != 0)) 
					{
						newWallType =  e.Duplicate(name) as Autodesk.Revit.DB.WallType;
						CompoundStructure compoundStructure = newWallType.GetCompoundStructure();
						
						//add all individual layers to wall compound structure layer
						IList<CompoundStructureLayer> layers = compoundStructure.GetLayers();
						
						layers[0].Width = wall.WallThickness;
						compoundStructure.SetLayers(layers);
						
						newWallType.SetCompoundStructure(compoundStructure);
						dict[wall.Type] = newWallType.Id;
					}
				}
				catch(Exception ex)
				{
					TaskDialog.Show("Exception in changing wall thickness", ex.Message);
				}
			}
		}
		
		
		private void LoadAllFamilies(Document document)
		{
			using (Transaction trans = new Transaction(document, "Load Families"))
			{
				trans.Start();
				trans.SetFailureHandlingOptions(trans.GetFailureHandlingOptions().SetFailuresPreprocessor(failureProcessor));
				mDocumentFamilySymbols.Clear();

				List<FamilySymbol> familySymbolsInDoc = new FilteredElementCollector(document)
					.WherePasses(new ElementClassFilter(typeof(FamilySymbol))).Cast<FamilySymbol>().ToList();
				
				int count = familySymbolsInDoc.Count();
				
				foreach (var familySymbol in familySymbolsInDoc)
				{
					string keyName = familySymbol.Family.Name + '_' + familySymbol.Name;
					mDocumentFamilySymbols.Add(keyName.ToLower(), familySymbol);
				}
				trans.Commit();
			}
		}
		
		private void PlaceWallComponents(Document document, List<WallComponent> wallComponent)
		{
			Dictionary<long, Autodesk.Revit.DB.Wall> wallMap = createWallIDDictionary(document);
			foreach (WallComponent wallComponentItem in wallComponent)
			{
				long wallID = wallComponentItem.Wall.RevitID;
				Autodesk.Revit.DB.Wall wall = wallMap[wallID];
				FamilySymbol symbol = null;

				symbol = GetFamilySymbolForHole(wallComponentItem, document);
				SetFamilySymbolDimensions(symbol, wallComponentItem);
				
				if (symbol != null)
				{
					if (symbol.IsActive == false)
						symbol.Activate();

					List<XYZ> wallEndPoints = getWallEndPointsInFeet(wallComponentItem);

					XYZ normal = new XYZ(wallComponentItem.NormalX, wallComponentItem.NormalY, wallComponentItem.NormalZ).Normalize();
					XYZ location = calculateElementLocation(wallEndPoints, wallComponentItem);

					XYZ locationOnWallFace = ProjectPointOnWallFace(document, wall, normal, location);
					SetWallComponentLocation(wallComponentItem, locationOnWallFace);
					
					FamilyInstance instance = CreateInstanceOnWall(document, wall, symbol, normal, location);
					
					instance.flipFacing();
					document.Regenerate();
					instance.flipFacing();
					document.Regenerate();

					if (false == isFaceOrientationCorrect(instance, normal))
					{
						instance.flipFacing();
					}

					wallComponentItem.RevitID = instance.Id.Value;

					if (instance.Room != null)
						wallComponentItem.RoomName = getRoomNameWithoutNumber(instance.Room);

					wallComponentItem.HostOrientation = getWallOrientation(wallEndPoints);
				}
			}
		}
		
		private FamilySymbol GetFamilySymbolForHole(WallComponent wallComponentItem, Document doc)
		{
			var collector = new FilteredElementCollector(doc);
			FamilySymbol symbol = null;
			
			try
			{
				string h = wallComponentItem.PropertyDictionary["height"].ToString();
				string w = wallComponentItem.PropertyDictionary["width"].ToString();
				double height = Math.Round(double.Parse(h), 2);
				double width = Math.Round(double.Parse(w), 2);
				if (wallComponentItem.Type == WallComponentType.Window)
				{
					string name = "Window" + width + "X" + height;
					collector.OfCategory(BuiltInCategory.OST_Windows).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>();
					symbol = collector.FirstOrDefault(x => x.Name.Equals( name)) as FamilySymbol;
				}
				else if (wallComponentItem.Type == WallComponentType.Door)
				{
					string name = "Door" + width + "X" + height;
					collector.OfCategory(BuiltInCategory.OST_Doors).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>();
					symbol = collector.FirstOrDefault(x => x.Name.Equals( name )) as FamilySymbol;
				}
			}
			catch (InvalidCastException e)
			{
				// Handle the exception if the conversion fails
				TaskDialog.Show("Conversion failed: ",  e.Message);
			}
			catch (FormatException e)
			{
				// Handle the exception if the value is not in a valid format
				TaskDialog.Show("Format is not valid: ", e.Message);
			}
			catch (Exception e)
			{
				// Handle any other exceptions
				TaskDialog.Show("An error occurred: ", e.Message);
			}
			return symbol;
		}
		
		private void SetFamilySymbolDimensions(FamilySymbol symbol, WallComponent wallComponentItem)
		{
			if (symbol == null) return;

			Parameter widthParameter = symbol.LookupParameter("Width");
			Parameter heightParameter = symbol.LookupParameter("Height");

			if (widthParameter != null && !widthParameter.IsReadOnly)
			{
				widthParameter.Set(Convert.ToDouble(wallComponentItem.PropertyDictionary["width"]));
			}

			if (heightParameter != null && !heightParameter.IsReadOnly)
			{
				heightParameter.Set(Convert.ToDouble(wallComponentItem.PropertyDictionary["height"]));
			}
		}

		private JArray getWallOrientation(List<XYZ> wallPoints)
		{
			XYZ minPoint = wallPoints.First();
			XYZ maxPoint = wallPoints.Last();

			XYZ difference = maxPoint - minPoint;
			XYZ normalised = difference.Normalize();

			return new JArray() { normalised.X, normalised.Y, normalised.Z };
		}
		
		public static String getRoomNameWithoutNumber(Autodesk.Revit.DB.Architecture.Room room)
		{
			return room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
		}
		
		private bool isFaceOrientationCorrect(FamilyInstance instance, XYZ normal)
		{
			double tol = 1e-3;

			if (instance.FacingOrientation.Normalize().IsAlmostEqualTo(normal.Normalize(), tol))
				return true;

			return false;
		}
		
		private FamilyInstance CreateInstanceOnWall(Document document, Autodesk.Revit.DB.Wall wall, FamilySymbol symbol, XYZ normal, XYZ location)
		{
			if (IsFaceHosted(symbol))
				return CreateInstanceHostedOnWallFace(document, wall, symbol, normal, location);
			else
				return document.Create.NewFamilyInstance(location, symbol, wall, document.GetElement(wall.LevelId) as Level, 
				                                         			Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
		}

		private FamilyInstance CreateInstanceHostedOnWallFace(Document document, Autodesk.Revit.DB.Wall wall, 
		                                                      		FamilySymbol symbol, XYZ normal, XYZ location)
		{
			XYZ refDir = normal.CrossProduct(XYZ.BasisZ).Multiply(-1);
			refDir.Normalize();
			XYZ locationOnWallFace = ProjectPointOnWallFace(document, wall, normal, location);
			var face = getWallFaceInNormalDirection(document, wall, normal);
			FamilyInstance instance = document.Create.NewFamilyInstance(face, locationOnWallFace, refDir, symbol);
			return instance;
		}

		private static bool IsFaceHosted(FamilySymbol symbol)
		{
			Family fam = symbol.Family;
			bool isFaceHosted;
			var hostParam = fam.GetOrderedParameters().First(p => p.Definition.Name == "Host");
			isFaceHosted = hostParam.AsValueString() == "Face";
			return isFaceHosted;
		}
		
		private static void SetWallComponentLocation(WallComponent wallComponentItem, XYZ location)
		{
			wallComponentItem.X = location.X;
			wallComponentItem.Y = location.Y;
			wallComponentItem.Z = location.Z;
		}
		
		private XYZ calculateFaceNormal(Face face)
		{
			if (face is PlanarFace)
				return (face as PlanarFace).FaceNormal;

			var bboxUV = face.GetBoundingBox();
			var center = (bboxUV.Max + bboxUV.Min) / 2.0;
			return face.ComputeNormal(center).Normalize();
		}
		
		private Reference getWallFaceInNormalDirection(Document document, Autodesk.Revit.DB.Wall wall, XYZ normal)
		{
			IList<Reference> interiorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior);
			Face interiorFace = document.GetElement(interiorFaces[0]).GetGeometryObjectFromReference(interiorFaces[0]) as Face;

			IList<Reference> exteriorFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);
			Face exteriorFace = document.GetElement(exteriorFaces[0]).GetGeometryObjectFromReference(exteriorFaces[0]) as Face;

			XYZ wallNormalExterierFace = calculateFaceNormal(exteriorFace).Normalize();
			double tol = 1e-6;

			if (wallNormalExterierFace.IsAlmostEqualTo(normal, tol))
				return exteriorFaces[0];

			return interiorFaces[0];
		}
		
		private XYZ projectPointOnHostFace(Face hostFace, XYZ point)
		{
			Surface surface = hostFace.GetSurface();
			UV uvProjectPoint = null;
			double distance = 0.0;
			surface.Project(point, out uvProjectPoint, out distance);

			return hostFace.Evaluate(uvProjectPoint);
		}
		
		private XYZ ProjectPointOnWallFace(Document document, Autodesk.Revit.DB.Wall wall, XYZ normal, XYZ point)
		{
			var face = getWallFaceInNormalDirection(document, wall, normal);
			//IF we still don't find a face with correct normal from then throw exception
			if (face == null)
				new Exception("Wall face was null for wall ID: " + wall.Id.ToString());

			Face wallFace = document.GetElement(face).GetGeometryObjectFromReference(face) as Face;
			XYZ locationOnWallFace = this.projectPointOnHostFace(wallFace, point);
			return locationOnWallFace;
		}
		
		private XYZ calculateElementLocation(List<XYZ> wallEndPoints, WallComponent wallComponentItem)
		{
			double unitConversionFactor = UnitUtils.ConvertToInternalUnits(1, UnitTypeId.Meters);
			double offset = wallComponentItem.Offset;
			XYZ offsetDir = wallEndPoints[1] - wallEndPoints[0];
			offset = offsetDir.GetLength() * offset;
			offsetDir = offsetDir.Normalize();
			double zValue = 0.0;
			if (wallComponentItem.PropertyDictionary.ContainsKey("altitude"))
			{
				zValue = double.Parse(wallComponentItem.PropertyDictionary["altitude"].ToString());
			}
			XYZ location = wallEndPoints[0] + (offsetDir * offset) + new XYZ(0, 0, zValue);

			return location;
		}
		
		private List<XYZ> getWallEndPointsInFeet(WallComponent wallComponentItem)
		{
			double unitConversionFactor = UnitUtils.ConvertToInternalUnits(1, UnitTypeId.Meters);

			XYZ p1 = new XYZ(
				wallComponentItem.Wall.VertextIDList[0].X, 
				wallComponentItem.Wall.VertextIDList[0].Y, 
				wallComponentItem.Wall.VertextIDList[0].Z);
			
			XYZ p2 = new XYZ(
				wallComponentItem.Wall.VertextIDList[1].X,
				wallComponentItem.Wall.VertextIDList[1].Y,
				wallComponentItem.Wall.VertextIDList[1].Z);
			
			List<XYZ> pointList = new List<XYZ>(){p1,p2};
			pointList = OrderWallVertices(pointList);
			p1 = pointList.First();
			p2 = pointList.Last();
			return new List<XYZ> {p1, p2};
		}
		
		private List<XYZ> OrderWallVertices(List<XYZ> list)
		{
			XYZ p1 = list.First();
			XYZ p2 = list.Last();
			double tol = 1e-5;
			if (Math.Abs(p1.X - p2.X) < tol)
				return list.OrderBy(point => point.Y).ToList<XYZ>();

			return list.OrderBy(point => point.X).ToList<XYZ>();
		}
		
		private Dictionary<long, Autodesk.Revit.DB.Wall> createWallIDDictionary(Document document)
		{
			FilteredElementCollector collector = new FilteredElementCollector(document);
			ICollection<Autodesk.Revit.DB.Element> collection = collector.OfClass(typeof(Autodesk.Revit.DB.Wall)).ToElements();

			Dictionary<long, Autodesk.Revit.DB.Wall> wallMap = new Dictionary<long, Autodesk.Revit.DB.Wall>();
			foreach (Autodesk.Revit.DB.Element wall in collection)
			{
				wallMap.Add(wall.Id.Value, wall as Autodesk.Revit.DB.Wall);
			}
			return wallMap;
		}
	}
}