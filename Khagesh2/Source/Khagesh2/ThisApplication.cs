/*
 * Created by SharpDevelop.
 * User: Ashish Kamble
 * Date: 27-05-2024
 * Time: 11:44
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;

namespace Khagesh2
{
	public class CustomFamilyLoadOption : IFamilyLoadOptions
	{
		public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
		{
			overwriteParameterValues = true;
			return true;
		}

		public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
		{
			source = FamilySource.Family;
			overwriteParameterValues = true;
			return true;
		}
	}
	
	public enum WallComponentType
	{
		Door,
		Window
	}
	
	public enum WallType
	{
		Glass_Wall,
		Partition_Wall,
		Exposed_Wall
	}
	
	public class Element
	{
		public string Name { get; internal set; }
		public String ID { get; set; }
		public long RevitID { get; set; }
		public String RoomName { get; set; }
		//Store in feet Unit as Revit uses data in feet unit
		public double X { get; set; }
		public double Y { get; set; }
		public double Z { get; set; }
		public Dictionary<String, Object> PropertyDictionary { get; set; }
		public double mRotation { get; set; }
		public JArray HostOrientation { get; set; }
	}
	
	public class Building : Element
	{
		public List<Wall> WallList { get; set; }
		public List<Room> RoomList { get; set; }
		public List<Vertex> VertexList { get; set; }
		public List<Component> ComponentList { get; set; }
		public List<WallComponent> WallComponent { get; set; }
		public List<List<UV>> SlabVertexList { get; set; }
		
	}
	
	public class Component : Element
	{
		public Component(double x, double y, double z, double mRotation, string name, string iD, Dictionary<String, Object> propertyDictionary)
		{
			X = x;
			Y = y;
			Z = z;
			this.mRotation = mRotation;
			Name = name;
			ID = iD;
			PropertyDictionary = propertyDictionary;
		}

		public double NormalX { get; set; }
		public double NormalY { get; set; }
		public double NormalZ { get; set; }

	}
	
	public class Room : Element
	{
		public Room(string name, string iD, List<Vertex> vertextList, double[] fluidPoint, Dictionary<String, Object> propertyDictionary)
		{
			Name = name;
			ID = iD;
			VertextList = vertextList;
			PropertyDictionary = propertyDictionary;
			FluidPoint = fluidPoint;
		}

		public List<Vertex> VertextList { get; set; }

		public double[] FluidPoint { get; set; }
	}
	
	public class Wall : Element
	{
		public Wall(WallType type, string name, string iD, List<Vertex> vertextIDList, double height, double baseOffset, Arc wallArc)
		{
			Type = type;
			Name = name;
			ID = iD;
			VertextIDList = vertextIDList;
			WallHeight = height;
			WallBaseOffset = baseOffset;
			WallArc = wallArc;
		}
		public WallType Type { get; set; }
		public List<WallComponent> ItemsOnWallIDList { get; set; }
		public List<Vertex> VertextIDList { get; set; }
		public double WallHeight {get; set;}
		public double WallBaseOffset { get; set; }
		public Arc WallArc {get; set;}
	}
	
	public class WallComponent : Element
	{
		public WallComponentType Type { get; set; }
		public WallComponent(WallComponentType type, string name, string iD, Dictionary<String, Object> propertyDictionary, double offset,double normalX, double normalY, double normalZ)
		{
			Type = type;
			Name = name;
			ID = iD;
			PropertyDictionary = propertyDictionary;
			//WallID = wallID;
			Offset = offset;
			NormalX = normalX;
			NormalY = normalY;
			NormalZ = normalZ;
			
		}
		
		[JsonIgnore]
		public Wall Wall { get; set; }
		
		public double Offset { get; set; }
		public double NormalX { get; set; }
		public double NormalY { get; set; }
		public double NormalZ { get; set; }
	}
	
	public class Vertex : Element
	{
		public Vertex(double x, double y, double z, string name, string iD)
		{
			X = x;
			Y = y;
			Z = z;
			Name = name;
			ID = iD;
		}

		public List<Wall> WallList { get; set; }
	}
	
	[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
	[Autodesk.Revit.DB.Macros.AddInId("78566F64-665A-4791-AA00-B008F4B08C50")]
	public partial class ThisApplication
	{
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
		private static double calculateUnitConversionFactor(string inputUnit)
		{
			//Convert to Meter Unit as we will be storing data in meters
			// while creating geometry in Revit we will convert X,Y,Z to Feet Unit as Revit internal Unit is feet

			switch (inputUnit)
			{
				case "cm":
					return UnitUtils.Convert(1, UnitTypeId.Centimeters, UnitTypeId.Meters);
				case "m":
					return UnitUtils.Convert(1, UnitTypeId.Meters, UnitTypeId.Meters);
				case "mm":
					return UnitUtils.Convert(1, UnitTypeId.Millimeters, UnitTypeId.Meters);
				case "in":
					return UnitUtils.Convert(1, UnitTypeId.Inches, UnitTypeId.Meters);
				case "ft":
					return UnitUtils.Convert(1, UnitTypeId.Feet, UnitTypeId.Meters);
				default:
					throw new Exception("Not Supported Unit");
			}
		}
		
		public void ReadBuildingJson()
		{
			Document doc = this.ActiveUIDocument.Document;

			string fileName = "E:\\output\\Dataset014.json";
			StreamReader streamReader = new StreamReader(fileName);
			string json = streamReader.ReadToEnd();
			streamReader.Close();
			
			LoadAllFamilies(doc);
			
			JObject builderRoot = JsonConvert.DeserializeObject<JObject>(json);
			
			int buildingCount = builderRoot["site"]["Buildings"].Count();
			
			try{
				
				for(int i = 1; i <= buildingCount; ++i)
				{
					Building building = new Building();
					string buildingName = "" + i;
					
					string unit = builderRoot["site"]["Buildings"][buildingName]["unit"].ToString();
					//double unitConversionFactor = calculateUnitConversionFactor(unit);
					int levelsCount = builderRoot["site"]["Buildings"][buildingName]["Levels"].Count();
					
					for(int j = 1; j <= levelsCount; ++j)
					{
						string levelName = "" + j;
						
//					double altitude = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["altitude"].ToObject<double>();
						double elevation = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["elevation"].ToObject<double>();
						//                	double wallBaseOffset = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["lines"].ToObject<double>();
						
						double levelAltitude = elevation;
						//                	if(elevation != 0)
						//                	{
						//                		levelAltitude = altitude + elevation;
						//                	}
						
						var vertices = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["vertices"].ToObject<JObject>();
						var lines = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["lines"].ToObject<JObject>();
						var areas = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["areas"].ToObject<JObject>();
						var holes = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["holes"].ToObject<JObject>();
						var slabs = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["slabs"].ToObject<JObject>();
						var structuralMembers = builderRoot["site"]["Buildings"][buildingName]["Levels"][levelName]["structuralMembers"].ToObject<JObject>();
						
						using (Transaction trans = new Transaction(doc)) {
							trans.Start("Column adding");
							
							List<XYZ> startPoints = new List<XYZ>();
							try {
								foreach (var item in structuralMembers)
								{
									var data = item.Value;
									
									double startPoint_x = data["StartPoint"]["X"].ToObject<double>();
									double startPoint_y = data["StartPoint"]["Y"].ToObject<double>();
									double startPoint_z = data["StartPoint"]["Z"].ToObject<double>();
									
									double endPoint_z = data["EndPoint"]["Z"].ToObject<double>();
									
									XYZ startPoint = new XYZ(startPoint_x,startPoint_y, startPoint_z);
									double columnHeight = endPoint_z - startPoint_z;
									
									
									FilteredElementCollector collector = new FilteredElementCollector(doc);
									collector.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralColumns);
									FamilySymbol columnType = collector.FirstElement() as FamilySymbol;
									ElementId levelId = Level.GetNearestLevelId(doc, elevation);
									
									Autodesk.Revit.DB.Element element = doc.GetElement(levelId);
									Level level = element as Level;
									FamilyInstance instance = null;
									if (null != columnType)
									{
										// Create a column at the origin
										//							XYZ origin = new XYZ(0, 0, 0);
										
										instance = doc.Create.NewFamilyInstance(startPoint, columnType, level, StructuralType.Column);
										
									}
									
									
								}
							} catch (Exception ex) {
								
								TaskDialog.Show("Column Exception", ex.Message);
							}
							
							trans.Commit();
							
						}
						
						
						
						Dictionary<String, Vertex> vertexMap = new Dictionary<String, Vertex>();
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
						
						Dictionary<String, WallComponent> wallComponentMap = new Dictionary<String, WallComponent>();
						foreach (var item in holes)
						{
							var data = item.Value.ToObject<JObject>();
							string id = item.Key;
							
							List<string> tags = data["tag"].Select(tag => tag.ToString().ToLower()).ToList<string>();
							if (tags.Contains("airsidesystem")) { continue; }
							
							string name = data["type"].ToString();
							string type = data["type"].ToString();
							Double offset = data["offset"].ToObject<Double>(); // This value is Unit less so no need to apply unit conversion factor
							double startPoint_x = Double.Parse(data["startPoint"]["Position"]["X"].ToString());
							double startPoint_y = (double)data["startPoint"]["Position"]["Y"];
							double startPoint_z = (double)data["startPoint"]["Position"]["Z"];
							
							XYZ startPoint = new XYZ(startPoint_x,startPoint_y, startPoint_z);
							
							double endPoint_x = data["startPoint"]["Position"]["X"].ToObject<double>();
							double endPoint_y = data["startPoint"]["Position"]["Y"].ToObject<double>();
							double endPoint_z = data["startPoint"]["Position"]["Z"].ToObject<double>();
							
							XYZ endPoint = new XYZ(endPoint_x,endPoint_y,endPoint_z);
							
							//midpoint is the point where the door actually located
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
							
							switch (type)
							{
								case "door":
									componentType = WallComponentType.Door;
									break;
								case "window":
									componentType = WallComponentType.Window;
									break;
							}
							
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
						
						List<Wall> wallList = new List<Wall>();
						try{
							foreach (var item in lines)
							{
								var data = item.Value;
								string id = item.Key;
								
								string name = data["name"].ToString();
								if(data["vertices"] == null)
								{
									TaskDialog.Show("Vertices Id", id);
									continue;
								}

//								if(data["vertices"] != null)
//								{
//									var vertexIDList = data["vertices"] as JArray;
//									
//									if (vertexIDList != null && vertexIDList.Count == 2)
//									{
//										//continue;
//										//throw new Exception("Line vertices ID array count is not equal to 2");
//										
//										string vertexID1 = vertexIDList[0].ToObject<String>();
//										string vertexID2 = vertexIDList[1].ToObject<String>();
//										
//										Vertex v1 = vertexMap[vertexID1];
//										Vertex v2 = vertexMap[vertexID2];
//										
//										List<Vertex> vertexList = new List<Vertex>();
//										vertexList.Add(v1);
//										vertexList.Add(v2);
//										
//										WallType wallType = WallType.Exposed_Wall;
//										Enum.TryParse(data["properties"]["type"].ToString(), out wallType);
//										double WallHeight = data["properties"]["Height"].ToObject<double>();
//										double WallBaseOffset = data["properties"]["BaseOffset"].ToObject<double>();
//										Wall wall = new Wall(wallType, name, id, vertexList, WallHeight, WallBaseOffset);
//										wall.PropertyDictionary = data["properties"].ToObject<Dictionary<string, Object>>();
//										
//										v1.WallList.Add(wall);
//										v2.WallList.Add(wall);
//										
//										List<WallComponent> wallComponentList = new List<WallComponent>();
//										if(data["holes"].Count() != 0)
//										{
//											var wallComponentIDList = data["holes"] as JArray;
//											foreach (var wallComponentIDItem in wallComponentIDList)
//											{
//												string wallComponentID = wallComponentIDItem.ToObject<String>();
//												
//												if (wallComponentMap.ContainsKey(wallComponentID))
//												{
//													// It's Component on Wall such as Door or Window
//													WallComponent wallComponent = wallComponentMap[wallComponentID];
//													wallComponent.Wall = wall;
//													wallComponentList.Add(wallComponent);
//												}
//											}
//											
//											//Add Wall Diffuser List
//											wall.ItemsOnWallIDList = wallComponentList;
//										}
//										
//										wallList.Add(wall);
//									}
//									else
//									{
//										continue;
//									}
//								}								
								
								
								
								Wall wall;
								
								WallType wallType = WallType.Exposed_Wall;
								Enum.TryParse(data["properties"]["type"].ToString(), out wallType);
								double WallHeight = data["properties"]["Height"].ToObject<double>();
								double WallBaseOffset = data["properties"]["BaseOffset"].ToObject<double>();
								
								var vertexIDList = data["vertices"] as JArray;
								string vertexID1 = vertexIDList[0].ToObject<String>();
								string vertexID2 = vertexIDList[1].ToObject<String>();
								
								Vertex v1 = vertexMap[vertexID1];
								Vertex v2 = vertexMap[vertexID2];
								
								List<Vertex> vertexList = new List<Vertex>();
								vertexList.Add(v1);
								vertexList.Add(v2);
								
																
								if(data["curatainArcWall"] != null && data["curatainArcWall"].HasValues == true)
								{
									var curtainArcWallPoints = data["curatainArcWall"];
									XYZ center = new XYZ((double)curtainArcWallPoints["Center"]["X"], (double)curtainArcWallPoints["Center"]["Y"], (double)curtainArcWallPoints["Center"]["Z"]);
									double startAngle = (double)curtainArcWallPoints["StartAngle"];
									double endAngle = (double)curtainArcWallPoints["EndAngle"];
									double radius = (double)curtainArcWallPoints["Radius"];
									XYZ xAxis = new XYZ( (double)curtainArcWallPoints["Xaxis"]["X"], (double)curtainArcWallPoints["Xaxis"]["Y"], (double)curtainArcWallPoints["Xaxis"]["Z"]);
									XYZ yAxis = new XYZ( (double)curtainArcWallPoints["Yaxis"]["X"], (double)curtainArcWallPoints["Yaxis"]["Y"], (double)curtainArcWallPoints["Yaxis"]["Z"]);
									
									Arc arcWall = Arc.Create(center, radius, startAngle, endAngle, xAxis, yAxis);
									wall = new Wall(wallType, name, id, vertexList, WallHeight, WallBaseOffset, arcWall);
									
									wall.PropertyDictionary = data["properties"].ToObject<Dictionary<string, Object>>();
									v1.WallList.Add(wall);
								v2.WallList.Add(wall);
								}
								else if(data["ArcWall"] != null && data["ArcWall"].HasValues == true)
								{
									var standardArcWallPoints = data["ArcWall"];
									XYZ startPoint = new XYZ((double)standardArcWallPoints["StartPoint"]["X"], (double)standardArcWallPoints["StartPoint"]["Y"], (double)standardArcWallPoints["StartPoint"]["Z"]);
									XYZ pointOnArc = new XYZ((double)standardArcWallPoints["PointOnArc"]["X"], (double)standardArcWallPoints["PointOnArc"]["Y"], (double)standardArcWallPoints["PointOnArc"]["Z"]);
									XYZ endPoint = new XYZ((double)standardArcWallPoints["EndPoint"]["X"], (double)standardArcWallPoints["EndPoint"]["Y"], (double)standardArcWallPoints["EndPoint"]["Z"]);
									
									Arc arcWall = Arc.Create(startPoint, endPoint, pointOnArc);
									
									
									
									wall = new Wall(wallType, name, id, vertexList, WallHeight, WallBaseOffset, arcWall);	
									wall.PropertyDictionary = data["properties"].ToObject<Dictionary<string, Object>>();
									v1.WallList.Add(wall);
								v2.WallList.Add(wall);
								}
								else{
									wall = new Wall(wallType, name, id, vertexList, WallHeight, WallBaseOffset, null);
									wall.PropertyDictionary = data["properties"].ToObject<Dictionary<string, Object>>();
									v1.WallList.Add(wall);
								v2.WallList.Add(wall);
								}
																		
//								wall = new Wall(wallType, name, id, vertexList, WallHeight, WallBaseOffset, null);
//								wall.PropertyDictionary = data["properties"].ToObject<Dictionary<string, Object>>();
//								v1.WallList.Add(wall);
//								v2.WallList.Add(wall);
								
								
								
								
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
						
						List<Room> roomList = new List<Room>();
						List<XYZ> vertexList1 = new List<XYZ>();
						using (Transaction trans = new Transaction(doc)) {
							
							trans.Start("Start");
							
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
						
						building.SlabVertexList = new List<List<UV>>();
						List<UV> slabVertexList = new List<UV>();
						List<XYZ> slabHolePointsList = new List<XYZ>();
						
						FloorType floorType = createFloorType(doc);
						CeilingType ceilingType = createCeilingType(doc);
						
						ElementId levelId1 = Level.GetNearestLevelId(doc, levelAltitude);
						
						foreach (var item in slabs)
						{
							
							var data = item.Value;
							string id = item.Key;
							var vertexList = data["SlabLoop"] as JArray;
							var floorHoles = data["Holes"] as JArray;
							var lowPoint = 	data["LowPoint"].ToObject<double>();
							var highPoint = data["HighPoint"].ToObject<double>();
							
							foreach (var v in vertexList)
							{
								double x = v[0].ToObject<double>();
								double y = v[1].ToObject<double>();
								
								UV slabPoint = new UV(x, y);
								slabVertexList.Add(slabPoint);
//		                	building.SlabVertexList(slabPoint);
							}
							//						building.SlabVertexList = slabVertexList;
							//		                building.SlabVertexList.Append(slabVertexList);
							//						slabVertexList.Clear();
							
							CurveLoop profile = new CurveLoop();
							CurveArray holePoints = new CurveArray();
							Floor floor = null;
							
							try  {
								using (Transaction trans = new Transaction(doc))
								{
									trans.Start("started");
									
									
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
									
									if (highPoint == 0.0 || lowPoint == 0.0) {
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
							} catch (Exception ex) {
								TaskDialog.Show("Floor Exception", ex.Message);
							}
							
							
							using (Transaction trans = new Transaction(doc))
							{
								trans.Start("started");
								FilteredElementCollector collector = new FilteredElementCollector(doc);
								collector.OfCategory(BuiltInCategory.OST_Levels);
								collector.OfClass(typeof(Level));
//							Level level = collector.ToElements().Cast<Level>().ToList<Level>().First();
								
								RoofType roofType = new FilteredElementCollector(doc).OfClass(typeof(RoofType)).FirstOrDefault() as RoofType;
								ElementId levelId = Level.GetNearestLevelId(doc, elevation);
								
								Autodesk.Revit.DB.Element element = doc.GetElement(levelId);
								Level level = element as Level;
								
								
//								ModelCurveArray footPrintToModelCurveMapping = new ModelCurveArray();
//								var footprintRoof = doc.Create.NewFootPrintRoof(holePoints, level, roofType, out footPrintToModelCurveMapping);
//								RoofBase roof = holePoints as RoofBase;
								
								
								
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
											foreach (var line in openingLines)
												openingProfile.Append(line);

//											var op2 = doc.Create.NewOpening(ceiling, openingProfile, true);
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
						
						building.VertexList = vertexMap.Values.ToList();
						building.WallComponent = wallComponentMap.Values.ToList();
						building.WallList = wallList;
						building.RoomList = roomList;
						
						
						CreateWindowSymbols(doc, building);
						
						createBuildingRevitFile(building, doc, levelAltitude, elevation, levelsCount, j);
					}
				}
			}
			catch(Exception ex){
				TaskDialog.Show("Exception", ex.StackTrace);
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
			createBuilding(building, document, ref centroidRoomNameMap, levelAltitude, altitude, levelsCount, j);
			createComponentsInBuilding(building, document, centroidRoomNameMap);
		}
		
		private void createComponentsInBuilding(Building building, Document document, Dictionary<XYZ, string> centroidRoomNameMap)
		{
			Transaction componentPlacementTransaction = new Transaction(document);// no need to to use "Using", Autodesk disposes of all the methods and data in Execute method.
			{
				componentPlacementTransaction.Start("create componets in building spaces");

				correctWallOrientation(document);

				assignRoomName(document, centroidRoomNameMap);

				// Component (Seating layout) place
				//placeComponent(document, building.ComponentList);
				
				try{
					PlaceWallComponents(document, building.WallComponent);
				}
				catch(Exception ex)
				{
					Console.WriteLine("Exception occurs at createComponentsInBuilding: " + ex.Message);
				}

				// Wall Component (Door, Window) place
				

				componentPlacementTransaction.Commit();
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
		
		private void correctWallOrientation(Document document)
		{
			FilteredElementCollector collector = new FilteredElementCollector(document);
			ICollection<Autodesk.Revit.DB.Element> collection = collector.OfClass(typeof(Autodesk.Revit.DB.Wall)).ToElements();

			foreach (Autodesk.Revit.DB.Element e in collection)
			{
				Autodesk.Revit.DB.Wall wall = e as Autodesk.Revit.DB.Wall;
				
				double wallThickness = 0.7916;

				IList<Reference> sideFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);

				if (sideFaces.Count == 0)
					continue;
				Face internalFace = document.GetElement(sideFaces[0]).GetGeometryObjectFromReference(sideFaces[0]) as Face;
				BoundingBoxUV bboxUV = internalFace.GetBoundingBox();
				UV center = (bboxUV.Max + bboxUV.Min) / 2.0;
				XYZ normal = internalFace.ComputeNormal(center).Normalize();
				XYZ centerPoint = internalFace.Evaluate(center);

				XYZ pointAtNormalDirection = centerPoint + normal * wallThickness * 2.0;

				Autodesk.Revit.DB.Architecture.Room otherRoom = document.GetRoomAtPoint(pointAtNormalDirection);

				if (otherRoom != null)
				{
					//IF we didnot get other room at point At Face Normal direction, then we should flip wall
					wall.Flip();
				}
			}
		}
		
		private void createBuilding(Building building, Document document, ref Dictionary<XYZ, string> centroidRoomNameMap, double levelAltitude, double altitude, int levelsCount, int j)
		{
			// Get a floor type for floor creation
			FloorType floorType = createFloorType(document);
			CeilingType ceilingType = createCeilingType(document);

			XYZ normal = XYZ.BasisZ;

			Transaction transaction = new Transaction(document);
			{
				transaction.Start("create building spaces");

				document.Regenerate();
				
				Dictionary<WallType, ElementId> wallTypeElemIdMap = GetWallTypeElementIDMap(document);
				CreateWalls(document, building.WallList, levelAltitude,altitude, wallTypeElemIdMap);
				try
				{
					createFloorsCeilings(building, document, levelAltitude, floorType, ceilingType, levelsCount, normal, j);
				}
				catch(Exception ex)
				{
					TaskDialog.Show("Error", ex.Message);
					TaskDialog.Show("Error", ex.StackTrace);
				}
				
				
				

				//createRooms(document, level_id1, ceilingHeight);

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
				WallType.Partition_Wall,
				WallType.Glass_Wall
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
		
		private void CreateWalls(Document doc,List<Wall> wallList, double levelAltitude, double altitude, Dictionary<WallType, ElementId> wallTypeIdMap)
		{
			foreach (var wallItem in wallList)
			{
//				XYZ p1 = new XYZ(wallItem.VertextIDList[0].X, wallItem.VertextIDList[0].Y, wallItem.VertextIDList[0].Z);
//				XYZ p2 = new XYZ(wallItem.VertextIDList[1].X, wallItem.VertextIDList[1].Y, wallItem.VertextIDList[1].Z);
//
//				//Convert to Revit Internal Unit
//				//p1 *= unitConversionFactor;
//				//p2 *= unitConversionFactor;
//
//				Curve curve = Line.CreateBound(p1, p2);
//
//				Autodesk.Revit.DB.Wall wall = Autodesk.Revit.DB.Wall.Create(doc, Line.CreateBound(p1, p2), wallTypeIdMap[wallItem.Type], Level.GetNearestLevelId(doc, levelAltitude), wallItem.WallHeight, wallItem.WallBaseOffset, true, false);
//				
//				wallItem.RevitID = wall.Id.Value;
//				//Parameter parm = wall.LookupParameter("Top Constraint");
//				//parm.Set(level_id2);
				
				Autodesk.Revit.DB.Wall wall;
				if(wallItem.WallArc == null)
				{
					XYZ p1 = new XYZ(wallItem.VertextIDList[0].X, wallItem.VertextIDList[0].Y, wallItem.VertextIDList[0].Z);
					XYZ p2 = new XYZ(wallItem.VertextIDList[1].X, wallItem.VertextIDList[1].Y, wallItem.VertextIDList[1].Z);
	
					Curve curve = Line.CreateBound(p1, p2);
					wall = Autodesk.Revit.DB.Wall.Create(doc, curve, wallTypeIdMap[wallItem.Type], Level.GetNearestLevelId(doc, levelAltitude), wallItem.WallHeight, wallItem.WallBaseOffset, true, false);
				}
				else
				{
					wall = Autodesk.Revit.DB.Wall.Create(doc, wallItem.WallArc, wallTypeIdMap[wallItem.Type], Level.GetNearestLevelId(doc, levelAltitude), wallItem.WallHeight, wallItem.WallBaseOffset, true, false);
				}
				wallItem.RevitID = wall.Id.Value;
				
			}
		}
		
		private void LoadAllFamilies(Document document)
		{
			using (Transaction trans = new Transaction(document, "Load Families"))
			{
				trans.Start();
				mDocumentFamilySymbols.Clear();

				List<FamilySymbol> familySymbolsInDoc = new FilteredElementCollector(document)
					.WherePasses(new ElementClassFilter(typeof(FamilySymbol)))
					.Cast<FamilySymbol>().ToList();
				
//	            List<FamilySymbol> familySymbolsInDoc = new FilteredElementCollector(document).OfClass(typeof(FamilySymbol))
//	                    .Cast<FamilySymbol>().ToList();
				
				int count = familySymbolsInDoc.Count();
				
				foreach (var familySymbol in familySymbolsInDoc)
				{
					string keyName = familySymbol.Family.Name + '_' + familySymbol.Name;
					
					mDocumentFamilySymbols.Add(keyName.ToLower(), familySymbol);
				}
				
				trans.Commit();
			}
			
		}
		
		private void CreateWindowSymbols(Document document, Building building)
		{
			try
			{
				List<WallComponent> wallComponentList = building.WallComponent;
				KeyValuePair<string, FamilySymbol> windowDict = (from symbol in mDocumentFamilySymbols where symbol.Value.Family.Name.ToLower() == "window" select symbol).FirstOrDefault();
				if (String.IsNullOrEmpty(windowDict.Key))
				{
					return;
				}
				
				FamilySymbol windowSymbol = windowDict.Value;
				
				foreach (WallComponent component in wallComponentList)
				{
					String familyName = component.Name.ToLower();
					String symbolName = String.Empty;
					
					if (familyName != "window")
					{
						continue;
					}
					
					FamilySymbol newWindowSymbol = CreateNewWindowType(document, component, windowSymbol, familyName, out symbolName);
					
					Transaction activateTransaction = new Transaction(document);
					activateTransaction.Start("Activate Window Symbol");
					if (newWindowSymbol.IsActive == false)
					{
						newWindowSymbol.Activate();
					}
					activateTransaction.Commit();
					
					String keyname = familyName + '_' + symbolName;
					
					if (mDocumentFamilySymbols.ContainsKey(keyname))
					{
						continue;
					}
					
//					mDocumentFamilySymbols.Add(keyname, newWindowSymbol);
				}
				
				mDocumentFamilySymbols.Remove(windowDict.Key);
				
//	            Transaction deleteOldSymbol = new Transaction(document);
//	            deleteOldSymbol.Start("Deleting Old Window Symbol");
//	            document.Delete(windowSymbol.Id);
//	            deleteOldSymbol.Commit();
			}
			catch(Exception ex)
			{
				TaskDialog.Show("Window error", ex.Message);
			}
			
		}
		
		private FamilySymbol CreateNewWindowType(Document document, WallComponent component, FamilySymbol symbol, String familyName, out String newSymbolName)
		{
			Family family = symbol.Family;
			newSymbolName = GetWindowSymbolFromComponent(component, '_');
			String componentSymbolName = familyName + '_' + newSymbolName;
			if (mDocumentFamilySymbols.ContainsKey(componentSymbolName))
			{
				return mDocumentFamilySymbols[componentSymbolName];
			}

			string[] paramValue = newSymbolName.Split('_');
			Dictionary<string, double> parameters = new Dictionary<string, double>()
			{
				{ "Width", Convert.ToDouble(paramValue[0]) },
				{ "Height", Convert.ToDouble(paramValue[1]) },
			};

			return CreateNewSymbol(document, parameters, family, newSymbolName);
		}
		
		private FamilySymbol CreateNewSymbol(Document document, Dictionary<string, double> parameters, Family family, String newFamilyTypeName)
		{
			if (null == family)
			{
				return null;    // could not get the family
			}

			// Get Family document for family
			Document familyDoc = document.EditFamily(family);
			if (null == familyDoc)
			{
				return null;    // could not open a family for edit
			}

			FamilyManager familyManager = familyDoc.FamilyManager;
			if (null == familyManager)
			{
				return null;  // cuould not get a family manager
			}

			// Start transaction for the family document
			using (Transaction newFamilyTypeTransaction = new Transaction(familyDoc, "Add Type to Family"))
			{
				newFamilyTypeTransaction.Start();

				// add a new type and edit its parameters
				FamilyType newFamilyType = familyManager.NewType(newFamilyTypeName);
				if (newFamilyType == null)
				{
					throw new Exception("Unable to create new type for " + newFamilyTypeName);
				}

				foreach (KeyValuePair<string, double> property in parameters)
				{
					FamilyParameter familyParameter = familyManager.get_Parameter(property.Key);
					if (familyParameter == null)
					{
						throw new Exception("Unable to find Width Param for Window");
					}
					familyManager.Set(familyParameter, property.Value);
				}

				newFamilyTypeTransaction.Commit();
			}

			FamilySymbol updatedSymbol = null;

			family = familyDoc.LoadFamily(document, new CustomFamilyLoadOption());
			if (null == family)
			{
				return updatedSymbol;
			}

			// find the new type
			ISet<ElementId> familySymbolIds = family.GetFamilySymbolIds();
			foreach (ElementId id in familySymbolIds)
			{
				FamilySymbol familySymbol = family.Document.GetElement(id) as FamilySymbol;
				if (familySymbol == null || familySymbol.Name != newFamilyTypeName)
				{
					continue;
				}

				using (Transaction changeSymbol = new Transaction(document, "Change Symbol Assignment"))
				{
					changeSymbol.Start();
					updatedSymbol = familySymbol;
					changeSymbol.Commit();
				}
				break;
			}

			return updatedSymbol;
		}
		
		private void PlaceWallComponents(Document document, List<WallComponent> wallComponent)
		{
			Dictionary<long, Autodesk.Revit.DB.Wall> wallMap = createWallIDDictionary(document);
			foreach (WallComponent wallComponentItem in wallComponent)
			{
				long wallID = wallComponentItem.Wall.RevitID;

				Autodesk.Revit.DB.Wall wall = wallMap[wallID];

				String familyName = String.Empty;
				String symbolName = String.Empty;
				GetComponentFamilyData(JObject.FromObject(wallComponentItem), out familyName, out symbolName);

				if (familyName == "window")
				{
					symbolName = GetWindowSymbolFromComponent(wallComponentItem, '_');
				}

				String symbolToGet = familyName + '_' + symbolName;

				if (mDocumentFamilySymbols.ContainsKey(symbolToGet) == false)
				{
					//                	TaskDialog.Show("Unable to Find Family for " , familyName);
					
					continue;
				}

				FamilySymbol symbol = mDocumentFamilySymbols[symbolToGet];
//				FilteredElementCollector collectorWindows = new FilteredElementCollector(document);
//				FamilySymbol symbol = collectorWindows.OfCategory(BuiltInCategory.OST_Doors).OfClass(typeof(FamilySymbol))
//														.Cast<FamilySymbol>().FirstOrDefault();
				
				//                Parameter widthParameter = symbol.LookupParameter("Width");
//			    Parameter heightParameter = symbol.LookupParameter("Height");
//
				
//			    if (widthParameter != null && !widthParameter.IsReadOnly)
//			    {
				////			        widthParameter.Set();
//
//
//			    	foreach(KeyValuePair<string, object> prop in wallComponentItem.PropertyDictionary)
//			    	{
//			    		if(prop.Key == "width")
//			    		{
//			    			double num = Convert.ToDouble(prop.Value);
//			    			widthParameter.Set(num);
				////			    			break;
				////			    			var num = (double)prop.Value;
//			    		}
//			    		break;
//			    	}
//			    }
				
//			    if (heightParameter != null && !heightParameter.IsReadOnly)
//			    {
//			        heightParameter.Set(properties.Height);
//			    }

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
				return document.Create.NewFamilyInstance(location, symbol, wall, document.GetElement(wall.LevelId) as Level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
		}

		private FamilyInstance CreateInstanceHostedOnWallFace(Document document, Autodesk.Revit.DB.Wall wall, FamilySymbol symbol, XYZ normal, XYZ location)
		{
			// Need to multiply refDir by -1 to correct orientation. Not sure why
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
			
			/*
            //Another way to get the face from wall if Wall in created in Ongoing transaction
            Options geomOptions = new Options();
            geomOptions.ComputeReferences = true;
            GeometryElement wallGeom = wall.get_Geometry(geomOptions);
            foreach (GeometryObject geomObj in wallGeom)
            {
                Solid geomSolid = geomObj as Solid;
                if (null == geomSolid)
                    continue;

                foreach (Face geomFace in geomSolid.Faces)
                {
                    XYZ normal1 = calculateFaceNormal(geomFace);

                    if (normal1.IsAlmostEqualTo(normal, 1e-3))
                        return geomFace;
                }
            }

            return null;*/
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
			if (wallComponentItem.Name.ToLower() == "window")
			{
				offset += (Convert.ToDouble(wallComponentItem.PropertyDictionary["width"]) / 2);
			}
			offsetDir = offsetDir.Normalize();

			double zValue = 2.95;
			if (wallComponentItem.PropertyDictionary.ContainsKey("altitude"))
			{
				zValue = double.Parse(wallComponentItem.PropertyDictionary["altitude"].ToString());
				//zValue *= unitConversionFactor;
			}

			XYZ location = wallEndPoints[0] + (offsetDir * offset) + new XYZ(0, 0, zValue);

			return location;
		}
		
		private List<XYZ> getWallEndPointsInFeet(WallComponent wallComponentItem)
		{
			double unitConversionFactor = UnitUtils.ConvertToInternalUnits(1, UnitTypeId.Meters);

			XYZ p1 = new XYZ(wallComponentItem.Wall.VertextIDList[0].X, wallComponentItem.Wall.VertextIDList[0].Y, wallComponentItem.Wall.VertextIDList[0].Z);
			XYZ p2 = new XYZ(wallComponentItem.Wall.VertextIDList[1].X, wallComponentItem.Wall.VertextIDList[1].Y, wallComponentItem.Wall.VertextIDList[1].Z);
			List<XYZ> pointList = new List<XYZ>() { p1, p2 };
			pointList = OrderWallVertices(pointList);
			p1 = pointList.First();
			p2 = pointList.Last();

			//convert to revit unit
			//p1 *= unitConversionFactor;
			//p2 *= unitConversionFactor;

			return new List<XYZ> { p1, p2 };
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
		
		private String GetWindowSymbolFromComponent(WallComponent component, char seperator)
		{
			const int RoundoffValue = 3;
			double width = Math.Round(Convert.ToDouble(component.PropertyDictionary["width"]), RoundoffValue);
			double height = Math.Round(Convert.ToDouble(component.PropertyDictionary["height"]), RoundoffValue);

			String heightString = height.ToString();
			String widthString = width.ToString();

			return widthString + seperator + heightString;
		}
		
		private void GetComponentFamilyData(JObject component, out String familyName, out String symbolName)
		{
			String componentName = component["Name"].ToString().ToLower();
			familyName = componentName;
			symbolName = "default";

			if (componentName.Contains(":") == true)
			{
				var componentNameDesc = componentName.Split(':');
				familyName = componentNameDesc[0];
				symbolName = componentNameDesc[1];
			}
			
			if (component["PropertyDictionary"] != null) {
				JObject properties = component["PropertyDictionary"] as JObject;
				if (properties.ContainsKey("size"))
				{
					symbolName = properties["size"].ToString();
				}
			}
			
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
		
		private void createFloorsCeilings(Building building, Document document, double levelAltitude, FloorType floorType, CeilingType ceilingType, int levelsCount, XYZ normal, int j)
		{
			ElementId levelId = Level.GetNearestLevelId(document, levelAltitude);
			
			//            foreach (var roomItem in building.RoomList)
			//            {
			//                XYZ roomFluidPoint = new XYZ(roomItem.FluidPoint[0], roomItem.FluidPoint[1], 0.0);
			//                //roomFluidPoint *= unitConversionFactor;
//
			//                XYZ roomFloorCentroid = new XYZ();
//
			//                CurveLoop profile = new CurveLoop();
			//                CurveLoop profile_top = new CurveLoop();
//
			//                var vertices = roomItem.VertextList;
//
			//                //double height = UnitUtils.ConvertFromInternalUnits(levelAltitude, UnitTypeId.Meters);
			//                double height = levelAltitude;
//
			//                for (int i = 0; i < vertices.Count; i++)
			//                {
			//                    var vertex1 = vertices[i];
			//                    XYZ p1 = new XYZ(vertex1.X, vertex1.Y, vertex1.Z);
			//                    XYZ p11 = new XYZ(vertex1.X, vertex1.Y, height);
//
			//                    //p1 *= unitConversionFactor;
			//                    //p11 *= unitConversionFactor;
//
			//                    int index = (i + 1) < vertices.Count ? i + 1 : 0;
			//                    var vertex2 = vertices[index];
//
			//                    XYZ p2 = new XYZ(vertex2.X, vertex2.Y, vertex2.Z);
			//                    XYZ p22 = new XYZ(vertex2.X, vertex2.Y, height);
//
			//                    //p2 *= unitConversionFactor;
			//                    //p22 *= unitConversionFactor;
//
			//                    roomFloorCentroid += p1;
//
			//                    var c = Line.CreateBound(p1, p2) as Curve;
			//                    var c1 = Line.CreateBound(p11, p22) as Curve;
//
//
			//                    profile.Append(c);
			//                    profile_top.Append(c1);
			//                }
//
			//                roomFloorCentroid /= vertices.Count;
			//                roomFloorCentroid += new XYZ(0, 0, levelAltitude / 2);
			//                roomFluidPoint += new XYZ(0, 0, levelAltitude / 2);
//
//
//
			//                Floor floor = Floor.Create(document, new List<CurveLoop> { profile }, floorType.Id, levelId);
//
//
			//                if(j == levelsCount)
			//                {
			//                	Ceiling ceiling = Ceiling.Create(document, new List<CurveLoop> { profile }, ceilingType.Id, levelId);
			//                }
//
			//            }
			
			//            foreach (var slabVertexList in building.SlabVertexList) {
//
			//            	 CurveLoop profile = new CurveLoop();
			//            	for (int i = 0; i < slabVertexList.Count; i++)
			//                {
			//                    var vertex1 = slabVertexList[i];
			//                    XYZ p1 = new XYZ(vertex1.U, vertex1.V, 0.0);
//
			//                    int index = (i + 1) < slabVertexList.Count ? i + 1 : 0;
			//                    var vertex2 = slabVertexList[index];
//
			//                    XYZ p2 = new XYZ(vertex2.U, vertex2.V, 0.0);
//
			//                    var c = Line.CreateBound(p1, p2) as Curve;
			//                    profile.Append(c);
			//                }
//
			//            	Floor floor = Floor.Create(document, new List<CurveLoop> { profile }, floorType.Id, levelId);
			//            }
			
		}
	}
}