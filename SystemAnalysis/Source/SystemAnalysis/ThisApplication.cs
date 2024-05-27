/*
 * Created by SharpDevelop.
 * User: Ashish Kamble
 * Date: 13-02-2024
 * Time: 12:09
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Analysis;
using System.Text;

namespace SystemAnalysis
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("DD768D52-1D13-4A43-B817-F877E31FB221")]
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
		public void openFile()
		{
			Document doc = this.ActiveUIDocument.Document;
			using (Transaction trans = new Transaction(doc))
			{
			
				string pathName = ".\\E:\\Revit_SystemAnalysis\\Ashish\\input.rvt";
				
				FilePath path = new FilePath(pathName);
			    RevitLinkOptions options = new RevitLinkOptions(false);
			    // Create new revit link storing absolute path to file
			    LinkLoadResult result = RevitLinkType.Create(doc, path, options);
			    //result.Dispose();
			    trans.Commit();
			}
		    
		}
		public void LocationAPI()
		{ 
			Document document = this.ActiveUIDocument.Document;
			
		    // Get the project location handle 
		    ProjectLocation projectLocation = document.ActiveProjectLocation;
		
		    // Show the information of current project location
		    XYZ origin = new XYZ(0, 0, 0);
		    ProjectPosition position = projectLocation.GetProjectPosition(origin);
		    if (null == position)
		    {
		        throw new Exception("No project position in origin point.");
		    }
		
		    // Format the prompt string to show the message.
		    String prompt = "Current project location information:\n";
		    prompt += "\n\t" + "Origin point position:";
		    prompt += "\n\t\t" + "Angle: " + position.Angle;
		    prompt += "\n\t\t" + "East to West offset: " + position.EastWest;
		    prompt += "\n\t\t" + "Elevation: " + position.Elevation;
		    prompt += "\n\t\t" + "North to South offset: " + position.NorthSouth;
		
		    // Angles are in radians when coming from Revit API, so we 
		    // convert to degrees for display
		    const double angleRatio = Math.PI / 180;   // angle conversion factor
		
		    SiteLocation site = projectLocation.GetSiteLocation();
		    prompt += "\n\t" + "Site location:";
		    prompt += "\n\t\t" + "Latitude: " + site.Latitude / angleRatio + "��";
		    prompt += "\n\t\t" + "Longitude: " + site.Longitude / angleRatio + "��";
		    prompt += "\n\t\t" + "TimeZone: " + site.TimeZone;
		
		    // Give the user some information
		    TaskDialog.Show("Revit",prompt);

		}
		
		public void GetProjectLocation()
		{
			Document document = this.ActiveUIDocument.Document;
			ProjectLocation currentLocation = document.ActiveProjectLocation;

		    //get the project position
		    XYZ origin = new XYZ(0, 0, 0);
		
		    const double angleRatio = Math.PI / 180;   // angle conversion factor
		
		    ProjectPosition projectPosition = currentLocation.GetProjectPosition(origin);
		    //Angle from True North
		    double angle = 0 * angleRatio;   // convert degrees to radian
		    double eastWest =73.8567;     //East to West offset
		    double northSouth = 18.5204;   //North to South offset
		    double elevation = 560.0;    //Elevation above ground level
		
		    //create a new project position
		    ProjectPosition newPosition =
		      document.Application.Create.NewProjectPosition(eastWest, northSouth, elevation, angle);
		
		    if (null != newPosition)
		    {
		        //set the value of the project position
		        currentLocation.SetProjectPosition(origin, newPosition);
		    }

		}
		
		
		
		public void SystemZone()
		{
			Document activeDoc = this.ActiveUIDocument.Document;
			
			List<Level> m_levels = new List<Level>();
			Dictionary<ElementId, List<Space>> spaceDictionary = new Dictionary<ElementId, List<Space>>();
			Dictionary<ElementId, List<Zone>> zoneDictionary = new Dictionary<ElementId, List<Zone>>();
						
			FilteredElementIterator levelsIterator = (new FilteredElementCollector(activeDoc)).OfClass(typeof(Level)).GetElementIterator();
			FilteredElementIterator spacesIterator =(new FilteredElementCollector(activeDoc)).WherePasses(new SpaceFilter()).GetElementIterator();
			FilteredElementIterator zonesIterator = (new FilteredElementCollector(activeDoc)).OfClass(typeof(Zone)).GetElementIterator();
			          
			levelsIterator.Reset();
			while (levelsIterator.MoveNext())
			{
			    Level level = levelsIterator.Current as Level;
			    if (level != null)
			    {
			        m_levels.Add(level);
			        spaceDictionary.Add(level.Id, new List<Space>());
			        zoneDictionary.Add(level.Id, new List<Zone>());
			    }
			}
			
			while (levelsIterator.MoveNext())
			{
					//string zoneName = "New Zone";
					Level level = levelsIterator.Current as Level;
					ElementId eleId = level.Id;

				Level level_1 = activeDoc.GetElement(eleId) as Level;
				
				// Define zone geometry using points (modify as needed)
				XYZ pt1 = new XYZ(-1.295676, 10.728976, 0);
				XYZ pt2 = new XYZ(-1.295676, 40.728978, 0);
				XYZ pt3 = new XYZ(-51.295676, 40.728978, 0);
				XYZ pt4 = new XYZ(-51.295676, 10.728976, 0);
				
				// Create a list of points for the zone boundary
				List<XYZ> points = new List<XYZ>() { pt1, pt2, pt3, pt4, pt1 };
				
//				// Create a new zone element
//				FamilySymbol zoneFamily = new FilteredElementCollector(activeDoc)
//                             .OfClass(typeof(FamilySymbol))
//                             .OfCategory(BuiltInCategory.OST_Zo)
//                             .FirstOrDefault(x => x.Name == "Generic Zone") as FamilySymbol;
//
//				// Create the zone element with the specified family symbol
//				Zone zone = activeDoc.Create.NewFamilyInstance(points, level, zoneFamily) as Zone;
//								
//				// Commit the changes in a transaction
//				using (Transaction trans = new Transaction(activeDoc))
//				{
//				  trans.Start("Add Zone");
//				  activeDoc.Insert(zone);
//				  trans.Commit();
//				}
				
				activeDoc.Regenerate();
				break;
			}
	}
		
		public void CreateZone(Level level, Phase phase, Document doc)
		{
			Dictionary<ElementId, List<Zone>> m_zoneDictionary = new Dictionary<ElementId, List<Zone>>();
		    Zone zone = doc.Create.NewZone(level, phase);
		    if (zone != null)
		    {
		        m_zoneDictionary[level.Id].Add(zone);
		    }
		}
		public void CreateEnergyAnalysis()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			// Collect space and surface data from the building's analytical thermal model
			EnergyAnalysisDetailModelOptions options = new EnergyAnalysisDetailModelOptions();
			options.Tier = EnergyAnalysisDetailModelTier.Final; // include constructions, schedules, and non-graphical data in the computation of the energy analysis model
			options.EnergyModelType = EnergyModelType.SpatialElement;   // Energy model based on rooms or spaces
			
			using (Transaction trans = new Transaction(doc))
				{
				  trans.Start("Add Zone");
				                  
				  EnergyAnalysisDetailModel eadm = EnergyAnalysisDetailModel.Create(doc, options); 
				  IList<EnergyAnalysisSpace> spaces = eadm.GetAnalyticalSpaces();
					StringBuilder builder = new StringBuilder();
					builder.AppendLine("Spaces: " + spaces.Count);
					foreach (EnergyAnalysisSpace space in spaces)
					{
					    SpatialElement spatialElement = doc.GetElement(space.CADObjectUniqueId) as SpatialElement;
					    ElementId spatialElementId = spatialElement == null ? ElementId.InvalidElementId : spatialElement.Id;
					    builder.AppendLine("   >>> " + space.SpaceName + " related to " + spatialElementId);
					    IList<EnergyAnalysisSurface> surfaces = space.GetAnalyticalSurfaces();
					    builder.AppendLine("       has " + surfaces.Count + " surfaces.");
					    foreach (EnergyAnalysisSurface surface in surfaces)
					    {
					        builder.AppendLine("            +++ Surface from " + surface.OriginatingElementDescription);
					    }
					}
				  TaskDialog.Show("EAM", builder.ToString());
				  trans.Commit();
				}
				
		}

		
		
		
		public void CreateGenericZone()
		{
			
			Document doc = this.ActiveUIDocument.Document;
			GenericZone genericZone = null;
			//GenericZoneDomainData domainData = Autodesk.Revit.DB.Analysis.GenericZoneDomainData;//GenericZone.GetDomainData();
			
			
			IList<Element> levelsIterator = (new FilteredElementCollector(doc)).OfClass(typeof(Level)).ToList();
			ElementId elementId = levelsIterator.First().Id;
			IList<CurveLoop> curveLoops = new List<CurveLoop>();
			//GeometryElement zoneGeometry = Zone.
			
			FilteredElementCollector wallCollector = new FilteredElementCollector(doc); 
            List<Element> walls = wallCollector
                .OfCategory(BuiltInCategory.OST_Walls) 
                .OfClass(typeof(Wall)) 
                .ToList();
            
            using (Transaction trans = new Transaction(doc))
				{
            	trans.Start();
		            foreach (var wall in walls) 
		            { 
		            	CurveLoop profileloop = new CurveLoop();
		                Wall wallData = wall as Wall;
		                LocationCurve wallLocation = wall.Location as LocationCurve; 
		                
		                Line line = Line.CreateBound(wallLocation.Curve.GetEndPoint(0), wallLocation.Curve.GetEndPoint(1));
		                profileloop.Append(line);
		                curveLoops.Add(profileloop);
		            }
		
		           //genericZone  = GenericZone.Create(doc, "newZone",  domainData, elementId, curveLoops);
		           	trans.Commit();
				}
		}
		
 }
}