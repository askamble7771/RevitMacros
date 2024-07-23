/*
 * Created by SharpDevelop.
 * User: Ashish Kamble
 * Date: 29-02-2024
 * Time: 17:01
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
using System.IO;
using Autodesk.Revit.DB.Architecture;
using System.Dynamic;

namespace EnergyAnalysis
{
	[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
	[Autodesk.Revit.DB.Macros.AddInId("EC16DC1A-503C-4E03-86B0-61FE06820197")]
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
		
		//This functions adds Air System, Water system in the model. Creates system zone and add model to it
		public void AddSystems()
		{
			Document doc = this.ActiveUIDocument.Document;

			using (Transaction trans = new Transaction(doc))
			{
				trans.Start("Set Water Loop");
				
				//The code below is with the reference to Priority 1 which adding Air System to the Revit model
				AnalyticalSystemDomain airDomain = AnalyticalSystemDomain.AirSystem;
				MEPAnalyticalSystem mepAirSystem = MEPAnalyticalSystem.Create(doc, airDomain, "airSystem");
				AirSystemData airSystem = mepAirSystem.GetAirSystemData();
				airSystem.CoolingCoilType = AirCoolingCoilType.DirectExpansion;
				airSystem.HeatingCoilType = AirHeatingCoilType.ElectricResistance;
				airSystem.AirFanType = AirFanType.VariableVolume;
				
//		        AnalyticalSystemDomain waterDomain = AnalyticalSystemDomain.WaterLoop;
//		        MEPAnalyticalSystem mepChilledWater = MEPAnalyticalSystem.Create(doc, waterDomain, "CW");
//		        WaterLoopData chilledWaterLoop = mepChilledWater.GetWaterLoopData();
//		        chilledWaterLoop.WaterLoopType = WaterLoopType.ChilledWater;
//		        chilledWaterLoop.ChillerType = WaterChillerType.AirCooled;
//
//		        MEPAnalyticalSystem mepHotWater = MEPAnalyticalSystem.Create(doc, waterDomain, "HW");
//		        WaterLoopData hotWaterLoop = mepHotWater.GetWaterLoopData();
//		        hotWaterLoop.WaterLoopType = WaterLoopType.HotWater;
				
				//Following the code adding the zone equipment to the Revit model
				List<string> str = new List<string>(){"EquipmentType", "EquipmentBehavior"};
				ZoneEquipment zoneEquipment = ZoneEquipment.Create(doc, "ZoneEquipment-1");
				ZoneEquipmentData zoneEquipmentData = zoneEquipment.GetZoneEquipmentData();
				zoneEquipmentData.EquipmentType = ZoneEquipmentHvacType.CAVBox;
				zoneEquipmentData.EquipmentBehavior = ZoneEquipmentBehavior.OnePerSpace;
				zoneEquipmentData.HeatingCoilType = AirHeatingCoilType.ElectricResistance;
				zoneEquipmentData.AirSystemId = mepAirSystem.Id;
				
				
//		        zoneEquipmentData.ChilledWaterLoopId = mepChilledWater.Id;
//		        zoneEquipmentData.HotWaterLoopId = mepHotWater.Id;
				
//		        SystemZoneData systemZoneData = SystemZoneData.Create();
//		        systemZoneData.ZoneEquipmentId = zoneEquipment.Id;
				
				SystemZoneData genericZoneDomainData = SystemZoneData.Create();
				ElementId levelId = Level.GetNearestLevelId(doc, 4.0);
				
				//Code to collect start and end point of the wall
				FilteredElementCollector collector = new FilteredElementCollector(doc);
				ICollection<ElementId> wallIds = collector.OfClass(typeof(Wall)).ToElementIds();
				List<Curve> curves = new List<Curve>();
				
//				FilteredElementCollector collector1 = new FilteredElementCollector(doc);
//				ICollection<ElementId> wallIds1 = collector.OfCategory(BuiltInCategory.wall).ToElementIds();
				
				foreach (ElementId id in wallIds)
				{
					Element element = doc.GetElement(id);
					if (element is Wall)
					{
						Wall wall = element as Wall;
						
						if (wall.WallType.Name == "Exposed_Wall")
						{
							//location curve of the wall
							LocationCurve locationCurve = wall.Location as LocationCurve;
							if (locationCurve != null)
							{
								// Get the start and end points of the wall
								XYZ startPoint = locationCurve.Curve.GetEndPoint(0);
								XYZ endPoint = locationCurve.Curve.GetEndPoint(1);
								
								// Create a line using start and end points
								Line line = Line.CreateBound(startPoint, endPoint);
								curves.Add(line);
							}
						}
					}
				}
				// Convert each Curve into a CurveLoop
				List<CurveLoop> curveLoops = new List<CurveLoop>();
				
				foreach (Curve curve in curves)
				{
					// Assuming you want to create a closed CurveLoop from each Curve
					CurveLoop curveLoop = new CurveLoop();
//				    curveLoop.Append(curve);
					curveLoop.Append(Line.CreateBound(curve.GetEndPoint(1), curve.GetEndPoint(0))); // Close the loop
					curveLoops.Add(curveLoop);
				}
				
				GenericZone genericZone = GenericZone.Create(doc,"systemZone1", genericZoneDomainData, levelId, curveLoops);
				genericZoneDomainData.ZoneEquipmentId = zoneEquipment.Id;
				
				trans.Commit();
			}
		}
		
		//Adds spaces in it
		public void AddSpaces()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			List<Level> m_levels = new List<Level>();
			Dictionary<ElementId, List<Space>> spaceDictionary = new Dictionary<ElementId, List<Space>>();
			Dictionary<ElementId, List<Zone>> zoneDictionary = new Dictionary<ElementId, List<Zone>>();
			Dictionary<ElementId, List<Room>> roomDictionary = new Dictionary<ElementId, List<Room>>();
			
			using (Transaction trans = new Transaction(doc))
			{
				trans.Start("Add Zone");
				
				FilteredElementIterator levelsIterator = (new FilteredElementCollector(doc)).OfClass(typeof(Level)).GetElementIterator();
				FilteredElementIterator spacesIterator =(new FilteredElementCollector(doc)).WherePasses(new SpaceFilter()).GetElementIterator();
				FilteredElementIterator zonesIterator = (new FilteredElementCollector(doc)).OfClass(typeof(Zone)).GetElementIterator();
				
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
				
				spacesIterator.Reset();
				while (spacesIterator.MoveNext())
				{
					Space space = spacesIterator.Current as Space;
					if (space != null)
					{
						spaceDictionary[space.LevelId].Add(space);
					}
				}
				
				zonesIterator.Reset();
				while (zonesIterator.MoveNext())
				{
					Zone zone = zonesIterator.Current as Zone;
					if (zone != null && doc.GetElement(zone.LevelId) != null)
					{
						zoneDictionary[zone.LevelId].Add(zone);
					}
				}
				
				Level m_currentLevel = m_levels[0];
				Parameter para = doc.ActiveView.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.VIEW_PHASE);
				Autodesk.Revit.DB.ElementId phaseId = para.AsElementId();
				Phase m_defaultPhase = doc.GetElement(phaseId) as Phase;
				
				try
				{
					ICollection<ElementId> elements = doc.Create.NewSpaces2(m_currentLevel, m_defaultPhase, doc.ActiveView);
					foreach (ElementId elem in elements)
					{
						Space space = doc.GetElement(elem) as Space;
						if (space != null)
						{
							spaceDictionary[m_currentLevel.Id].Add(space);
						}
					}
					if (elements == null || elements.Count == 0)
					{
						Autodesk.Revit.UI.TaskDialog.Show("Revit", "There is no enclosed loop in " + m_currentLevel.Name);
					}
					else{
						Autodesk.Revit.UI.TaskDialog.Show("Revit", "Spaces added successfully ");
					}
					
				}
				catch (Exception ex)
				{
					Autodesk.Revit.UI.TaskDialog.Show("Revit", ex.Message);
				}
				trans.Commit();
			}
		}
		
		//Creates energy model
		public void CreateEnergyModel()
		{
			
			Document doc = this.ActiveUIDocument.Document;
			
			// Collect space and surface data from the building's analytical thermal model
			EnergyAnalysisDetailModelOptions options = new EnergyAnalysisDetailModelOptions();
			options.Tier = EnergyAnalysisDetailModelTier.Final; // include constructions, schedules, non-graphical data
			options.EnergyModelType = EnergyModelType.SpatialElement; // Energy model based on rooms or spaces
			
			using (Transaction trans = new Transaction(doc))
			{
				trans.Start("Add Zone");
				
				//Turn on volume calculation in Areas and Volume Computations
				AreaVolumeSettings areaVolumeSettings = AreaVolumeSettings.GetAreaVolumeSettings(doc);
				areaVolumeSettings.ComputeVolumes = true;
				
				EnergyDataSettings energyDataSetting = EnergyDataSettings.GetFromDocument(doc);
				if (energyDataSetting.AnalysisType == AnalysisMode.ConceptualMassesAndBuildingElements) {
					energyDataSetting.AnalysisType = AnalysisMode.RoomsOrSpaces;
				}
				energyDataSetting.IncludeThermalProperties = true;
//				ConceptualConstructionType ca = null;
//
//				ConceptualConstructionType cs = null;
//				ConceptualConstructionFloorSlabType.HighMassConstructionColdClimateSlabInsulation;
				////					= ConstructionType.Floor;

				//setting for changing building type
				energyDataSetting.BuildingType  = gbXMLBuildingType.Gymnasium;
				
				//Extra part that can add the building in the in the Building Type
//				HVACLoadType hvacLoadType  = HVACLoadSpaceType.Create(doc, "Space1");
//				HVACLoadType hvacLoadType  = HVACLoadBuildingType.Create(doc, "Building1");
//				hvacLoadType.AreaPerPerson = 60.0;
//				hvacLoadType.SensibleHeatGainPerPerson = 251.00;
//				hvacLoadType.LatentHeatGainPerPerson = 221.00;

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
//				TaskDialog.Show("EAM", builder.ToString());
				SystemsAnalysisOptions sa = new SystemsAnalysisOptions();
				string weatherFileName = sa.WeatherFile;
//				TaskDialog.Show("WeatherFile", "Weather file name is " + weatherFileName);
				trans.Commit();
			}
		}
		
		//Generates report
		public void GenearateReport()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			ViewSystemsAnalysisReport newReport = null;

			SystemsAnalysisOptions theOptions = new SystemsAnalysisOptions();
			string openStudioPath = Path.GetFullPath(Path.Combine(doc.Application.SystemsAnalysisWorkfilesRootPath, @"..\"));
			// The following are the default values for systems analysis if not specified.
			// If the weather file is not specified, the analysis will use the weather at the current site location.
			string[] filePaths= { @"workflows\HVAC Systems Loads and Sizing.osw", @"workflows\Annual Building Energy Simulation.osw"};
			
			foreach(string filePath in filePaths)
			{
				using (Transaction transaction = new Transaction(doc))
				{
					transaction.Start("Create Systems Analysis View");
					theOptions.WorkflowFile = Path.Combine(openStudioPath, filePath);
					
					//It takes the .epw file of Pune Location
					theOptions.WeatherFile = Path.GetFullPath(@"E:\WeatherFile_PuneLocation\IND_Pune.430630_ISHRAE.epw");
					theOptions.OutputFolder = Path.GetFullPath(@"E:\Revit_SystemAnalysis\SystemAnalysis_DownloadReports");
//					theOptions.OutputFolder = Path.GetTempPath();
					
					newReport = ViewSystemsAnalysisReport.Create(doc, "APITestView13");
					// Create a new report of systems analysis.
					if (newReport != null)
					{
						newReport.RequestSystemsAnalysis(theOptions);
						// Request the systems analysis in the background process. When the systems analysis is completed,
						// the result is automatically updated in the report view and the analytical space elements.
						// You may check the status by calling newReport.IsAnalysisCompleted().
						transaction.Commit();
					}
					else
					{
						transaction.RollBack();
					}
				}
			}
			
		}
		
		//Get model details
		public void GetModelDetails()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			EnergyAnalysisDetailModel eadm = EnergyAnalysisDetailModel.GetMainEnergyAnalysisDetailModel(doc);
			bool  status = eadm.SimplifyCurtainSystems;
			
			EnergyAnalysisDetailModelTier eadmTier = eadm.Tier;
			
			EnergyDataSettings energyDataSetting = EnergyDataSettings.GetFromDocument(doc);
			AnalysisMode analysisMode = energyDataSetting.AnalysisType;
			
			IList<EnergyAnalysisOpening> energyAnalysisOpeningList = eadm.GetAnalyticalOpenings();
			IList<EnergyAnalysisSurface> energyAnalysisSurface = eadm.GetAnalyticalShadingSurfaces();
			IList<EnergyAnalysisSpace> energyAnalysisSpace = eadm.GetAnalyticalSpaces();
			
		}
		
		//This function is combination of 4-5 function above which did all operations of system Analysis
		public void SystemAnalysis()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			List<Level> m_levels = new List<Level>();
			Dictionary<ElementId, List<Space>> spaceDictionary = new Dictionary<ElementId, List<Space>>();
			Dictionary<ElementId, List<Zone>> zoneDictionary = new Dictionary<ElementId, List<Zone>>();
			
			using (Transaction trans = new Transaction(doc))
			{
				trans.Start("Add Zone");
				
				FilteredElementIterator levelsIterator = (new FilteredElementCollector(doc)).OfClass(typeof(Level)).GetElementIterator();
				FilteredElementIterator spacesIterator =(new FilteredElementCollector(doc)).WherePasses(new SpaceFilter()).GetElementIterator();
				FilteredElementIterator zonesIterator = (new FilteredElementCollector(doc)).OfClass(typeof(Zone)).GetElementIterator();
				
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
				
				spacesIterator.Reset();
				while (spacesIterator.MoveNext())
				{
					Space space = spacesIterator.Current as Space;
					if (space != null)
					{
						spaceDictionary[space.LevelId].Add(space);
					}
				}
				
				zonesIterator.Reset();
				while (zonesIterator.MoveNext())
				{
					Zone zone = zonesIterator.Current as Zone;
					if (zone != null && doc.GetElement(zone.LevelId) != null)
					{
						zoneDictionary[zone.LevelId].Add(zone);
					}
				}
				
				Level m_currentLevel = m_levels[0];
				Parameter para = doc.ActiveView.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.VIEW_PHASE);
				Autodesk.Revit.DB.ElementId phaseId = para.AsElementId();
				Phase m_defaultPhase = doc.GetElement(phaseId) as Phase;
				
				try
				{
					ICollection<ElementId> elements = doc.Create.NewSpaces2(m_currentLevel, m_defaultPhase, doc.ActiveView);
					foreach (ElementId elem in elements)
					{
						Space space = doc.GetElement(elem) as Space;
						if (space != null)
						{
							spaceDictionary[m_currentLevel.Id].Add(space);
						}
					}
					if (elements == null || elements.Count == 0)
					{
						Autodesk.Revit.UI.TaskDialog.Show("Revit", "There is no enclosed loop in " + m_currentLevel.Name);
					}
					else{
						Autodesk.Revit.UI.TaskDialog.Show("Revit", "Spaces added successfully ");
					}
					
				}
				catch (Exception ex)
				{
					Autodesk.Revit.UI.TaskDialog.Show("Revit", ex.Message);
				}
				
				
				
				trans.Commit();
			}
			
			using (Transaction trans = new Transaction(doc))
			{
				trans.Start("Set Water Loop");
				
				//The code below is with the reference to Priority 1
				AnalyticalSystemDomain airDomain = AnalyticalSystemDomain.AirSystem;
				MEPAnalyticalSystem mepAirSystem = MEPAnalyticalSystem.Create(doc, airDomain, "airSystem");
				AirSystemData airSystem = mepAirSystem.GetAirSystemData();
				airSystem.CoolingCoilType = AirCoolingCoilType.DirectExpansion;
				airSystem.HeatingCoilType = AirHeatingCoilType.ElectricResistance;
				airSystem.AirFanType = AirFanType.VariableVolume;
				
//		        AnalyticalSystemDomain waterDomain = AnalyticalSystemDomain.WaterLoop;
//		        MEPAnalyticalSystem mepChilledWater = MEPAnalyticalSystem.Create(doc, waterDomain, "CW");
//		        WaterLoopData chilledWaterLoop = mepChilledWater.GetWaterLoopData();
//		        chilledWaterLoop.WaterLoopType = WaterLoopType.ChilledWater;
//		        chilledWaterLoop.ChillerType = WaterChillerType.AirCooled;
//
//		        MEPAnalyticalSystem mepHotWater = MEPAnalyticalSystem.Create(doc, waterDomain, "HW");
//		        WaterLoopData hotWaterLoop = mepHotWater.GetWaterLoopData();
//		        hotWaterLoop.WaterLoopType = WaterLoopType.HotWater;
				
				ZoneEquipment zoneEquipment = ZoneEquipment.Create(doc, "ZoneEquipment-1");
				ZoneEquipmentData zoneEquipmentData = zoneEquipment.GetZoneEquipmentData();
				zoneEquipmentData.EquipmentType = ZoneEquipmentHvacType.CAVBox;
				zoneEquipmentData.EquipmentBehavior = ZoneEquipmentBehavior.OnePerSpace;
				zoneEquipmentData.HeatingCoilType = AirHeatingCoilType.ElectricResistance;
				zoneEquipmentData.AirSystemId = mepAirSystem.Id;
//		        zoneEquipmentData.ChilledWaterLoopId = mepChilledWater.Id;
//		        zoneEquipmentData.HotWaterLoopId = mepHotWater.Id;
				
//		        SystemZoneData systemZoneData = SystemZoneData.Create();
//		        systemZoneData.ZoneEquipmentId = zoneEquipment.Id;
				
				SystemZoneData genericZoneDomainData = SystemZoneData.Create();
				ElementId levelId = Level.GetNearestLevelId(doc, 4.0);
				
				//Code to collect start and end point of the wall
				FilteredElementCollector collector = new FilteredElementCollector(doc);
				ICollection<ElementId> wallIds = collector.OfClass(typeof(Wall)).ToElementIds();
				List<Curve> curves = new List<Curve>();
				
				foreach (ElementId id in wallIds)
				{
					Element element = doc.GetElement(id);
					if (element is Wall)
					{
						Wall wall = element as Wall;
						
						if (wall.WallType.Name.Contains("Exposed_Wall"))
						{
							//location curve of the wall
							LocationCurve locationCurve = wall.Location as LocationCurve;
							if (locationCurve != null)
							{
								// Get the start and end points of the wall
								XYZ startPoint = locationCurve.Curve.GetEndPoint(0);
								XYZ endPoint = locationCurve.Curve.GetEndPoint(1);
								
								// Create a line using start and end points
								Line line = Line.CreateBound(startPoint, endPoint);
								curves.Add(line);
							}
						}
					}
				}
				// Convert each Curve into a CurveLoop
				List<CurveLoop> curveLoops = new List<CurveLoop>();
				foreach (Curve curve in curves)
				{
					// Assuming you want to create a closed CurveLoop from each Curve
					CurveLoop curveLoop = new CurveLoop();
//				    curveLoop.Append(curve);
					curveLoop.Append(Line.CreateBound(curve.GetEndPoint(1), curve.GetEndPoint(0))); // Close the loop
					curveLoops.Add(curveLoop);
				}
				
				IList<CurveLoop> curveLoopList = curveLoops;
				
				GenericZone genericZone = GenericZone.Create(doc,"systemZone1", genericZoneDomainData,levelId, curveLoopList);
				genericZoneDomainData.ZoneEquipmentId = zoneEquipment.Id;
				
				trans.Commit();
			}
			
			// Collect space and surface data from the building's analytical thermal model
			EnergyAnalysisDetailModelOptions options = new EnergyAnalysisDetailModelOptions();
			options.Tier = EnergyAnalysisDetailModelTier.Final; // include constructions, schedules, non-graphical data
			options.EnergyModelType = EnergyModelType.SpatialElement; // Energy model based on rooms or spaces
			options.SimplifyCurtainSystems = true;
			
			using (Transaction trans = new Transaction(doc))
			{
				trans.Start("Add Zone");
				
				EnergyDataSettings energyDataSetting = EnergyDataSettings.GetFromDocument(doc);
				if (energyDataSetting.AnalysisType == AnalysisMode.ConceptualMassesAndBuildingElements) {
					energyDataSetting.AnalysisType = AnalysisMode.RoomsOrSpaces;
				}
				energyDataSetting.IncludeThermalProperties = true;
				
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
			
			ViewSystemsAnalysisReport newReport = null;

			SystemsAnalysisOptions theOptions = new SystemsAnalysisOptions();
			string openStudioPath = Path.GetFullPath(Path.Combine(doc.Application.SystemsAnalysisWorkfilesRootPath, @"..\"));
			// The following are the default values for systems analysis if not specified.
			// If the weather file is not specified, the analysis will use the weather at the current site location.
			string[] filePaths= { @"workflows\HVAC Systems Loads and Sizing.osw", @"workflows\Annual Building Energy Simulation.osw"};
			
			foreach(string filePath in filePaths)
			{
				using (Transaction transaction = new Transaction(doc))
				{
					transaction.Start("Create Systems Analysis View");
					theOptions.WorkflowFile = Path.Combine(openStudioPath, filePath);
					theOptions.OutputFolder = Path.GetFullPath(@"E:\Revit_SystemAnalysis\SystemAnalysis_DownloadReports");
//					theOptions.OutputFolder = Path.GetTempPath();
					
					newReport = ViewSystemsAnalysisReport.Create(doc, "APITestView13");
					// Create a new report of systems analysis.
					if (newReport != null)
					{
						newReport.RequestSystemsAnalysis(theOptions);
						// Request the systems analysis in the background process. When the systems analysis is completed,
						// the result is automatically updated in the report view and the analytical space elements.
						// You may check the status by calling newReport.IsAnalysisCompleted().
						transaction.Commit();
					}
					else
					{
						transaction.RollBack();
					}
				}
			}
		}
		
		//This function creates new schedule and changed the parameters in it
		public void ChangeBuildingSchedule()
		{
			
			Document doc = this.ActiveUIDocument.Document;
			
			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("Changing Building Schedule");
//				HVACLoadType hvacLoadType  = HVACLoadSpaceType.Create(doc, "Space1");
				
				if(HVACLoadBuildingType.IsNameUnique(doc, "Office_Duplicate")){
					HVACLoadBuildingType hvacLoadType  = HVACLoadBuildingType.Create(doc, "Office_Duplicate");

					hvacLoadType.AreaPerPerson = 10.0;
					
					// convert user-defined heatGain value to BritishThermalUnitsPerHour from deciwatt prior to setting
					double sensibleHeatGainPerPerson = UnitUtils.Convert(100, UnitTypeId.Watts, UnitTypeId.BritishThermalUnitsPerHour);//check - Value available -> Watts -  required->deciwatt
					hvacLoadType.SensibleHeatGainPerPerson = sensibleHeatGainPerPerson;//check - Value changing - deciwatt
					
					double latentHeatGainPerPerson = UnitUtils.Convert(150, UnitTypeId.Watts, UnitTypeId.BritishThermalUnitsPerHour);//check - Value available -> Watts -  required->deciwatt
					hvacLoadType.LatentHeatGainPerPerson = latentHeatGainPerPerson;//check - Value changing - deciwatt
					
					hvacLoadType.LightingLoadDensity = 0.95;
					
					double powerLoadDensity = UnitUtils.Convert(10.0, UnitTypeId.WattsPerSquareFoot, UnitTypeId.WattsPerSquareMeter);
					hvacLoadType.PowerLoadDensity = powerLoadDensity; //check - Value changing - (this)w/sq.m. to w/sq.ft.(Revit)
					
//					hvacLoadType.InfiltrationAirFlowPerArea = ;
					hvacLoadType.PlenumLighting = 0.25;
					
//					Reference myRef = doc.GetElement(hvacLoadType.Id);
//					Element e = doc.GetElement();
//					Parameter p = e.GetParameter(BuiltInParameter.SPACE_OCCUPANCY_SCHEDULE_PARAM);
//					p.Set(3);
					
					
					//    FilteredElementCollector scheduleCollector = new FilteredElementCollector(doc);
					//        ScheduleDefinition schedule = scheduleCollector.OfClass(typeof(ViewSchedule))
					//                                                       .Cast<ViewSchedule>().FirstOrDefault().Definition;
//
//					Parameter para = hvacLoadType.GetParameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM);
//					para.Set(schedule.CategoryId.Value);
					
//					hvacLoadType.OccupancySchedule = ;
//					hvacLoadType.LightingSchedule = ;
//					hvacLoadType.PowerSchedule = ;
//					SchedulableField sf = ViewSchedule.
//
//					SchedulableField sf = new SchedulableField();      			//It creates seperate schedule fields which is not part of the Lighting or Occupancy schedule
//					sf.FieldType = ScheduleFieldType.
					
					double outdoorAirPerPerson = UnitUtils.Convert(10.0, UnitTypeId.CubicFeetPerMinute, UnitTypeId.CubicFeetPerMinute);//CubicFeetPerSeconds not present
					hvacLoadType.OutdoorAirPerPerson = 12; //check - Value changing - (this)CFS = CFM(Revit)
					hvacLoadType.OutdoorAirPerArea = 0.08; //check - Value changing - (this)CFS/CF = CFM/CF(Revit)
					hvacLoadType.AirChangesPerHour = 0.001;
//					hvacLoadType.OutdoorAirMethod = ;
					hvacLoadType.OpeningTime = "8.00 AM";
					hvacLoadType.ClosingTime = "6.00 PM";
					hvacLoadType.OutdoorAirFlowStandard = OutdoorAirFlowStandard.ByACH;
					
					
					double unoccupiedCoolingSetPoint = UnitUtils.Convert(85, UnitTypeId.Fahrenheit, UnitTypeId.Kelvin);
					hvacLoadType.UnoccupiedCoolingSetPoint = unoccupiedCoolingSetPoint;
					
					double heatingSetPoint = UnitUtils.Convert(70, UnitTypeId.Fahrenheit, UnitTypeId.Kelvin);
					hvacLoadType.HeatingSetPoint= heatingSetPoint;

					double coolingSetPoint = UnitUtils.Convert(85, UnitTypeId.Fahrenheit, UnitTypeId.Kelvin);
					hvacLoadType.CoolingSetPoint = coolingSetPoint;
					hvacLoadType.HumidificationSetPoint = 0.001;
					hvacLoadType.DehumidificationSetPoint = 0.72;

					
					
					
					TaskDialog.Show("Result", "Building Schedule created Successfully");
				}
				else{
					TaskDialog.Show("Result", "It is already exists");
				}
				
				
//					hvacLoadType.SensibleHeatGainPerPerson = 251.00;
//					hvacLoadType.LatentHeatGainPerPerson = 221.00;
//				hvacLoadType.airf
				transaction.Commit();
			}
		}
		
		//I have not written code here
		public void SystemZone()
		{
			
			
		}
		
		//This function deletes the energy model which created by UI (manually)
		public void DeleteEnergyModel()
		{
			Document doc = this.ActiveUIDocument.Document;
			RevitCommandId id = RevitCommandId.LookupPostableCommandId(PostableCommand.DeleteEnergyModel);
			PostCommand(id);
		}
		
		//This function adds single system zones for the given spaces
		public void SpaceSystemZone()
		{
			Document doc = this.ActiveUIDocument.Document;

			using (Transaction transaction = new Transaction(doc))
			{
				try
				{
					transaction.Start("Creating System Zones");

					IEnumerable<Space> spaces = new FilteredElementCollector(doc, doc.ActiveView.Id)
						.WhereElementIsNotElementType()
						.OfCategory(BuiltInCategory.OST_MEPSpaces)
						.Cast<Space>();
					List<CurveLoop> curveLoops = new List<CurveLoop>();

					foreach (Space space in spaces)
					{
//						if( space.Name == "Space 8" || space.Name == "Space 4" || space.Name == "Space 7"){
						
						
						using (GeometryElement geometryElement = space.get_Geometry(new Options())) // Use using statement
						{
							foreach (GeometryObject geometryObject in geometryElement)
							{
								int count1 = 0; // Move inside the loop
								Solid solid = geometryObject as Solid;
								if (solid != null)
								{
									foreach (Face face in solid.Faces)
									{
										IList<XYZ> edgePoints = face.Triangulate().Vertices;

										if (count1 >= 2) // To avoid the roofs and floor surface
										{
											XYZ startPoint = edgePoints[0];
											XYZ endPoint = edgePoints[1];

											CurveLoop curveLoop = new CurveLoop();
											curveLoop.Append(Line.CreateBound(startPoint, endPoint));
											curveLoops.Add(curveLoop);
										}
										count1++;
									}
								}
							}
//							}
						}
					}
					
					// After collecting curves for this space, create the zone
					string systemZoneName = "SystemZone-B";
					SystemZoneData genericZoneDomainData = SystemZoneData.Create();
					ElementId levelId = Level.GetNearestLevelId(doc, 4.0);
					GenericZone genericZone = GenericZone.Create(doc, systemZoneName, genericZoneDomainData, levelId, curveLoops);

					transaction.Commit();
				}
				catch (Exception ex)
				{
					// Handle exceptions
					TaskDialog.Show("Error", ex.Message);
					transaction.RollBack();
				}
			}
		}

		//In this function, we pass the line through centroid of the spaces. The line can include the unnecessary spaces also ehich came intermediate
		public void CentroidSystemZone()
		{
			Document doc = this.ActiveUIDocument.Document;

			using (Transaction transaction = new Transaction(doc))
			{
				try
				{
					transaction.Start("Creating System Zones");

					IEnumerable<Space> spaces = new FilteredElementCollector(doc, doc.ActiveView.Id)
						.WhereElementIsNotElementType()
						.OfCategory(BuiltInCategory.OST_MEPSpaces)
						.Cast<Space>();
					List<CurveLoop> curveLoops = new List<CurveLoop>();
					List<XYZ> centroids = new List<XYZ>();

					foreach (Space space in spaces)
					{
						if (space.Name == "Space 17" || space.Name == "Space 18" || space.Name == "Space 22")
						{
							
							
							using (GeometryElement geometryElement = space.get_Geometry(new Options())) // Use using statement
							{
								foreach (GeometryObject geometryObject in geometryElement)
								{
									Solid solid = geometryObject as Solid;
									if (solid != null)
									{
										XYZ centroid = solid.ComputeCentroid();
										centroids.Add(centroid);
									}
								}
							}
						}
					}
					
					for (int i = 0; i < centroids.Count()-1; i++) {
						XYZ startPoint = centroids[i];
						XYZ endPoint = centroids[i+1];

						CurveLoop curveLoop = new CurveLoop();
						curveLoop.Append(Line.CreateBound(startPoint, endPoint));
						curveLoops.Add(curveLoop);
						
					}
					
					
					// After collecting curves for this space, create the zone
					string systemZoneName = "SystemZone-B";
					SystemZoneData genericZoneDomainData = SystemZoneData.Create();
					ElementId levelId = Level.GetNearestLevelId(doc, 4.0);
					GenericZone genericZone = GenericZone.Create(doc, systemZoneName, genericZoneDomainData, levelId, curveLoops);

					transaction.Commit();
					
				}
				catch (Exception ex)
				{
					// Handle exceptions
					TaskDialog.Show("Error", ex.Message);
					transaction.RollBack();
				}
			}
		}
		
		//It is the new schedule adjecent with Building schedule which has its own paramenters
		public void ChangeSpaceSchedule()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("Changing Building Schedule");
				
				
				if(HVACLoadSpaceType.IsNameUnique(doc, "Space1")){
					HVACLoadSpaceType hvacLoadType = HVACLoadSpaceType.Create(doc, "Space1");
					
					hvacLoadType.AreaPerPerson = 10.0;
					
					// convert user-defined heatGain value to BritishThermalUnitsPerHour from deciwatt prior to setting
					double sensibleHeatGainPerPerson = UnitUtils.Convert(100, UnitTypeId.Watts, UnitTypeId.BritishThermalUnitsPerHour);//check - Value available -> Watts -  required->deciwatt
					hvacLoadType.SensibleHeatGainPerPerson = sensibleHeatGainPerPerson;//check - Value changing - deciwatt
					
					double latentHeatGainPerPerson = UnitUtils.Convert(150, UnitTypeId.Watts, UnitTypeId.BritishThermalUnitsPerHour);//check - Value available -> Watts -  required->deciwatt
					hvacLoadType.LatentHeatGainPerPerson = latentHeatGainPerPerson;//check - Value changing - deciwatt
					
					hvacLoadType.LightingLoadDensity = 0.95;
					
					
					double powerLoadDensity = UnitUtils.Convert(10.0, UnitTypeId.WattsPerSquareFoot, UnitTypeId.WattsPerSquareMeter);
					hvacLoadType.PowerLoadDensity = powerLoadDensity; //check - Value changing - (this)w/sq.m. to w/sq.ft.(Revit)
//					hvacLoadType.infiltrationAirflowPerArea

					hvacLoadType.PlenumLighting = 0.25;

//					schedules
					
					double outdoorAirPerPerson = UnitUtils.Convert(10.0, UnitTypeId.CubicFeetPerMinute, UnitTypeId.CubicFeetPerMinute);//CubicFeetPerSeconds not present
					hvacLoadType.OutdoorAirPerPerson = 12; //check - Value changing - (this)CFS = CFM(Revit)
					hvacLoadType.OutdoorAirPerArea = 0.08; //check - Value changing - (this)CFS/CF = CFM/CF(Revit)
					hvacLoadType.AirChangesPerHour = 0.001;
//					hvacLoadType.OutdoorAirMethod = ;
					
					double heatingSetPoint = UnitUtils.Convert(70, UnitTypeId.Fahrenheit, UnitTypeId.Kelvin);
					hvacLoadType.HeatingSetPoint= heatingSetPoint;

					double coolingSetPoint = UnitUtils.Convert(85, UnitTypeId.Fahrenheit, UnitTypeId.Kelvin);
					hvacLoadType.CoolingSetPoint = coolingSetPoint;
					hvacLoadType.HumidificationSetPoint = 0.001;
					hvacLoadType.DehumidificationSetPoint = 0.72;
					
					TaskDialog.Show("Result", "Building Schedule created Successfully");
				}
				else{
					TaskDialog.Show("Result", "It is already exists");
				}

				
				
				
				
				transaction.Commit();
			}
		}
		
		public void UpdateBuildinSchedule()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			FilteredElementCollector buildingTypeCollector = new FilteredElementCollector(doc);
			IList<Element> buildingTypeElementsList = buildingTypeCollector.OfCategory(BuiltInCategory.OST_HVAC_Load_Building_Types).ToList();
			
			FilteredElementCollector scheduleCollector = new FilteredElementCollector(doc);
			IList<Element> hvacSheduleList = scheduleCollector.OfCategory(BuiltInCategory.OST_HVAC_Load_Schedules).ToList();
//			IList<Element> hvacSheduleList = scheduleCollector.OfCategory(BuiltInCategory.OST_HVAC_).ToList();
			
			ElementId scheduleElementId = new ElementId(137396);
			
			using (Transaction transaction = new Transaction(doc))
			{
				transaction.Start("Update Schedules");
				try
				{
					foreach(Element buildingTypeElement in buildingTypeElementsList)
					{
						HVACLoadType hvacLoadType = buildingTypeElement as HVACLoadType;
						String loadTypeName = hvacLoadType.Name;
						
						if(hvacLoadType.Name == "Office")
						{
							IList<Parameter> orderParamList = buildingTypeElement.GetOrderedParameters();
						
							foreach(Parameter param in orderParamList)
							{								
								if(param.Definition.Name == "Occupancy Schedule")
								{
									//ForgeTypeId forgeTypeId = 
									param.Set(scheduleElementId);
									TaskDialog.Show("Parameter Name",param.AsValueString(), TaskDialogCommonButtons.Ok);
								}
							}
						}	
					}
					
					transaction.Commit();
				}
				catch (Exception ex)
                {
                    // Handle exceptions
                    TaskDialog.Show("Error", ex.Message);
                    transaction.RollBack();
                }
			}
		}
		
		//It adds the equipment for the group spaces
		public void GroupsSpacesEquipment()
		{
			Document doc = this.ActiveUIDocument.Document;

			using (Transaction trans = new Transaction(doc))
			{
				trans.Start("Adding group spaces for single equipment");
				
				//The code below is with the reference to Priority 1 which adding Air System to the Revit model
				AnalyticalSystemDomain airDomain = AnalyticalSystemDomain.AirSystem;
				MEPAnalyticalSystem mepAirSystem = MEPAnalyticalSystem.Create(doc, airDomain, "airSystem");
				AirSystemData airSystem = mepAirSystem.GetAirSystemData();
				airSystem.CoolingCoilType = AirCoolingCoilType.DirectExpansion;
				airSystem.HeatingCoilType = AirHeatingCoilType.ElectricResistance;
				airSystem.AirFanType = AirFanType.VariableVolume;
				
				//Following the code adding the zone equipment to the Revit model
				ZoneEquipment zoneEquipment = ZoneEquipment.Create(doc, "ZoneEquipment-1");
				ZoneEquipmentData zoneEquipmentData = zoneEquipment.GetZoneEquipmentData();
				zoneEquipmentData.EquipmentType = ZoneEquipmentHvacType.CAVBox;
				zoneEquipmentData.EquipmentBehavior = ZoneEquipmentBehavior.GroupSpaces;
				zoneEquipmentData.HeatingCoilType = AirHeatingCoilType.ElectricResistance;
				zoneEquipmentData.AirSystemId = mepAirSystem.Id;
				
				IEnumerable<Space> spaces = new FilteredElementCollector(doc, doc.ActiveView.Id)
					.WhereElementIsNotElementType()
					.OfCategory(BuiltInCategory.OST_MEPSpaces)
					.Cast<Space>();
				List<CurveLoop> curveLoops = new List<CurveLoop>();

				foreach (Space space in spaces)
				{
//						if( space.Name == "Space 8" || space.Name == "Space 2" || space.Name == "Space 7"){
					
					
					using (GeometryElement geometryElement = space.get_Geometry(new Options())) // Use using statement
					{
						foreach (GeometryObject geometryObject in geometryElement)
						{
							int count1 = 0; // Move inside the loop
							Solid solid = geometryObject as Solid;
							if (solid != null)
							{
								foreach (Face face in solid.Faces)
								{
									IList<XYZ> edgePoints = face.Triangulate().Vertices;

									if (count1 >= 2) // To avoid the roofs and floor surface
									{
										XYZ startPoint = edgePoints[0];
										XYZ endPoint = edgePoints[1];

										CurveLoop curveLoop = new CurveLoop();
										curveLoop.Append(Line.CreateBound(startPoint, endPoint));
										curveLoops.Add(curveLoop);
									}
									count1++;
								}
							}
						}
//							}
					}
				}
				
				// After collecting curves for this space, create the zone
				string systemZoneName = "SystemZone-B";
				SystemZoneData genericZoneDomainData = SystemZoneData.Create();
				ElementId levelId = Level.GetNearestLevelId(doc, 4.0);
				GenericZone genericZone = GenericZone.Create(doc, systemZoneName, genericZoneDomainData, levelId, curveLoops);
				
				genericZoneDomainData.ZoneEquipmentId = zoneEquipment.Id;
				
				trans.Commit();
			}
			
		}
		
		//It adds the schedule for very spaces but code is incomplete
		public void AddSpaceSchedule()
		{
			
			
//			Document doc = this.ActiveUIDocument.Document;
//
//			using (Transaction transaction = new Transaction(doc))
//			{
//				try
//				{
//					transaction.Start("Creating System Zones");
//
//					IEnumerable<Space> spaces = new FilteredElementCollector(doc, doc.ActiveView.Id)
//						.WhereElementIsNotElementType()
//						.OfCategory(BuiltInCategory.OST_MEPSpaces)
//						.Cast<Space>();
//					List<CurveLoop> curveLoops = new List<CurveLoop>();
//
//					foreach (Space space in spaces)
//					{
//						space.SpaceType = SpaceType.kActiveStorage;
//
//
//						if( space.Name == "Space 8" || space.Name == "Space 2" || space.Name == "Space 7"){
//
//
//							using (GeometryElement geometryElement = space.get_Geometry(new Options())) // Use using statement
//							{
//								foreach (GeometryObject geometryObject in geometryElement)
//								{
//									int count1 = 0; // Move inside the loop
//									Solid solid = geometryObject as Solid;
//									if (solid != null)
//									{
//										foreach (Face face in solid.Faces)
//										{
//											IList<XYZ> edgePoints = face.Triangulate().Vertices;
//
//											if (count1 >= 2) // To avoid the roofs and floor surface
//											{
//												XYZ startPoint = edgePoints[0];
//												XYZ endPoint = edgePoints[1];
//
//												CurveLoop curveLoop = new CurveLoop();
//												curveLoop.Append(Line.CreateBound(startPoint, endPoint));
//												curveLoops.Add(curveLoop);
//											}
//											count1++;
//										}
//									}
//								}
//							}
//						}
//					}
//
//					// After collecting curves for this space, create the zone
//					string systemZoneName = "SystemZone-B";
//					SystemZoneData genericZoneDomainData = SystemZoneData.Create();
//					ElementId levelId = Level.GetNearestLevelId(doc, 4.0);
//					GenericZone genericZone = GenericZone.Create(doc, systemZoneName, genericZoneDomainData, levelId, curveLoops);
//
//					transaction.Commit();
//				}
//				catch (Exception ex)
//				{
//					// Handle exceptions
//					TaskDialog.Show("Error", ex.Message);
//					transaction.RollBack();
//				}
		}
		

		public void ChangeEnergySetting()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			// Collect space and surface data from the building's analytical thermal model
			EnergyAnalysisDetailModelOptions options = new EnergyAnalysisDetailModelOptions();
			options.Tier = EnergyAnalysisDetailModelTier.Final; // include constructions, schedules, non-graphical data
			options.EnergyModelType = EnergyModelType.SpatialElement; // Energy model based on rooms or spaces
			
			using (Transaction trans = new Transaction(doc))
			{
				trans.Start("Add Zone");
				
				//Turn on volume calculation in Areas and Volume Computations
				AreaVolumeSettings areaVolumeSettings = AreaVolumeSettings.GetAreaVolumeSettings(doc);
				areaVolumeSettings.ComputeVolumes = true;
				
				EnergyDataSettings energyDataSetting = EnergyDataSettings.GetFromDocument(doc);
				if (energyDataSetting.AnalysisType == AnalysisMode.ConceptualMassesAndBuildingElements) {
					energyDataSetting.AnalysisType = AnalysisMode.RoomsOrSpaces;
				}

				//setting for changing building type
				energyDataSetting.BuildingType  = gbXMLBuildingType.Gymnasium;
				
				
				//Code for overriding the material properties
				ElementId id = EnergyDataSettings.GetBuildingConstructionSetElementId(doc);
				MEPBuildingConstruction mep = doc.GetElement(id) as MEPBuildingConstruction;
				mep.SetBuildingConstructionOverride(ConstructionType.ExteriorWindow, true);
				mep.SetBuildingConstructionOverride(ConstructionType.Floor, true);
				mep.SetBuildingConstructionOverride(ConstructionType.InteriorWall, true);

//					mep.SetBuildingConstruction(ConstructionType.Floor, ConceptualConstructionWallType.LightweightConstructionNoInsulationInterior);
//					mep.SetBuildingConstruction(ConstructionType.Floor, ConceptualConstructionWallType.LightweightConstructionNoInsulationInterior);
				
				/*The best overloaded method match for 'Autodesk.Revit.DB.Mechanical.MEPBuildingConstruction.SetBuildingConstruction(Autodesk.Revit.DB.Analysis.ConstructionType, Autodesk.Revit.DB.Construction)' has some invalid arguments (CS1502) -
				 * 						C:\ProgramData\Autodesk\Revit\Macros\2024\Revit\AppHookup\EnergyAnalysis\Source\EnergyAnalysis\ThisApplication.cs:1084,6*/

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
//				TaskDialog.Show("EAM", builder.ToString());
				SystemsAnalysisOptions sa = new SystemsAnalysisOptions();
				string weatherFileName = sa.WeatherFile;
//				TaskDialog.Show("WeatherFile", "Weather file name is " + weatherFileName);
				trans.Commit();
			}
		}
		
		
		public void GetView()
		{
//			View3D view3D;
//			var document = RevitCommandData.Document;
//			using (Transaction transaction = new Transaction(document,"Sample"))
//			{
//				transaction.Start();
//
//				var collector = new FilteredElementCollector(document);
//				var list = collector.OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().
//										First(v => v.ViewFamily == ViewFamily.ThreeDimensional);
//
//				view3D = View3D.CreateIsometric(document, list.Id);
//				view3D.Name = "Sample 3D ";
//
//				transaction.Commit();
//			}
//
//			RevitCommandData.UiDocument.ActiveView = view3D;
		}
		
		
		
		public class RevitData
		{
			// Properties used in GetDataForAnalysis
			public double AreaPerPerson { get; set; }
			public double SensibleHeatGainPerPerson { get; set; }
			public double LatentHeatGainPerPerson { get; set; }
			public double LightingLoadDensity { get; set; }
			public double PowerLoadDensity { get; set; }
			public double OutdoorAirPerPerson { get; set; }
			public double OutdoorAirPerArea { get; set; }
			public double AirChangesPerHour { get; set; }
			public string OpeningTime { get; set; }
			public string ClosingTime { get; set; }
			public double UnoccupiedCoolingSetPoint { get; set; }
			public double HeatingSetPoint { get; set; }
			public double CoolingSetPoint { get; set; }
			public double HumidificationSetPoint { get; set; }
			public double DehumidificationSetPoint { get; set; }

			// Constructor (if needed)
			public RevitData()
			{
				// Initialize properties if required
			}
		}

		public class RevitHVACLoadBuildingType
		{
			// Properties used in GetDataForAnalysis
			public double AreaPerPerson { get; set; }
			public double SensibleHeatGainPerPerson { get; set; }
			public double LatentHeatGainPerPerson { get; set; }
			public double LightingLoadDensity { get; set; }
			public double PowerLoadDensity { get; set; }
			public double OutdoorAirPerPerson { get; set; }
			public double OutdoorAirPerArea { get; set; }
			public double HeatingSetPoint { get; set; }
			public double CoolingSetPoint { get; set; }
			public double HumidificationSetPoint { get; set; }
			public double DehumidificationSetPoint { get; set; }

			// Constructor (if needed)
			public RevitHVACLoadBuildingType()
			{
				// Initialize properties if required
			}
		}

		public class RevitConceptualConstructionType
		{
			// Properties used in GetDataForAnalysis
			public Category Category { get; set; }
			public string FamilyName { get; set; }
			public Location Location { get; set; }
			public string Name { get; set; }

			// Constructor (if needed)
			public RevitConceptualConstructionType()
			{
				// Initialize properties if required
			}
		}

		public void GetDataForAnalysis()
		{
			Document _doc = this.ActiveUIDocument.Document;

			IList<Element> buildingTypeCollector = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_HVAC_Load_Building_Types).ToList();
			List<RevitData> dataList = new List<RevitData>();

			foreach (Element ele in buildingTypeCollector)
			{
				RevitData data = new RevitData();
				HVACLoadBuildingType loadBuildingType = ele as HVACLoadBuildingType;

				//Commented parameters below are not avaialable in the Revit API
				data.AreaPerPerson = loadBuildingType.AreaPerPerson;
				data.SensibleHeatGainPerPerson = loadBuildingType.SensibleHeatGainPerPerson;
				data.LatentHeatGainPerPerson = loadBuildingType.LatentHeatGainPerPerson;
				data.LightingLoadDensity = loadBuildingType.LightingLoadDensity;
				data.PowerLoadDensity = loadBuildingType.PowerLoadDensity;
				//Infiltration Airflow per area
				//Plenum lighting contribution'
				//Occupancy schedule
				//Lighting Schedule
				data.OutdoorAirPerPerson = loadBuildingType.OutdoorAirPerPerson;
				data.OutdoorAirPerArea = loadBuildingType.OutdoorAirPerArea;
				data.AirChangesPerHour = loadBuildingType.AirChangesPerHour;
				//Outdoor Air Method
				data.OpeningTime = loadBuildingType.OpeningTime;
				data.ClosingTime = loadBuildingType.ClosingTime;
				data.UnoccupiedCoolingSetPoint = loadBuildingType.UnoccupiedCoolingSetPoint;
				data.HeatingSetPoint = loadBuildingType.HeatingSetPoint;
				data.CoolingSetPoint = loadBuildingType.CoolingSetPoint;
				data.HumidificationSetPoint = loadBuildingType.HumidificationSetPoint;
				data.DehumidificationSetPoint = loadBuildingType.DehumidificationSetPoint;

				dataList.Add(data);
			}

			IList<Element> spaceTypeCollector = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_HVAC_Load_Space_Types).ToList();
			List<RevitHVACLoadBuildingType> hvacLoadSpaceType = new List<RevitHVACLoadBuildingType>();

			foreach (Element ele in spaceTypeCollector)
			{
				RevitHVACLoadBuildingType revitLoadSpaceType = new RevitHVACLoadBuildingType();
				HVACLoadSpaceType loadSpaceType = ele as HVACLoadSpaceType;

				//Commented parameters below are not avaialable in the Revit API
				revitLoadSpaceType.AreaPerPerson = loadSpaceType.AreaPerPerson;
				revitLoadSpaceType.SensibleHeatGainPerPerson = loadSpaceType.SensibleHeatGainPerPerson;
				revitLoadSpaceType.LatentHeatGainPerPerson = loadSpaceType.LatentHeatGainPerPerson;
				revitLoadSpaceType.LightingLoadDensity = loadSpaceType.LightingLoadDensity;
				revitLoadSpaceType.PowerLoadDensity = loadSpaceType.PowerLoadDensity;
				//Infiltration Airflow per area
				//Plenum lighting contribution
				//Occupancy schedule
				//Lighting Schedule
				revitLoadSpaceType.OutdoorAirPerPerson = loadSpaceType.OutdoorAirPerPerson;
				revitLoadSpaceType.OutdoorAirPerArea = loadSpaceType.OutdoorAirPerArea;
				//AirChangesPerHour
				//Outdoor Air Method
				revitLoadSpaceType.HeatingSetPoint = loadSpaceType.HeatingSetPoint;
				revitLoadSpaceType.CoolingSetPoint = loadSpaceType.CoolingSetPoint;
				revitLoadSpaceType.HumidificationSetPoint = loadSpaceType.HumidificationSetPoint;
				revitLoadSpaceType.DehumidificationSetPoint = loadSpaceType.DehumidificationSetPoint;

				hvacLoadSpaceType.Add(revitLoadSpaceType);
			}

			IList<Element> conceptualConstructionTypeCollector = new FilteredElementCollector(_doc).OfClass(typeof(ConceptualConstructionType)).ToList();
			List<RevitConceptualConstructionType> conceptualConstructionType = new List<RevitConceptualConstructionType>();

			foreach (Element ele in conceptualConstructionTypeCollector)
			{
				RevitConceptualConstructionType revitConceptualConstructionType = new RevitConceptualConstructionType();
				ConceptualConstructionType hvBt = ele as ConceptualConstructionType;

				// Populate RevitConceptualConstructionType properties from ConceptualConstructionType
				revitConceptualConstructionType.Category = hvBt.Category;
				revitConceptualConstructionType.FamilyName = hvBt.FamilyName;
				revitConceptualConstructionType.Location = hvBt.Location;
				revitConceptualConstructionType.Name = hvBt.Name;

				conceptualConstructionType.Add(revitConceptualConstructionType);
			}


			IList<Element> windowTypes = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_MassGlazingAll).ToList();
			foreach(Element ele in windowTypes)
			{
				
			}
			
			IList<Element> massTypeCollector = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_Mass).ToList();
			foreach(Element ele in massTypeCollector)
			{
				//string name = ele.Name;
				//ConceptualConstructionType hvBt = ele as ConceptualConstructionType;
				//string conceptualConstructionTypeName = hvBt.FamilyName;
			}
			
			ElementId id = EnergyDataSettings.GetBuildingConstructionSetElementId(_doc);
			MEPBuildingConstruction mep = _doc.GetElement(id) as MEPBuildingConstruction;
			List<ConstructionType> constructionTypeList = new List<ConstructionType>();
			Dictionary<string, List<string>> constructionTypeDict = new Dictionary<string, List<string>>();
			
			
			constructionTypeList.Add(ConstructionType.Ceiling);
			constructionTypeList.Add(ConstructionType.Door);
			constructionTypeList.Add(ConstructionType.ExteriorWall);
			constructionTypeList.Add(ConstructionType.ExteriorWindow);
			constructionTypeList.Add(ConstructionType.Floor);
			constructionTypeList.Add(ConstructionType.InteriorWall);
			constructionTypeList.Add(ConstructionType.InteriorWindow);
			constructionTypeList.Add(ConstructionType.Roof);
			constructionTypeList.Add(ConstructionType.Skylight);
			constructionTypeList.Add(ConstructionType.Slab);
			constructionTypeList.Add(ConstructionType.UndergroundWall);
			try {
				for (int i = 0; i < constructionTypeList.Count; i++)
				{
					ICollection<Construction> constructions = mep.GetConstructions(constructionTypeList[i]);
					List<string> constructionNames = new List<string>();
					
					foreach(Construction construction in constructions)
					{
						constructionNames.Add(construction.Name);
					}
					string constructionTypeName = constructionTypeList[i].ToString();
					constructionTypeDict.Add(constructionTypeName, constructionNames);
				}
			} catch (Exception e) {
				TaskDialog.Show("Exception", e.Message);
			}
			
			
			
			
			for (int i = 0; i < 5; i++) {
				
			}
			
		}
		
	}
}
