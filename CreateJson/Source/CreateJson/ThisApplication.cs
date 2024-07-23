/*
 * Created by SharpDevelop.
 * User: Ashish Kamble
 * Date: 11-07-2024
 * Time: 18:30
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

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace CreateJson
{
	public class MainClass
	{
		public Dictionary<string, BuildingTypeParameter> buildingTypeData = new Dictionary<string, BuildingTypeParameter>();
		public Dictionary<string, SpaceTypeParameter> spaceTypeData = new Dictionary<string, SpaceTypeParameter>();
		public Dictionary<string, List<string>> conceptualType = new Dictionary<string, List<string>>();
		public Dictionary<string, List<string>> schematicType = new Dictionary<string, List<string>>();
	}
	
	public class BuildingTypeParameter
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
		public string infiltrationAirflowPerArea {get; set;}
		public string plenumLightingContribution{get; set;}
		
		public List<string> OccupancyScheduleList {get; set;}
		public List<string> LightingScheduleList {get; set;}
		public List<string> PowerScheduleList {get; set;}

		// Constructor (if needed)
		public BuildingTypeParameter()
		{
			// Initialize properties if required
		}
	}

	public class SpaceTypeParameter
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
		
		public List<string> occupancyScheduleList {get; set;}
		public List<string> lightingScheduleList {get; set;}
		public List<string> PowerScheduleList {get; set;}
		public string infiltrationAirflowPerArea {get; set;}
		public string plenumLightingContribution{get; set;}
		public double AirChangesPerHour { get; set; }
		
		
		// Constructor (if needed)
		public SpaceTypeParameter()
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

	[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
	[Autodesk.Revit.DB.Macros.AddInId("438726CE-9D82-4388-833C-A3CC63827CDF")]
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
		
		public void GetDataForAnalysis()
		{
			Document _doc = this.ActiveUIDocument.Document;
			

			IList<Element> buildingTypeCollector = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_HVAC_Load_Building_Types).ToList();
			MainClass mainClass = new MainClass();			
			FilteredElementCollector scheduleCollector = new FilteredElementCollector(_doc);
			List<Element> hvacSheduleList = scheduleCollector.OfCategory(BuiltInCategory.OST_HVAC_Load_Schedules).ToList();
//			List<Element> hvacSheduleList = scheduleCollector.OfCategory(BuiltInCategory.).ToList();
			List<string> scheduleNameList = hvacSheduleList.Select(element => element.Name).ToList();
			List<string> scheduleList = new List<string>();
			
			foreach (var element in hvacSheduleList) {
				scheduleList.Add(element.ToString());				
			}

			foreach (Element ele in buildingTypeCollector)
			{
				string plenumLightingContribution = ele.get_Parameter(BuiltInParameter.ROOM_PLENUM_LIGHTING_PARAM).Definition.Name;
				string infiltrationAirFlowPerArea = ele.get_Parameter(BuiltInParameter.SPACE_INFILTRATION_PARAM).Definition.Name;
				Parameter para = ele.get_Parameter(BuiltInParameter.ROOM_OUTDOOR_AIRFLOW_STANDARD_PARAM);
				Element e = para.Element;
				
				
//				BuiltInParameterGroup builtInParameterGroup = para.Definition.ParameterGroup;
				
				BuildingTypeParameter data = new BuildingTypeParameter();
				HVACLoadBuildingType loadBuildingType = ele as HVACLoadBuildingType;

				//Commented parameters below are not avaialable in the Revit API
				data.AreaPerPerson = loadBuildingType.AreaPerPerson;
				data.SensibleHeatGainPerPerson = loadBuildingType.SensibleHeatGainPerPerson;
				data.LatentHeatGainPerPerson = loadBuildingType.LatentHeatGainPerPerson;
				data.LightingLoadDensity = loadBuildingType.LightingLoadDensity;
				data.PowerLoadDensity = loadBuildingType.PowerLoadDensity;
				data.infiltrationAirflowPerArea = infiltrationAirFlowPerArea;
				data.plenumLightingContribution = plenumLightingContribution;
				data.OccupancyScheduleList = scheduleNameList;
				data.LightingScheduleList = scheduleNameList;
				data.PowerScheduleList = scheduleNameList;
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
				
				mainClass.buildingTypeData.Add(ele.Name, data);
			}

			IList<Element> spaceTypeCollector = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_HVAC_Load_Space_Types).ToList();
			foreach (Element ele in spaceTypeCollector)
			{
				string plenumLightingContribution = ele.get_Parameter(BuiltInParameter.ROOM_PLENUM_LIGHTING_PARAM).Definition.Name;
				string infiltrationAirFlowPerArea = ele.get_Parameter(BuiltInParameter.SPACE_INFILTRATION_PARAM).Definition.Name;
//				string infiltrationAirFlowPerArea = ele.get_Parameter(BuiltInParameter.).Definition.Name;
				
				SpaceTypeParameter revitLoadSpaceType = new SpaceTypeParameter();
				HVACLoadSpaceType loadSpaceType = ele as HVACLoadSpaceType;

				//Commented parameters below are not avaialable in the Revit API
				revitLoadSpaceType.AreaPerPerson = loadSpaceType.AreaPerPerson;
				revitLoadSpaceType.SensibleHeatGainPerPerson = loadSpaceType.SensibleHeatGainPerPerson;
				revitLoadSpaceType.LatentHeatGainPerPerson = loadSpaceType.LatentHeatGainPerPerson;
				revitLoadSpaceType.LightingLoadDensity = loadSpaceType.LightingLoadDensity;
				revitLoadSpaceType.PowerLoadDensity = loadSpaceType.PowerLoadDensity;
				revitLoadSpaceType.infiltrationAirflowPerArea = infiltrationAirFlowPerArea;	
				revitLoadSpaceType.plenumLightingContribution = plenumLightingContribution;
				revitLoadSpaceType.occupancyScheduleList = scheduleNameList;
				revitLoadSpaceType.lightingScheduleList = scheduleNameList;
				revitLoadSpaceType.PowerScheduleList = scheduleNameList;
				revitLoadSpaceType.OutdoorAirPerPerson = loadSpaceType.OutdoorAirPerPerson;
				revitLoadSpaceType.OutdoorAirPerArea = loadSpaceType.OutdoorAirPerArea;
				revitLoadSpaceType.AirChangesPerHour = loadSpaceType.AirChangesPerHour;
				//Outdoor Air Method
				revitLoadSpaceType.HeatingSetPoint = loadSpaceType.HeatingSetPoint;
				revitLoadSpaceType.CoolingSetPoint = loadSpaceType.CoolingSetPoint;
				revitLoadSpaceType.HumidificationSetPoint = loadSpaceType.HumidificationSetPoint;
				revitLoadSpaceType.DehumidificationSetPoint = loadSpaceType.DehumidificationSetPoint;
				
				
				mainClass.spaceTypeData.Add(ele.Name, revitLoadSpaceType);
			}
			
			IList<Element> conceptualSurfaceTypeCollector = new FilteredElementCollector(_doc).OfClass(typeof(ConceptualSurfaceType)).ToList();

			foreach (Element ele in conceptualSurfaceTypeCollector)
			{
				ConceptualSurfaceType cct = ele as ConceptualSurfaceType;
				ICollection<ElementId> constructionTypesIds = cct.GetConstructionTypeIds();
				List<string> elementNames = new List<string>();
				foreach (ElementId elementId in constructionTypesIds) {
					elementNames.Add(_doc.GetElement(elementId).Name);
				}
				mainClass.conceptualType.Add(ele.Name, elementNames);
			}
			
			
			ElementId id = EnergyDataSettings.GetBuildingConstructionSetElementId(_doc);
			MEPBuildingConstruction mep = _doc.GetElement(id) as MEPBuildingConstruction;
			
			List<ConstructionType> constructionTypeList = new List<ConstructionType>();
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
			try
			{
				for (int i = 0; i < constructionTypeList.Count; i++)
				{
					ICollection<Construction> constructions = mep.GetConstructions(constructionTypeList[i]);
				List<string> elementNames = new List<string>();
					
					foreach(Construction construction in constructions)
					{
						elementNames.Add(construction.Name);
					}
					mainClass.schematicType.Add(constructionTypeList[i].ToString(),elementNames);
				}
			}
			catch (Exception e)
			{
				TaskDialog.Show("Exception", e.Message);
			}
			
			string buildingJson = JsonConvert.SerializeObject(mainClass, Formatting.Indented);
			System.IO.File.WriteAllText("E:\\buildingA.json", buildingJson);
		}
	}
}