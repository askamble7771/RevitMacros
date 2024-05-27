/*
 * Created by SharpDevelop.
 * User: Ashish Kamble
 * Date: 24-04-2024
 * Time: 11:14
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
	/// <summary>
	/// Description of CreateWalls.
	/// </summary>
	public class CreateWalls
	{
		public CreateWalls()
		{
			TaskDialog.Show("Result", "CreateWalls constructor called");
		}
		
		public static void createWalls()
		{
			
			
		}

	}
}
