/*
 * Created by SharpDevelop.
 * User: Ashish Kamble
 * Date: 22-04-2024
 * Time: 17:12
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
	/// <summary>
	/// Description of Class1.
	/// </summary>
	public class Class1
	{
//		public Class1()
//		{
//			
//				TaskDialog.Show("Result", "Hello World!!!");
//			
//		}
		
		public static void hello(){
			TaskDialog.Show("result" , "Hello World!!!");
		}
	}
}
