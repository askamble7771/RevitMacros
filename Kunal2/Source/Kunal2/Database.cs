/*
 * Created by SharpDevelop.
 * User: Ashish Kamble
 * Date: 09-07-2024
 * Time: 12:02
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
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
	public enum WallComponentType
	{
		Door,
		Window
	}
	
	public enum WallType
	{
		Curtain,
		Exposed_Wall
	}
	
	public class Element
	{
		public string Name { get; internal set; }
		public String ID { get; set; }
		public long RevitID { get; set; }
		public String RoomName { get; set; }
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
		public Wall(WallType type, string name, string iD, List<Vertex> vertextIDList, double height, double baseOffset, Arc wallArc, double wallThickness = 0.0)
		{
			Type = type;
			Name = name;
			ID = iD;
			VertextIDList = vertextIDList;
			WallHeight = height;
			WallBaseOffset = baseOffset;
			WallThickness = wallThickness;
			WallArc = wallArc;
		}
		public WallType Type { get; set; }
		public List<WallComponent> ItemsOnWallIDList { get; set; }
		public List<Vertex> VertextIDList { get; set; }
		public double WallHeight {get; set;}
		public double WallBaseOffset { get; set; }
		public double WallThickness { get; set; }
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
	
	public class ColumnType : EqualityComparer<ColumnType>
	{
		float[] _dim = new float[2];

		public float H
		{
			get { return _dim[ 0 ]; }
		}

		public float W
		{
			get { return _dim[ 1 ]; }
		}

		public string Name
		{
			get { return "W" + W.ToString() + "X" + H.ToString(); }
		}

		public ColumnType( float d1, float d2 )
		{
			_dim = new float[] {Math.Max(d1, d2), Math.Min(d1, d2)};
		}

		public override bool Equals( ColumnType x, ColumnType y )
		{
			return x.H == y.H && x.W == y.W;
		}

		public override int GetHashCode( ColumnType obj )
		{
			return obj.Name.GetHashCode();
		}
	}
	
	public class HoleType
	{
		public HoleType(float d1, float d2)
		{
			_dim = new float[] { d1, d2 };
		}

		public float H
		{
			get { return _dim[1]; }
		}

		public float W
		{
			get { return _dim[0]; }
		}
		
		private float[] _dim = new float[2];
	}

	public class WindowType : HoleType, IEqualityComparer<WindowType>
	{
		public WindowType(float d1, float d2) : base(d1, d2) { }

		public string Name
		{
			get { return "Window" + W.ToString() + "X" + H.ToString(); }
		}

		public bool Equals(WindowType x, WindowType y)
		{
			if (x == null && y == null) return true;
			if (x == null || y == null) return false;
			return x.H == y.H && x.W == y.W;
		}

		public int GetHashCode(WindowType obj)
		{
			if (obj == null) return 0;
			return obj.H.GetHashCode() ^ obj.W.GetHashCode();
		}
	}

	public class DoorType : HoleType, IEqualityComparer<DoorType>
	{
		public DoorType(float d1, float d2) : base(d1, d2) { }

		public string Name
		{
			get { return "Door" + W.ToString() + "X" + H.ToString(); }
		}

		public bool Equals(DoorType x, DoorType y)
		{
			if (x == null && y == null) return true;
			if (x == null || y == null) return false;
			return x.H == y.H && x.W == y.W;
		}

		public int GetHashCode(DoorType obj)
		{
			if (obj == null) return 0;
			return obj.H.GetHashCode() ^ obj.W.GetHashCode();
		}
	}
	
	/// <summary>
	/// Description of Database.
	/// </summary>
	public class Database
	{
		public Database()
		{
		}
	}
}
