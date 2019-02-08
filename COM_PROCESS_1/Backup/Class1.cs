using System;
using System.Runtime.InteropServices;

namespace Tester
{
	[Guid("D6F88E95-8A27-4ae6-B6DE-0542A0FC7039")]
	[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
	public interface _Numbers
	{
		[DispId(1)]
		int GetDay();
		
		[DispId(2)]
		int GetMonth();

		[DispId(3)]
		int GetYear();

		[DispId(4)]
		int DayOfYear();
	}

	[Guid("13FE32AD-4BF8-495f-AB4D-6C61BD463EA4")]
	[ClassInterface(ClassInterfaceType.None)]
	[ProgId("Tester.Numbers")]
	public class Numbers : _Numbers
	{
		public Numbers(){}

		public int GetDay()
		{
			return(DateTime.Today.Day);
		}

		public int GetMonth()
		{
			return(DateTime.Today.Month);
		}

		public int GetYear()
		{
			return(DateTime.Today.Year);
		}

		public int DayOfYear()
		{
			return(DateTime.Now.DayOfYear);
		}
	}
}
