using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;

namespace NavfertyCommon
{
	[DebuggerStepThrough]
	public static class EnumExtensions
	{

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static string GetEnumDescription(this Enum value)
		{
			var fi = value.GetType().GetField(value.ToString());
			var attributes = fi.GetCustomAttributes(typeof(DescriptionAttribute), false) as DescriptionAttribute[];

			if (attributes != null && attributes.Any())
			{
				return attributes.First().Description;
			}

			return value.ToString();
		}
	}
}
