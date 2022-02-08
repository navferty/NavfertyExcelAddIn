#nullable enable

#region Structures to interoperate with the Windows API

using word = System.UInt16;
using dword = System.UInt32;
using hwnd = System.IntPtr;
using large_int = System.Int64;
using ulong_ptr = System.IntPtr;

#endregion

namespace Navferty.Common.WinAPI
{

	internal static class Core
	{
		public const string WINDLL_USER = "user32.dll";
	}

}
