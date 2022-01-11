using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.WorksheetProtectorUnprotector
{
	public class WsProtectorUnprotector : IWsProtectorUnprotector
	{
		private readonly IDialogService dialogService;

		public WsProtectorUnprotector(IDialogService dialogService)
		{
			this.dialogService = dialogService;
		}

		public void ProtectUnprotectSelectedWorksheets(Workbook wb)
		{

			MessageBox.Show("sdf");
		}

		/*
		 
		 

		public void UnprotectWorkbookWithAllWorksheets(string path)
		{
			using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
			{
				UnprotectOpenXml(zip);
				UnlockVbaFromZip(zip);
			}
		}

		private void UnprotectOpenXml(ZipArchive zip)
		{
			// unable to use DocumentFormat.OpenXml because it needs IsolatedStorage for large documents
			// when COM creates the AppDomain instance it doesn’t provide any evidence
			// so the call to create an IsolatedStorage stream fails
			// https://www.lyquidity.com/devblog/?p=65


			var wb = zip.Entries.First(x => x.Name.Contains("workbook.xml") && !x.FullName.Contains("rels"));

			RemoveProtectionNodes(wb, "workbookProtection");

			var worksheets = zip.Entries.Where(x => x.FullName.Contains("xl/worksheets/")
													&& !x.FullName.Contains("rels")
													&& x.FullName.Contains(".xml"))
										.ToArray();


			foreach (var worksheet in worksheets)
			{
				RemoveProtectionNodes(worksheet, "sheetProtection");
			}

		}

		private void RemoveProtectionNodes(ZipArchiveEntry entry, string nodeName)
		{

			using (var stream = entry.Open())
			{
				var xml = new XmlDocument();
				xml.Load(stream);

				var nodes = xml.DocumentElement
					.GetElementsByTagName(nodeName)
					.Cast<XmlElement>().ToArray();
				foreach (var node in nodes)
				{
					node.ParentNode.RemoveChild(node);
				}

				stream.SetLength(0);
				xml.Save(stream);
			}
		}

		private void UnlockVbaFromZip(ZipArchive zip)
		{

			var vba = zip.Entries.FirstOrDefault(x => x.Name.Contains("vbaProject.bin"));
			if (vba == null)
			{
				return;
			}

			using (var vbaStream = vba.Open())
			{
				UnlockVba(vbaStream);
			}

			dialogService.ShowInfo(UIStrings.VbaUnprotected);
		}

		private void UnlockVba(Stream stream)
		{

			using (var memoryStream = new MemoryStream())
			{
				stream.CopyTo(memoryStream);

				// find position of "DPB"
				var ss = memoryStream.ToArray();
				var bytes = Encoding.UTF8.GetBytes("DPB"); // 68, 80, 66
				var position = StartingIndex(ss, bytes);

				if (position.Count == 0)
				{
					return;
				}


				// replace "DPB" to "DPx"
				ss[position.First() + 2] = Encoding.UTF8.GetBytes("x").First(); //120
				stream.Position = 0;
				stream.Write(ss, 0, ss.Length);
			}
		}

		private static IReadOnlyCollection<int> StartingIndex(byte[] main, byte[] sequence)
		{
			var index = Enumerable.Range(0, main.Length - sequence.Length + 1);
			for (var i = 0; i < sequence.Length; i++)
			{
				index = index.Where(n => main[n + i] == sequence[i]).ToArray();
			}
			return index.ToArray();
		}
		*/

	}
}
