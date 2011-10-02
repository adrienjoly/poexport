using System;
using System.IO;
using System.Reflection;
using System.Collections;
using PocketOutlook;

namespace POOMClient
{
	// Used to get contacts list using GetDefaultFolder method
	/// <summary>
	/// Summary description for POExport.
	/// </summary>
	public class POExport
	{
		private const int folderContactsConst = 10;

		private string separator = ";";
		private string lineSeparator = "\n";

		public POExport()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		static public void Main()
		{
			POExport po = new POExport();
			po.ExportWithTemplate("\\contacts.csv", "\\csvtemplate.csv");
		}

		public void ExportWithTemplate(string csvFilename, string templateFilename)
		{
			PropertyInfo[] fields;
			string header = "";
			Type cType = Type.GetType("PocketOutlook.Contact"); //pa.GetType();

			// open the output file
			FileStream fs = new FileStream(csvFilename, FileMode.Create);

			/*
			// Include all properties
			Type cType = Type.GetType("PocketOutlook.Contact");
			PropertyInfo[] fields = cType.GetProperties();

			// write the header: fields names
			{
				// add the name of each field in the header record
				foreach (PropertyInfo f in fields)
					header += ProtectValue(f.Name) + separator;
			}
			*/

			// read a custom mapping of fields
			ArrayList fieldsNames = new ArrayList();
			StreamReader template = new StreamReader(templateFilename);
			for(;;)
			{
				string line = template.ReadLine();
				if (line == null || line.Length == 0) break;

				string[] record = line.Split(',');
				fieldsNames.Add(UnprotectValue(record[0]));

				string label = record.Length > 1 ? record[1] : record[0];
				header += /*ProtectValue*/(label) + separator;
			}

			// Replace the last comma with a CR then save into file
			byte[] data = EncodeRecord(header.Substring(0, header.Length - separator.Length) + lineSeparator);
			fs.Write(data, 0, data.Length);

			fields = new PropertyInfo[fieldsNames.Count];

			for(int i=0; i<fieldsNames.Count; ++i)
			{
				fields[i] = cType.GetProperty((string) fieldsNames[i]);
			}

			//fields = (PropertyInfo[]) fieldsNames.ToArray(Type.GetType("PropertyInfo"));

			Export(fields, fs);
		}

		private void Export(PropertyInfo[] fields, FileStream fs)
		{
			try 
			{
				// The Application object is the only object which can be viewed
				// by external libraries. Other PocketOutlook objects are created
				// by calling various methods on the Application object.
				PocketOutlook.Application app = new PocketOutlook.Application();
            
				// Log the user onto a Pocket Outlook session 
				app.Logon();

				// get contacts info
				ItemCollection poicContactsCollection = app.GetDefaultFolder(folderContactsConst).Items;

				PocketOutlook.Contact pa;
				            
				// Add all the Contacts to a ListView.
				for (int i = 0; i < poicContactsCollection.Count; i++) 
				{
					pa = (PocketOutlook.Contact) poicContactsCollection.Item(i + 1);      // Starting with first item, item 'zero' does not exist
                    
					// Add contact information to the list view
					//string[] displayInfo = new string[3];
					//this.setupListViewItem(pa, out displayInfo[0], out displayInfo[1], out displayInfo[2]);
					//this.lvContacts.Items.Add(new ListViewItem(displayInfo));

					// Put the data values into a record string
					string record = "";
					foreach (PropertyInfo f in fields)
						record += ProtectValue(f.GetValue(pa, null).ToString()) + separator;

					// Replace the last comma with a CR then save into file
					byte[] data = EncodeRecord(record.Substring(0, record.Length - separator.Length) + lineSeparator);
					fs.Write(data, 0, data.Length);
				}
                
				// log the user out
				app.Logoff();
			}
			catch (Exception exception) 
			{
				System.Windows.Forms.MessageBox.Show(exception.ToString());
			}

		}

		private string ProtectValue (string str)
		{
			if (str.Length > 0)
			{
				str = str.Replace("\n", "; ");
				str = str.Replace("\r", "");
				str = str.Replace("\"", "''");
				str = "\"" + str + "\"";
			}
			return str;
		}

		private string UnprotectValue (string str)
		{
			if (str.StartsWith("\"") && str.EndsWith("\""))
			{
				str = str.Substring(1, str.Length-2);
			}
			return str;
		}

		private byte[] EncodeRecord (string str)
		{
			return System.Text.Encoding.Default.GetBytes(str);
		}
	}
}
