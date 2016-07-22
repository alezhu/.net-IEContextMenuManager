using System;
using System.IO;
using System.IO.IsolatedStorage;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace IEContextMenuManager
{
	/// <summary>
	/// Summary description for Settings.
	/// </summary>
//	[Serializable]
	public class Settings
	{
		#region DoNotEditThis
		public Settings()
		{
			//
			// TODO: Add constructor logic here
			//
		}
		public Settings(string Section) {
			this.Section  = Section;
		}
		private string _Section;
		[XmlIgnore]
		public string Section {
			get {return _Section;}
            set
            {
                if (_Section == value)
                    return;
                _Section = value;
            }
		}

		public void Save() {
			XmlSerializer ser = new XmlSerializer(this.GetType());
			using(IsolatedStorageFileStream fs = new IsolatedStorageFileStream(Section + ".xml",
					  FileMode.Create,IsolatedStorageFile.GetUserStoreForDomain())) {
				ser.Serialize(fs,this);
			}
		}
		public Settings Load() {
			return Load(this.Section, this.GetType());
		}
		public static Settings Load(string Section, Type type) {
			XmlSerializer ser = new XmlSerializer(type);
			Settings result = null;
			try {
				using (IsolatedStorageFileStream fs = new IsolatedStorageFileStream(Section+".xml",
						   FileMode.Open,IsolatedStorageFile.GetUserStoreForDomain())) {
					result = (Settings)ser.Deserialize(fs);
				}
				result.Section = Section;
			} catch(FileNotFoundException) {
				result = new Settings(Section);
			}
			return result;
		}
		#endregion

		public int FormTop = -1;
		public int FormLeft= -1;
		public int FormWidth = -1;
		public int FormHeight = -1;
		public FormWindowState WindowState = FormWindowState.Normal;
		public int Column1Width= 100;
		public int Column2Width = 200;
	}
}
