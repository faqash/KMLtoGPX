using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Text;
using Microsoft.Win32;
using System.Reflection;


namespace KML2GPX
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtFileName;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnExit;
		private System.Windows.Forms.TextBox txtOutputFolder;
		private System.Windows.Forms.Button btnExtract;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
		private System.Windows.Forms.Button btnBrowseOpenFile;
		private System.Windows.Forms.Button btnBrowseFolder;
		private System.Windows.Forms.Button button1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Form1));
			this.label1 = new System.Windows.Forms.Label();
			this.txtFileName = new System.Windows.Forms.TextBox();
			this.txtOutputFolder = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.btnExit = new System.Windows.Forms.Button();
			this.btnExtract = new System.Windows.Forms.Button();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
			this.btnBrowseOpenFile = new System.Windows.Forms.Button();
			this.btnBrowseFolder = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(80, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "KML Filename:";
			// 
			// txtFileName
			// 
			this.txtFileName.Location = new System.Drawing.Point(128, 8);
			this.txtFileName.Name = "txtFileName";
			this.txtFileName.Size = new System.Drawing.Size(448, 20);
			this.txtFileName.TabIndex = 2;
			this.txtFileName.Text = "";
			// 
			// txtOutputFolder
			// 
			this.txtOutputFolder.Location = new System.Drawing.Point(128, 40);
			this.txtOutputFolder.Name = "txtOutputFolder";
			this.txtOutputFolder.Size = new System.Drawing.Size(448, 20);
			this.txtOutputFolder.TabIndex = 5;
			this.txtOutputFolder.Text = "";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(104, 16);
			this.label2.TabIndex = 4;
			this.label2.Text = "Output Folder:";
			// 
			// btnExit
			// 
			this.btnExit.Location = new System.Drawing.Point(312, 96);
			this.btnExit.Name = "btnExit";
			this.btnExit.Size = new System.Drawing.Size(120, 23);
			this.btnExit.TabIndex = 6;
			this.btnExit.Text = "Exit";
			this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
			// 
			// btnExtract
			// 
			this.btnExtract.Location = new System.Drawing.Point(176, 96);
			this.btnExtract.Name = "btnExtract";
			this.btnExtract.Size = new System.Drawing.Size(120, 23);
			this.btnExtract.TabIndex = 9;
			this.btnExtract.Text = "Convert";
			this.btnExtract.Click += new System.EventHandler(this.btnExtract_Click);
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.DefaultExt = "kml";
			this.openFileDialog1.Filter = "KML|*.kml";
			this.openFileDialog1.Title = "Select KML file to convert";
			// 
			// btnBrowseOpenFile
			// 
			this.btnBrowseOpenFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnBrowseOpenFile.Location = new System.Drawing.Point(584, 8);
			this.btnBrowseOpenFile.Name = "btnBrowseOpenFile";
			this.btnBrowseOpenFile.Size = new System.Drawing.Size(24, 18);
			this.btnBrowseOpenFile.TabIndex = 10;
			this.btnBrowseOpenFile.Text = "...";
			this.btnBrowseOpenFile.Click += new System.EventHandler(this.btnBrowseOpenFile_Click);
			// 
			// btnBrowseFolder
			// 
			this.btnBrowseFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnBrowseFolder.Location = new System.Drawing.Point(584, 40);
			this.btnBrowseFolder.Name = "btnBrowseFolder";
			this.btnBrowseFolder.Size = new System.Drawing.Size(24, 18);
			this.btnBrowseFolder.TabIndex = 11;
			this.btnBrowseFolder.Text = "...";
			this.btnBrowseFolder.Click += new System.EventHandler(this.btnBrowseFolder_Click);
			// 
			// button1
			// 
			this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button1.Location = new System.Drawing.Point(584, 112);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(24, 18);
			this.button1.TabIndex = 12;
			this.button1.Text = "?";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(618, 143);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.btnBrowseFolder);
			this.Controls.Add(this.btnBrowseOpenFile);
			this.Controls.Add(this.btnExtract);
			this.Controls.Add(this.btnExit);
			this.Controls.Add(this.txtOutputFolder);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtFileName);
			this.Controls.Add(this.label1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "KML to GPX Converter";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Form1_Closing);
			this.Load += new System.EventHandler(this.Form1_Load);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			GetSettings();
		}
		
		
		private void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			SaveSettings();
		}		
		
		
		private void btnExit_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		
		private void btnExtract_Click(object sender, System.EventArgs e)
		{						
			string sOutFilePath = txtFileName.Text.Substring(txtFileName.Text.LastIndexOf(@"\")+1, txtFileName.Text.Length - txtFileName.Text.LastIndexOf(@"\") - 1);
			string sOutFileName = sOutFilePath.Substring(0, sOutFilePath.LastIndexOf("."));
			sOutFilePath = txtOutputFolder.Text + @"\" + sOutFileName + ".gpx";

			if (CheckPaths(txtFileName.Text, txtOutputFolder.Text))
			{
				
				string sFileName = txtFileName.Text;
				bool bFileWasFixed = FixXML(ref sFileName);

				bool bProcessWorked = ProcessKML(sFileName, sOutFilePath, sOutFileName);
				if (bFileWasFixed)
				{
					if (File.Exists(sFileName))
					{
						File.Delete(sFileName);
					}
				}
				if (bProcessWorked)
				{
					MessageBox.Show("Conversion Completed Successfully.");
				}
			}
		}	
		
		
		private bool CheckPaths(string sFilename, string sOutputFolder)
		{
			if (!File.Exists(sFilename))
			{
				MessageBox.Show("Input file does not exists.");
				return false;
			}
			if (!Directory.Exists(sOutputFolder))
			{
				MessageBox.Show("Output Folder does not exists.");
				return false;
			}
			return true;
		}

		
		private bool FixXML(ref string sFilename)
		{
			// For some reason the xml does not read correctly with "xmlns=" in the file
			// So a crude fix for now is to remove it from the text and work with a new temp file

			System.IO.StreamReader srr = File.OpenText(sFilename);
			string s = srr.ReadToEnd();
			srr.Close();
			if (s.IndexOf("xmlns=") > 0)
			{
				s = s.Replace("xmlns=","xmlns:t=");
				sFilename = Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\KML2GPX.tmp";

				if (File.Exists(sFilename))
				{
					File.Delete(sFilename);
				}

				using (StreamWriter sw = new StreamWriter(sFilename))
				{
					sw.Write(s);
					return true;
				}
			}
			else
			{
				return false;
			}
		}

		
		private bool ProcessKML(string sFilename, string sOutputFilePath, string sOutputFileName)
		{
						
			string sName = "";

			try
			{						
				XPathDocument doc = new XPathDocument(sFilename);

				XPathNavigator nav = doc.CreateNavigator();

				StreamWriter sr = null;

//				<kml xmlns="http://earth.google.com/kml/2.0">
//				<Document xmlns:xlink="http://www.w3/org/1999/xlink">
//				<name>Grand Canyon</name>

				XPathNodeIterator ni = nav.Select("/kml/Document/name");
				string sDocName = "";
				if(ni.MoveNext())
				{
					sDocName = ni.Current.Value;
				}
				if (sDocName == "")
				{
					sDocName = sOutputFileName;
				}				

				sr = File.CreateText(sOutputFilePath);
			
				sr.WriteLine(@"<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""yes""?>");
				sr.WriteLine(@"<gpx");
				sr.WriteLine(@"version=""1.1""");
				sr.WriteLine(@"creator=""KML2GPX""");
				sr.WriteLine(@"xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""");
				sr.WriteLine(@"xmlns=""http://www.topografix.com/GPX/1/1""");
				sr.WriteLine(@"xsi:schemaLocation=""http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd"">");
				//sr.WriteLine(@"<metadata>");
				//sr.WriteLine(@"<name>" + sDocName + @"</name>");
				//sr.WriteLine(@"<desc>" + sDocName + @"</desc>");
				//sr.WriteLine(@"<link href=""http://www.garmin.com"">");
				//sr.WriteLine(@"<text>Garmin International</text>");
				//sr.WriteLine(@"</link>");
				//sr.WriteLine(@"</metadata>");

				// Declare variables
				XPathNavigator nav2;

				// The waypoints must all be first then all the tracks
				// I'm sure there must be a clean way to select the "Placemarks"
				// depending on sub-nodes, but I'm lazy so we will just loop though
				// every node twice...

				// Loop thorough the placemarks but output only waypoints
				ni = nav.SelectDescendants("Placemark", "", false);
				while (ni.MoveNext()) // Through the placemarks (waypoints)
				{												
					sName = "";
					nav2 = ni.Current.Clone();
					nav2.MoveToFirstChild();				
					ProcessChild(ref nav2, ref sName, ref sr, true);
					while (nav2.MoveToNext())
					{
						ProcessChild(ref nav2, ref sName, ref sr, true);
					}	
				}				

				// Now Loop thorough the placemarks again and output only tracks
				nav.MoveToRoot(); // Start from top
				ni = nav.SelectDescendants("Placemark", "", false);
				while (ni.MoveNext()) // Through the placemarks (tracks)
				{												
					sName = "";
					nav2 = ni.Current.Clone();
					nav2.MoveToFirstChild();				
					ProcessChild(ref nav2, ref sName, ref sr, false);
					while (nav2.MoveToNext())
					{
						ProcessChild(ref nav2, ref sName, ref sr, false);
					}	
				}				
				
				//sr.WriteLine(@"<extensions>");
				//sr.WriteLine(@"</extensions>");
				sr.WriteLine(@"</gpx>");
				

				sr.Close();
				return true;
				
			}
			catch (System.Exception ex)
			{
				MessageBox.Show (ex.Message);
				return false;
			}
			
		}


		private void ProcessChild(ref XPathNavigator nav, ref string sName, ref StreamWriter sr, bool bDoWaypoints)
		{

			if (nav.Name == "name")
			{
				sName = nav.Value;
				sName = sName.Replace(@"&", @"&amp;");
				sName = sName.Replace(@"<", @"&lt;");
				sName = sName.Replace(@">", @"&gt;");
				sName = sName.Replace(@"'", @"&apos;");
				sName = sName.Replace(@"""", @"&quot;");

			}
			else if (nav.Name == "Point" && bDoWaypoints)
			{
//				<wpt lat="33.540270000" lon="-112.022800000">
//				<ele>33.832800</ele>
//				<name>PIESTEWA P</name>
//				<sym>Trail Head</sym>
//				</wpt>

				// Select the coordinates
				XPathNodeIterator ni = nav.SelectDescendants("coordinates", "", false);
				if (ni.MoveNext())
				{												
					XPathNavigator nav2 = ni.Current.Clone();
					nav2.MoveToFirstChild();				
			
					string sLat = nav2.Value.Split(",".ToCharArray())[1];;
					string sLon = nav2.Value.Split(",".ToCharArray())[0];;
					string sEle = nav2.Value.Split(",".ToCharArray())[2];;
					
					sr.WriteLine(@"<wpt lat=""" + sLat + @""" lon=""" + sLon + @""">");
					if(sEle != "0")
					{
						sr.WriteLine(@"<ele>" + sEle + @"</ele>");
					}
					sr.WriteLine(@"<name>" + sName + @"</name>");
					sr.WriteLine(@"</wpt>");
				}
			}
			else if ((nav.Name == "MultiGeometry" || nav.Name == "LineString") && !bDoWaypoints)
			{
//				<trk>
//				<name>Piestewa Peak Circumference</name>
//				<number>1</number>
//				<trkseg>
//				<trkpt lat="33.543250000" lon="-112.025140000">
//					<ele>131.06399755026854</ele>
//					<sym>Waypoint</sym>
//				</trkpt>
//				<trkpt lat="33.543440000" lon="-112.025950000"></trkpt>
//				</trkseg>
//				</trk>				


				// Select the coordinates collection		

				XPathNodeIterator xni = nav.SelectDescendants("coordinates", "", false);
				xni.MoveNext();
				XPathNavigator xnav2 = xni.Current;

				// Display the content of each element node.
				XPathNodeIterator xni2 = xnav2.SelectDescendants(XPathNodeType.Text, false); 
				while (xni2.MoveNext())
				{
					Console.WriteLine(xni2.Current.Name + " " + xni2.Current.Value);

					sr.WriteLine(@"<rte>");
					sr.WriteLine(@"<name>" + sName + "</name>");
					//sr.WriteLine(@"<trkseg>");

					// Get the points array
					string[] aPoints = xni2.Current.Value.Trim().Split(" ".ToCharArray());

					// Process the points
					foreach(string sPoint in aPoints)
					{
						string sLat = sPoint.Split(",".ToCharArray())[1];
						string sLon = sPoint.Split(",".ToCharArray())[0];
						string sEle = sPoint.Split(",".ToCharArray())[2];

						if(sEle == "0" || sEle == "")
						{
							sr.WriteLine(@"<rtept lat=""" + sLat + @""" lon=""" + sLon + @"""></rtept>");
						}
						else
						{
							sr.WriteLine(@"<rtept lat=""" + sLat + @""" lon=""" + sLon + @""">");
							sr.WriteLine(@"<ele>" + sEle + @"</ele>");
							sr.WriteLine(@"</rtept>");
						}
					}
					//sr.WriteLine(@"</trkseg>");
					sr.WriteLine(@"</rte>");
				}
			}
		}


		private string NullToNullString(object input)
		{
			if(input == null)
			{
				return "";
			}
			else
			{
				return input.ToString();
			}

		}

		
		private void GetSettings()
		{
			
			try
			{
				RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Derek\KML2GPX", true);		
				if(key == null)
				{
					Registry.CurrentUser.CreateSubKey(@"Software\Derek\KML2GPX");
				}
				else
				{
					this.txtFileName.Text = NullToNullString(key.GetValue("FileName"));
					this.txtOutputFolder.Text = NullToNullString(key.GetValue("OutputFolder"));				
				}
						
			}
			catch
			{
				MessageBox.Show("Failed to read settings");
			}
		}


		private void SaveSettings()
		{

			try
			{
				// Save settings to registry
				RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Derek\KML2GPX", true);		
				if(key == null)
				{
					Registry.CurrentUser.CreateSubKey(@"Software\Derek\KML2GPX");
				}

				key.SetValue("FileName", this.txtFileName.Text);
				key.SetValue("OutputFolder", this.txtOutputFolder.Text);

			}
			catch
			{
				MessageBox.Show("Failed to save settings.");
			}
		}

		
		private void btnBrowseOpenFile_Click(object sender, System.EventArgs e)
		{
			openFileDialog1.ShowDialog();
			txtFileName.Text = openFileDialog1.FileName;
		}

		
		private void btnBrowseFolder_Click(object sender, System.EventArgs e)
		{
			folderBrowserDialog1.ShowDialog();
			txtOutputFolder.Text = folderBrowserDialog1.SelectedPath;
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
		
			System.Reflection.AssemblyName an = Assembly.GetExecutingAssembly().GetName();
			string sVer = an.Version.Major.ToString() + "." + an.Version.Minor.ToString();
			string sVer2 = an.Version.Major.ToString() + "." + an.Version.Minor.ToString() + "." + an.Version.Revision.ToString();

			MessageBox.Show("Version: " + sVer2 + "\n\nThis utility is freeware and comes as is, with no support.","KML2GPX " + sVer + "  By Derek Rosen");
		
		}

	}
}

