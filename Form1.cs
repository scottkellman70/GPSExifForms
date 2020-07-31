using LevDan.Exif;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace GPSExifForms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void ButtonSource_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                textBoxSource.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void buttonDestination_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                textBoxDestination.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void buttonProcess_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            double latitude = 0;
            double longitude = 0;
            string altitude = string.Empty;
            string dateTime = string.Empty;
            string make = string.Empty;
            string model = string.Empty;
            List<string> photos = new List<string>();
            DirectoryInfo source = new DirectoryInfo(textBoxSource.Text);
            DirectoryInfo destination = new DirectoryInfo(textBoxDestination.Text);

            foreach (FileInfo path in source.GetFiles("*", SearchOption.AllDirectories))
            {
                sb.Append(Path.GetFileName(path.FullName));

                try
                {
                    ExifGPSLatLonTagCollection exif = new ExifGPSLatLonTagCollection(path.FullName);

                    if (exif.Count() >= 3)//datetime, lat, long
                    {
                        foreach (ExifTag tag in exif)
                        {
                            string latRef = string.Empty;
                            string lonRef = string.Empty;

                            foreach (ExifTag tag2 in exif)
                            {
                                switch (tag2.FieldName)
                                {
                                    case "GPSLatitudeRef":
                                        {
                                            latRef = tag2.Value;
                                            break;
                                        }
                                    case "GPSLongitudeRef":
                                        {
                                            lonRef = tag2.Value;
                                            break;
                                        }
                                }
                            }
                            switch (tag.FieldName)
                            {
                                case "GPSLatitude":
                                    {
                                        if (!string.IsNullOrEmpty(latRef))
                                        {
                                            latitude = Utilities.GPS.GetLatLonFromDMS(latRef.Substring(0, 1) + tag.Value);
                                        }
                                        latitude = Utilities.GPS.GetLatLonFromDMS(tag.Value);
                                        break;
                                    }
                                case "GPSLongitude":
                                    {
                                        if (!string.IsNullOrEmpty(lonRef))
                                        {
                                            longitude = Utilities.GPS.GetLatLonFromDMS(lonRef.Substring(0, 1) + tag.Value);
                                        }
                                        longitude = Utilities.GPS.GetLatLonFromDMS(tag.Value);
                                        break;
                                    }
                                case "GPSAltitude":
                                    {
                                        altitude = tag.Value;
                                        break;
                                    }
                                case "DateTime":
                                    {
                                        dateTime = tag.Value;
                                        break;
                                    }
                                case "Make":
                                    {
                                        make = tag.Value;
                                        break;
                                    }
                                case "Model":
                                    {
                                        model = tag.Value;
                                        break;
                                    }
                            }
                        }
                        sb.Append(Path.GetFileName(path.FullName) + "," + latitude + "," + longitude + "," + altitude + "," + make + "," + model);
                        sb.Append(Environment.NewLine);
                        photos.Add(longitude + "," + latitude + "," + altitude + "," + dateTime + "," + Path.GetFileName(path.FullName) + "," + make + "," + model);
                        latitude = 0;
                        longitude = 0;
                        altitude = string.Empty;
                        dateTime = string.Empty;
                        make = string.Empty;
                    }
                    else
                    {
                        sb.Clear();
                    }
                }
                catch (Exception) { }
            }

            if (photos.Count > 0)
            {
                string path3 = Path.Combine(@"C:\Users\OCONUS4\Desktop\Cases\Results", "photos.kml");
                KML.Create(photos, path3);

                string path2 = Path.Combine(@"C:\Users\OCONUS4\Desktop\Cases\Results", "report.txt");
                File.WriteAllText(path2, sb.ToString());
            }
            else
            {
                MessageBox.Show("No locational data in the selected photo set.");
            }

            MessageBox.Show("Done processing files");
        }
        public static class KML
        {
            public static void Create(List<string> points, string path)
            {
                StringBuilder sb = new StringBuilder();

                using (XmlWriter writer = XmlWriter.Create(path))
                {
                    writer.WriteStartElement("Document");
                    writer.WriteElementString("name", "photos.xml");
                    writer.WriteElementString("open", "1");

                    writer.WriteStartElement("Style");
                    writer.WriteStartElement("LabelStyle");
                    writer.WriteElementString("color", "ff0000cc");
                    writer.WriteEndElement();//LabelStyle
                    writer.WriteEndElement();//Style

                    foreach (string p in points)
                    {
                        string[] splitUp = p.Split(',');
                        sb.Append("Latitude: " + splitUp[0].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Longitude: " + splitUp[1].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Altitude: " + splitUp[2].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Date Time: " + splitUp[3].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("File Name: " + splitUp[4].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Make: " + splitUp[5].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Model: " + splitUp[6].ToString());

                        if (splitUp[0].ToString() != "0" && splitUp[1].ToString() != "0")
                        {
                            writer.WriteStartElement("Placemark");
                            writer.WriteElementString("description", sb.ToString());
                            writer.WriteElementString("name", splitUp[4].ToString());
                            writer.WriteStartElement("Point");
                            writer.WriteElementString("coordinates", splitUp[0].ToString() + "," + splitUp[1].ToString());
                            writer.WriteEndElement();//Point
                            writer.WriteEndElement();//Placemark
                        }
                        sb.Clear();
                    }

                    writer.WriteEndElement();//Document
                    writer.Flush();
                }
            }
        }
    }
}
