using ClosedXML.Excel;
using LevDan.Exif;
using System;
using System.Collections.Generic;
using System.Data;
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

        private void ButtonDestination_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                textBoxDestination.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void ButtonProcess_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            DataTable dtGPSData = new DataTable();

            dtGPSData.Columns.AddRange(new DataColumn[11]
            {
                new DataColumn("Filename", typeof(string)),
                new DataColumn("Latitude", typeof(double)),
                new DataColumn("Longitude", typeof(double)),
                new DataColumn("Altitude",typeof(string)),
                new DataColumn("Make", typeof(string)),
                new DataColumn("Model",typeof(string)),
                new DataColumn("ModifyDateTime", typeof(string)),
                new DataColumn("GPSTimeAtomicClock", typeof(string)),
                new DataColumn("DateTimeOriginal",typeof(string)),
                new DataColumn("DateTimeDigitized", typeof(string)),
                new DataColumn("GPSDateStamp", typeof(string))
            });

            double latitude = 0;
            double longitude = 0;
            string altitude = string.Empty;
            string modifyDateTime = string.Empty;
            string make = string.Empty;
            string model = string.Empty;
            string gpsTimeStamp = string.Empty;
            string dateTimeOriginal = string.Empty;
            string dateTimeDigitized = string.Empty;
            string gpsDateStamp = string.Empty;

            List<string> photos = new List<string>();
            DirectoryInfo source = new DirectoryInfo(textBoxSource.Text);
            DirectoryInfo destination = new DirectoryInfo(textBoxDestination.Text);

            foreach (FileInfo path in source.GetFiles("*", SearchOption.AllDirectories))
            {
                sb.Append(Path.GetFileName(path.FullName));

                try
                {
                    ExifGPSLatLonTagCollection exif = new ExifGPSLatLonTagCollection(path.FullName);

                    if (exif.Count() >= 3)
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
                                        modifyDateTime = tag.Value;
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
                                case "DateTimeOriginal":
                                    {
                                        dateTimeOriginal = tag.Value;
                                        break;
                                    }
                                case "DateTimeDigitized":
                                    {
                                        dateTimeDigitized = tag.Value;
                                        break;
                                    }
                                case "GPSDateStamp":
                                    {
                                        gpsDateStamp = tag.Value;
                                        break;
                                    }
                                case "GPSTimeStamp":
                                    {
                                        gpsTimeStamp = tag.Value;
                                        break;
                                    }
                            }
                        }

                        if (latitude > 0 && longitude > 0)
                        {
                            dtGPSData.Rows.Add(Path.GetFileName(path.FullName).ToString(), latitude.ToString(),
                                longitude.ToString(), altitude.ToString(), make.ToString(), model.ToString(),
                                modifyDateTime.ToString(), dateTimeOriginal.ToString(), dateTimeDigitized.ToString(),
                                gpsDateStamp.ToString(), gpsTimeStamp.ToString());

                            photos.Add(longitude + "," + latitude + "," + altitude + "," + modifyDateTime + "," +
                                Path.GetFileName(path.FullName) + "," + make + "," + model);
                        }

                        latitude = 0;
                        longitude = 0;
                        altitude = string.Empty;
                        modifyDateTime = string.Empty;
                        make = string.Empty;
                        model = string.Empty;
                    }
                }
                catch { }
            }

            if (photos.Count > 0)
            {
                string path3 = Path.Combine(destination.FullName, "Photo GPS Report.kml");
                KML.Create(photos, path3);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                string path = Path.Combine(destination.FullName, "Photo Location Report.xlsx");
                wb.Worksheets.Add(dtGPSData, "GPS");
                wb.SaveAs(path);
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
                    writer.WriteElementString("name", "gps_photo_report.xml");
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

                        writer.WriteStartElement("Placemark");
                        writer.WriteElementString("description", sb.ToString());
                        writer.WriteElementString("name", splitUp[4].ToString());
                        writer.WriteStartElement("Point");
                        writer.WriteElementString("coordinates", splitUp[0].ToString() + "," + splitUp[1].ToString());
                        writer.WriteEndElement();//Point
                        writer.WriteEndElement();//Placemark

                        sb.Clear();
                    }

                    writer.WriteEndElement();//Document
                    writer.Flush();
                }
            }
        }

    }
}
