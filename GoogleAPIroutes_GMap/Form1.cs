using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Net;
using Word = Microsoft.Office.Interop.Word;
using GMap.NET;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.Markers;

namespace GoogleAPIroutes_GMap
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            gMapControl1.Bearing = 0;
            gMapControl1.CanDragMap = true;
            gMapControl1.DragButton = MouseButtons.Left;
            gMapControl1.GrayScaleMode = true;
            gMapControl1.MarkersEnabled = true;
            gMapControl1.MaxZoom = 18;
            gMapControl1.MinZoom = 2;
            gMapControl1.MouseWheelZoomType = MouseWheelZoomType.MousePositionAndCenter;
            gMapControl1.NegativeMode = false;
            gMapControl1.PolygonsEnabled = true;
            gMapControl1.RoutesEnabled = true;
            gMapControl1.ShowTileGridLines = false;
            gMapControl1.Zoom = 2;
            gMapControl1.MapProvider = GMap.NET.MapProviders.GMapProviders.GoogleMap;
            GMaps.Instance.Mode = AccessMode.ServerOnly;
            GMap.NET.MapProviders.GMapProvider.WebProxy =
                WebRequest.GetSystemWebProxy();
            GMap.NET.MapProviders.GMapProvider.WebProxy.Credentials =
                CredentialCache.DefaultCredentials;
            dtRouter = new DataTable();
            dtRouter.Columns.Add("Шаг");
            dtRouter.Columns.Add("Нач. точка (latitude)");
            dtRouter.Columns.Add("Нач. точка (longitude)");
            dtRouter.Columns.Add("Кон. точка (latitude)");
            dtRouter.Columns.Add("Кон. точка (longitude)");
            dtRouter.Columns.Add("Время пути");
            dtRouter.Columns.Add("Расстояние");
            dtRouter.Columns.Add("Описание маршрута");          
            dataGridView1.DataSource = dtRouter;
            dataGridView1.Columns[7].Width = 250;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.ReadOnly = false;
            comboBox1.Items.Add("Автомобильные маршруты");
            comboBox1.Items.Add("Пешеходные маршруты");
            comboBox1.Items.Add("Велосипедные маршруты");
            comboBox1.Items.Add("Маршруты общественного транспорта");
            comboBox1.SelectedIndex = 0;
        }

        DataTable dtRouter;
        private void button1_Click(object sender, EventArgs e)
        {
            dtRouter.Rows.Clear();
            List<string> mode = new List<string>();
            mode.Add("driving");
            mode.Add("walking");
            mode.Add("bicycling");
            mode.Add("transit");

            string url = string.Format(
                "http://maps.googleapis.com/maps/api/directions/xml?origin={0},&destination={1}&sensor=false&language=ru&mode={2}",
                Uri.EscapeDataString(textBox1.Text), Uri.EscapeDataString(textBox2.Text), Uri.EscapeDataString(mode[comboBox1.SelectedIndex]));

            HttpWebRequest request =
                (HttpWebRequest)System.Net.WebRequest.Create(url);
            WebResponse response = request.GetResponse();
            Stream dataStream = response.GetResponseStream();
            StreamReader sreader = new System.IO.StreamReader(dataStream);     
            string responsereader = sreader.ReadToEnd();
            response.Close();

            System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();

            xmldoc.LoadXml(responsereader);

            if (xmldoc.GetElementsByTagName("status")[0].ChildNodes[0].InnerText == "OK")
            {   
                System.Xml.XmlNodeList nodes =
                    xmldoc.SelectNodes("//leg//step");

                object[] dr;
                for (int i = 0; i < nodes.Count; i++)
                {
                    dr = new object[8];
                    dr[0] = i;
                    dr[1] = xmldoc.SelectNodes("//start_location").Item(i).SelectNodes("lat").Item(0).InnerText.ToString();
                    dr[2] = xmldoc.SelectNodes("//start_location").Item(i).SelectNodes("lng").Item(0).InnerText.ToString();
                    dr[3] = xmldoc.SelectNodes("//end_location").Item(i).SelectNodes("lat").Item(0).InnerText.ToString();
                    dr[4] = xmldoc.SelectNodes("//end_location").Item(i).SelectNodes("lng").Item(0).InnerText.ToString();
                    dr[5] = xmldoc.SelectNodes("//duration").Item(i).SelectNodes("text").Item(0).InnerText.ToString();
                    dr[6] = xmldoc.SelectNodes("//distance").Item(i).SelectNodes("text").Item(0).InnerText.ToString();
                    dr[7] = HtmlToPlainText(xmldoc.SelectNodes("//html_instructions").Item(i).InnerText.ToString());
                    dtRouter.Rows.Add(dr);
                }

                textBox1.Text = xmldoc.SelectNodes("//leg//start_address").Item(0).InnerText.ToString();
                textBox2.Text = xmldoc.SelectNodes("//leg//end_address").Item(0).InnerText.ToString();

                double latStart = 0.0;
                double lngStart = 0.0;
                double latEnd = 0.0;
                double lngEnd = 0.0;

                latStart = System.Xml.XmlConvert.ToDouble(xmldoc.GetElementsByTagName("start_location")[nodes.Count].ChildNodes[0].InnerText);
                lngStart = System.Xml.XmlConvert.ToDouble(xmldoc.GetElementsByTagName("start_location")[nodes.Count].ChildNodes[1].InnerText);
         
                latEnd = System.Xml.XmlConvert.ToDouble(xmldoc.GetElementsByTagName("end_location")[nodes.Count].ChildNodes[0].InnerText);
                lngEnd = System.Xml.XmlConvert.ToDouble(xmldoc.GetElementsByTagName("end_location")[nodes.Count].ChildNodes[1].InnerText);

                dataGridView1.DataSource = dtRouter;
                gMapControl1.Position = new GMap.NET.PointLatLng(latStart, lngStart); 
                GMapOverlay markersOverlay = new GMapOverlay("marker");
                GMarkerGoogle markerG =
                    new GMarkerGoogle(new PointLatLng(latStart, lngStart), GMarkerGoogleType.green_pushpin);
                markerG.ToolTip =
                    new GMap.NET.WindowsForms.ToolTips.GMapRoundedToolTip(markerG);

                markerG.ToolTipMode = MarkerTooltipMode.Always;

                string[] wordsG = textBox1.Text.Split(',');
                string dataMarkerG = string.Empty;
                foreach (string word in wordsG)
                {
                    dataMarkerG += word + ";\n";
                }              
                markerG.ToolTipText = dataMarkerG;
                GMarkerGoogle markerR =
                    new GMarkerGoogle(
                    new PointLatLng(latEnd, lngEnd), GMarkerGoogleType.red_pushpin);
                markerG.ToolTip =
                    new GMap.NET.WindowsForms.ToolTips.GMapRoundedToolTip(markerG);
                markerR.ToolTipMode = MarkerTooltipMode.Always;
                string[] wordsR = textBox2.Text.Split(',');
                string dataMarkerR = string.Empty;
                foreach (string word in wordsR)
                {
                    dataMarkerR += word + ";\n";
                }
       
                markerR.ToolTipText = dataMarkerR;
                markersOverlay.Markers.Add(markerG);
                markersOverlay.Markers.Add(markerR);
                gMapControl1.Overlays.Clear();               
                List<PointLatLng> list = new List<PointLatLng>();
                for (int i = 0; i < dtRouter.Rows.Count; i++)
                {                    
                    double dbStartLat = double.Parse(dtRouter.Rows[i].ItemArray[1].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                    double dbStartLng = double.Parse(dtRouter.Rows[i].ItemArray[2].ToString(), System.Globalization.CultureInfo.InvariantCulture);

                    list.Add(new PointLatLng(dbStartLat, dbStartLng));

                    double dbEndLat = double.Parse(dtRouter.Rows[i].ItemArray[3].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                    double dbEndLng = double.Parse(dtRouter.Rows[i].ItemArray[4].ToString(), System.Globalization.CultureInfo.InvariantCulture);

                    list.Add(new PointLatLng(dbEndLat, dbEndLng));
                }

                markersOverlay.Routes.Clear();
                GMapRoute r = new GMapRoute(list, "Route");
                r.IsVisible = true;
                r.Stroke.Color = Color.DarkRed;
                markersOverlay.Routes.Add(r);
                gMapControl1.Overlays.Add(markersOverlay);
                gMapControl1.Zoom = 15;
                gMapControl1.Refresh();

                


            }
        }
        public string HtmlToPlainText(string html)
       {        
            html = html.Replace("</b>","");
            return html.Replace("<b>", "");
        }

        private void button2_Click(object sender, EventArgs e)
        {
                //Делаем скриншот и отправляем его в Word
               Bitmap bmp = new Bitmap(pictureBox1.Width, pictureBox1.Height);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(this.Location.X+20, this.Location.Y+150, 0, 0, new Size(pictureBox1.Width, pictureBox1.Height));
                }
                pictureBox1.Image = bmp;

                Word.Application application = new Word.Application();
                Object missing = Type.Missing;
                application.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                Clipboard.SetImage(bmp);
                application.ActiveDocument.Paragraphs[1].Range.Paste();
                application.Visible = true;
           
        }
    }
}
