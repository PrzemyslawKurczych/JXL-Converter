using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Xml;

namespace JXL_Converter
{
    public partial class JXL : Form
    {
        public JXL()
        {
            InitializeComponent();
        }

        private void btEnd_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btConvert_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                InitialDirectory = "c:\\",
                Filter = "Trimlbe JXL|*.jxl"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string fileName = ofd.FileName;
                string filePath = Path.GetDirectoryName(fileName);
                string onlyfileName = ofd.SafeFileName;
                string PKTname;
                string subPKTname = "_REF1";
                string domiar2;
                string domiar1;
                string offset1 = "0";
                XmlDocument XmlDoc = new XmlDocument();
                try
                {
                    XmlDoc.Load(fileName);
                    int CountPoint = XmlDoc.GetElementsByTagName("Point").Count;
                    int CountGrid = XmlDoc.GetElementsByTagName("Grid").Count;
                    int CountVRS = -1;

                    XmlNodeList itemlist = XmlDoc.SelectNodes("//FieldBook//PointRecord//Method");
                    foreach (XmlNode item in itemlist)
                    {
                        if (item.InnerText == "FromBase")
                        {
                            CountVRS++;
                        }
                    }
                    int CountSum = CountVRS + CountGrid-CountPoint;
                    for (int i = 0; i < CountPoint; i++)
                    {
                        PKTname = XmlDoc.GetElementsByTagName("Name").Item(i).InnerText;
                        if (PKTname.Contains(subPKTname) == true)
                        {
                            //Obliczenie przyrostów współrzędnych i odległości między punktami *_REF1 i *_REF2 a punktem wcinanym
                            double deltaN = Convert.ToDouble(XmlDoc.GetElementsByTagName("North").Item(i-1 + (CountSum)).InnerText) - Convert.ToDouble(XmlDoc.GetElementsByTagName("North").Item(i + 1 + (CountSum)).InnerText);
                            double deltaE = Convert.ToDouble(XmlDoc.GetElementsByTagName("East").Item(i-1 + (CountSum)).InnerText) - Convert.ToDouble(XmlDoc.GetElementsByTagName("East").Item(i + 1 + (CountSum)).InnerText);
                            double deltaN2 = Math.Pow(deltaN, 2);
                            double deltaE2 = Math.Pow(deltaE, 2);
                            domiar1 = Math.Round(Math.Sqrt(deltaE2 + deltaN2), 4).ToString();
                            deltaN = Convert.ToDouble(XmlDoc.GetElementsByTagName("North").Item(i  + (CountSum)).InnerText) - Convert.ToDouble(XmlDoc.GetElementsByTagName("North").Item(i + 1 + (CountSum)).InnerText);
                            deltaE = Convert.ToDouble(XmlDoc.GetElementsByTagName("East").Item(i  + (CountSum)).InnerText) - Convert.ToDouble(XmlDoc.GetElementsByTagName("East").Item(i + 1 + (CountSum)).InnerText);
                            deltaN2 = Math.Pow(deltaN,2);
                            deltaE2 = Math.Pow(deltaE, 2);
                            domiar2 = Math.Round(Math.Sqrt(deltaE2 + deltaN2), 4).ToString();
                            double elevationPKT = Convert.ToDouble(XmlDoc.GetElementsByTagName("Elevation").Item(i + 1 + (CountSum)).InnerText);
                            double elevationREF1 = Convert.ToDouble(XmlDoc.GetElementsByTagName("Elevation").Item(i - 1 + (CountSum)).InnerText);
                            double elevationREF2 = Convert.ToDouble(XmlDoc.GetElementsByTagName("Elevation").Item(i + (CountSum)).InnerText);
                            if (elevationPKT != (elevationREF1+elevationREF2)/2)
                            {
                                offset1 = (elevationPKT - (elevationREF1 + elevationREF2) / 2).ToString();
                            }
                             //określenie nazwy punktu wcinanego, bez końcówki "_REF1"
                            string FindName = PKTname.Substring(0, PKTname.Length - 5);
                            XmlNodeList nodelist = XmlDoc.SelectNodes("//FieldBook//PointRecord");
                            foreach (XmlNode node in nodelist)
                            {
                                foreach (XmlNode child in node.ChildNodes)
                                {
                                    if (child.Name == "Name" && child.InnerText.Equals(FindName))
                                    {
                                        //wyszukanie odpowiedniego miejsca w xml-u i wprowadzanie danych niezbędnych do zmiany typu na wcięcie
                                        XmlNode EnteredData = null;
                                        XmlNode Method = null;
                                        XmlNode PointOne = null;
                                        XmlNode PointTwo = null;
                                        XmlNode DistanceOne = null;
                                        XmlNode DistanceTwo = null;
                                        XmlNode VerticalOffset = null;
                                        EnteredData = XmlDoc.CreateNode(XmlNodeType.Element, "EnteredData", "");
                                        node.AppendChild(EnteredData);
                                        Method = XmlDoc.CreateNode(XmlNodeType.Element, "Method", "");
                                        Method.InnerText = "DistanceDistanceIntersection";
                                        EnteredData.AppendChild(Method);
                                        PointOne = XmlDoc.CreateNode(XmlNodeType.Element, "PointOne", "");
                                        PointOne.InnerText = XmlDoc.GetElementsByTagName("Name").Item(i).InnerText;
                                        EnteredData.AppendChild(PointOne);
                                        PointTwo = XmlDoc.CreateNode(XmlNodeType.Element, "PointTwo", "");
                                        PointTwo.InnerText = XmlDoc.GetElementsByTagName("Name").Item(i + 1).InnerText;
                                        EnteredData.AppendChild(PointTwo);
                                        DistanceOne = XmlDoc.CreateNode(XmlNodeType.Element, "DistanceOne", "");
                                        DistanceOne.InnerText = domiar1;
                                        EnteredData.AppendChild(DistanceOne);
                                        DistanceTwo = XmlDoc.CreateNode(XmlNodeType.Element, "DistanceTwo", "");
                                        DistanceTwo.InnerText = domiar2;
                                        EnteredData.AppendChild(DistanceTwo);
                                        VerticalOffset = XmlDoc.CreateNode(XmlNodeType.Element, "VerticalOffset", "");
                                        VerticalOffset.InnerText = offset1;
                                        EnteredData.AppendChild(VerticalOffset);
                                    }
                                }
                                //zmiana metody z COGO, na wcięcie z dwóch punktów
                                foreach (XmlNode child2 in node.ChildNodes)
                                {
                                    if (child2.Name == "Method" && child2.InnerText.Contains("CogoCalculated"))
                                    {
                                        child2.InnerText = child2.InnerText.Replace("CogoCalculated", "DistanceDistanceIntersection");
                                    }
                                }
                            }
                        }
                    }
                   
                    string newfileName = "zmienione" + onlyfileName;
                    string fileSave = Path.Combine(filePath, newfileName);
                    if (File.Exists(fileSave) != true)
                    {
                        XmlDoc.Save(fileSave);
                        MessageBox.Show("Plik " + onlyfileName + " został dostosowany do importu do C-Geo.", "Konwersja pliku " + onlyfileName + " została zakończona pomyślnie.");
                    }
                    else
                    {
                        var result = MessageBox.Show("Plik " + newfileName + " już istnieje. Jeżeli chcesz go nadpisać naciśnij 'Tak', utwożyć kopię naciśnij 'nie'.", "Czy chcesz nadpisać plik " + newfileName, MessageBoxButtons.YesNoCancel);
                        if (result == DialogResult.Yes)
                        {
                            XmlDoc.Save(fileSave);
                            MessageBox.Show("Plik " + onlyfileName + " został dostosowany do importu do C-Geo.", "Konwersja pliku " + onlyfileName + " została zakończona pomyślnie.");
                        }
                        else if (result == DialogResult.No)
                        {
                            newfileName = "kopia" + newfileName;
                            fileSave = Path.Combine(filePath, newfileName);
                            XmlDoc.Save(fileSave);
                            MessageBox.Show("Plik " + onlyfileName + " został dostosowany do importu do C-Geo.", "Konwersja pliku " + onlyfileName + " została zakończona pomyślnie.");
                        }
                        else
                        {
                            MessageBox.Show("Plik " + onlyfileName + " nie został dostosowany do importu do C-Geo.", "Konwersja pliku " + onlyfileName + " została zakończona niepowodzeniem.");
                        }
                    }
                    
                }

                catch (XmlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}

