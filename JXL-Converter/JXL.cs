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
        private double X1;
        private double Y1;
        private double H1;
        private string Ref1Name;
        private double X2;
        private double Y2;
        private double H2;
        private double X0;
        private double Y0;
        private double H0;
        private string Ref2Name;

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
                Filter = "Trimble JXL|*.jxl"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string fileName = ofd.FileName;
                string filePath = Path.GetDirectoryName(fileName);
                string onlyfileName = ofd.SafeFileName;
                string subPKTname = null;
                string domiar2;
                string domiar1;
                string DistanceTwoValue;
                string DistanceOneValue;
                string offset1 = "0";
                string PointOneName;
                string PointTwoName;
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
                    if (CountVRS != 0)
                    { CountVRS = 1; }

                    int CountSum = CountVRS + CountGrid - CountPoint;
                    XmlNodeList reductionslist = XmlDoc.SelectNodes("//Reductions//Point");
                    foreach (XmlNode reductions in reductionslist)
                    {
                        //  foreach (XmlNode point in reductions.ChildNodes) 
                        if (reductions.ChildNodes.Item(1).InnerText.Contains("_REF1"))
                        {
                            X1 = Convert.ToDouble(reductions.ChildNodes.Item(6).ChildNodes.Item(0).InnerText);
                            Y1 = Convert.ToDouble(reductions.ChildNodes.Item(6).ChildNodes.Item(1).InnerText);
                            H1 = Convert.ToDouble(reductions.ChildNodes.Item(6).ChildNodes.Item(2).InnerText);
                            Ref1Name = reductions.ChildNodes.Item(1).InnerText;
                        }
                        if (reductions.ChildNodes.Item(1).InnerText.Contains("_REF2"))
                        {
                            X2 = Convert.ToDouble(reductions.ChildNodes.Item(6).ChildNodes.Item(0).InnerText);
                            Y2 = Convert.ToDouble(reductions.ChildNodes.Item(6).ChildNodes.Item(1).InnerText);
                            H2 = Convert.ToDouble(reductions.ChildNodes.Item(6).ChildNodes.Item(2).InnerText);
                            subPKTname = reductions.ChildNodes.Item(1).InnerText.Substring(0, reductions.ChildNodes.Item(1).InnerText.Length - 5);
                            Ref2Name = reductions.ChildNodes.Item(1).InnerText;
                        }

                        if (subPKTname != null && reductions.ChildNodes.Item(1).InnerText == subPKTname)
                        {
                            X0 = Convert.ToDouble(reductions.ChildNodes.Item(6).ChildNodes.Item(0).InnerText);
                            Y0 = Convert.ToDouble(reductions.ChildNodes.Item(6).ChildNodes.Item(1).InnerText);
                            H0 = Convert.ToDouble(reductions.ChildNodes.Item(6).ChildNodes.Item(2).InnerText);
                        
                        //Obliczenie przyrostów współrzędnych i odległości między punktami *_REF1 i *_REF2 a punktem wcinanym

                        double deltaN = X0 - X1;
                        double deltaE = Y0 - Y1;
                        double deltaN2 = Math.Pow(deltaN, 2);
                        double deltaE2 = Math.Pow(deltaE, 2);
                        domiar1 = Math.Round(Math.Sqrt(deltaE2 + deltaN2), 4).ToString();
                        deltaN = X0 - X2;
                        deltaE = Y0 - Y2;
                        deltaN2 = Math.Pow(deltaN, 2);
                        deltaE2 = Math.Pow(deltaE, 2);
                        domiar2 = Math.Round(Math.Sqrt(deltaE2 + deltaN2), 4).ToString();
                        if (H0 != (H1 + H2) / 2)
                        {
                            offset1 = (H0 - (H1 + H2) / 2).ToString();
                        }
                            //sprawdzanie czy punkt jest na prawo od REF1 REF2, Kąt w radianach między Ref1 - Ref2 - Punkt musi być pomiędzy 0 - pi
                            double deltaN20 = X2 - X0;
                            double deltaE20 = Y2 - Y0;
                            double phi20 = Math.Atan2(deltaE20, deltaN20);
                            if (deltaN20 < 0 && deltaE20 > 0 )
                            {
                                phi20 = phi20 + Math.PI;
                            }
                            else if (deltaN20 <0 && deltaE20 <0)
                            {
                                phi20 = phi20 + Math.PI;
                            }
                            else if (deltaN20 > 0 && deltaE20 < 0)
                            {
                                phi20 = phi20 + 2*Math.PI;
                            }
                            double deltaN21 = X2 - X1;
                            double deltaE21 = Y2 - Y1;
                            double phi21 = Math.Atan2(deltaE21, deltaN21);
                            if (deltaN21 < 0 && deltaE21 > 0)
                            {
                                phi21 = phi21 + Math.PI;
                            }
                            else if (deltaN21 < 0 && deltaE21 < 0)
                            {
                                phi21 = phi21 + Math.PI;
                            }
                            else if (deltaN21 > 0 && deltaE21 < 0)
                            {
                                phi21 = phi21 + 2 * Math.PI;
                            }
                            double alpha120 = phi21 - phi20;
                            if (alpha120 < Math.PI)
                            {
                                PointOneName = Ref1Name;
                                PointTwoName = Ref2Name;
                                DistanceOneValue = domiar1;
                                DistanceTwoValue = domiar2;
                            }
                            else
                            {
                                PointOneName = Ref2Name;
                                PointTwoName = Ref1Name;
                                DistanceOneValue = domiar2;
                                DistanceTwoValue = domiar1;
                            }

                            XmlNodeList fieldbooklist = XmlDoc.SelectNodes("//FieldBook//PointRecord");
                            foreach (XmlNode pointrecord in fieldbooklist)
                            {
                                foreach (XmlNode child in pointrecord.ChildNodes)
                                {
                                    if (child.Name == "Name" && child.InnerText.Equals(subPKTname))
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
                                        pointrecord.AppendChild(EnteredData);
                                        Method = XmlDoc.CreateNode(XmlNodeType.Element, "Method", "");
                                        Method.InnerText = "DistanceDistanceIntersection";
                                        EnteredData.AppendChild(Method);
                                        PointOne = XmlDoc.CreateNode(XmlNodeType.Element, "PointOne", "");
                                        PointOne.InnerText = PointOneName;
                                        EnteredData.AppendChild(PointOne);
                                        PointTwo = XmlDoc.CreateNode(XmlNodeType.Element, "PointTwo", "");
                                        PointTwo.InnerText = PointTwoName;
                                        EnteredData.AppendChild(PointTwo);
                                        DistanceOne = XmlDoc.CreateNode(XmlNodeType.Element, "DistanceOne", "");
                                        DistanceOne.InnerText = DistanceOneValue;
                                        EnteredData.AppendChild(DistanceOne);
                                        DistanceTwo = XmlDoc.CreateNode(XmlNodeType.Element, "DistanceTwo", "");
                                        DistanceTwo.InnerText = DistanceTwoValue;
                                        EnteredData.AppendChild(DistanceTwo);
                                        VerticalOffset = XmlDoc.CreateNode(XmlNodeType.Element, "VerticalOffset", "");
                                        VerticalOffset.InnerText = offset1;
                                        EnteredData.AppendChild(VerticalOffset);
                                    }
                                }
                                //zmiana metody z COGO, na wcięcie z dwóch punktów
                                foreach (XmlNode child2 in pointrecord.ChildNodes)
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


