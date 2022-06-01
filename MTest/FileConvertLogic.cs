using ClosedXML.Excel;
using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;

namespace MTest {
    class FileConvertLogic {
        public static FileConvertLogic Instance;
        XmlDocument doc = new XmlDocument();
        bool isPhotosNumberLowerThan2 = false;
        bool isMarkupLowerThan20Percent = false;
        int row = 2, column = 1;
        public static FileConvertLogic GetInstance() {
            if (Instance == null) {
                Instance = new FileConvertLogic();
            }
            return Instance;
        }
        public void CreateFile(string path) {
            XmlNodeList nodes = doc.SelectNodes("/produkty/produkt");
            IXLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("1");
            IXLWorksheet ws2 = wb.Worksheets.Add("2");

            ws.Cell(1, 1).Value = "id";
            ws.Cell(1, 2).Value = "nazwa";
            ws.Cell(1, 3).Value = "dlugi_opis";
            ws.Cell(1, 4).Value = "waga";
            ws.Cell(1, 5).Value = "kod";
            ws.Cell(1, 6).Value = "ean";
            ws.Cell(1, 7).Value = "status";
            ws.Cell(1, 8).Value = "typ";
            ws.Cell(1, 9).Value = "cena_zewnetrzna_hurt";
            ws.Cell(1, 10).Value = "cena_zewnetrzna";
            ws.Cell(1, 11).Value = "ilosc_wariantow";
            ws.Cell(1, 12).Value = "ilosc_zdjec";
            ws.Cell(1, 13).Value = "marża";

            ws.Column("B").Width = 45;
            ws.Column("C").Width = 150;
            ws.Column("F").Width = 20;
            ws.Column("C").Style.Alignment.WrapText = true;
            ws.Column("F").Style.NumberFormat.SetNumberFormatId((int)XLPredefinedFormat.Number.Integer);

            ws2.Cell(1, 1).Value = "id";
            ws2.Cell(1, 2).Value = "zdjecia";
            ws2.Column("B").Style.Alignment.WrapText = true;
            ws2.Column("B").Width = 100;

            foreach (XmlNode node in nodes) {
                string id = node.ChildNodes[0]?.InnerText;
                string nazwa = node.ChildNodes[1]?.InnerText;
                string dlugi_opis = node.ChildNodes[2]?.InnerText;
                string waga = node.ChildNodes[4]?.InnerText;
                string kod = node.ChildNodes[6]?.InnerText;
                string ean = node.ChildNodes[7]?.InnerText;
                string status = node.ChildNodes[8]?.InnerText;
                string typ = node.ChildNodes[9]?.InnerText;
                string cena_zewnetrzna_hurt = node.ChildNodes[10]?.InnerText;
                string cena_zewnetrzna = node.ChildNodes[11]?.InnerText;
                string ilosc_wariantow = node.ChildNodes[13]?.InnerText;

                while (column < 14) {
                    switch (column) {
                        case 1:
                            ws.Cell(row, column).Value = id;
                            ws2.Cell(row, 1).Value = id;
                            break;
                        case 2:
                            ws.Cell(row, column).Value = nazwa;
                            break;
                        case 3:
                            dlugi_opis = Regex.Replace(dlugi_opis, "<.*?>", String.Empty);
                            ws.Cell(row, column).Value = dlugi_opis;
                            break;
                        case 4:
                            ws.Cell(row, column).Value = waga;
                            break;
                        case 5:
                            ws.Cell(row, column).Value = kod;
                            break;
                        case 6:
                            ws.Cell(row, column).Value = ean;
                            break;
                        case 7:
                            ws.Cell(row, column).Value = status;
                            break;
                        case 8:
                            ws.Cell(row, column).Value = typ;
                            break;
                        case 9:
                            ws.Cell(row, column).Value = cena_zewnetrzna_hurt;
                            break;
                        case 10:
                            ws.Cell(row, column).Value = cena_zewnetrzna;
                            break;
                        case 11:
                            ws.Cell(row, column).Value = ilosc_wariantow;
                            break;
                        case 12:
                            XmlNode photosNode = node.ChildNodes[5];
                            ws.Cell(row, column).Value = photosNode.ChildNodes.Count;
                            if (photosNode.ChildNodes.Count < 2) {
                                ws.Row(row).Style.Fill.BackgroundColor = XLColor.Red;
                                isPhotosNumberLowerThan2 = true;
                            }
                            string photosString = "";
                            if (photosNode.HasChildNodes) {
                                for (int i = 0; i < photosNode.ChildNodes.Count; i++) {
                                    photosString += ($"{photosNode.ChildNodes[i].InnerText}\n");
                                }
                            }
                            ws2.Cell(row, 2).Value = photosString;
                            break;
                        case 13:
                            float cena_zewnetrzna_hurtFloat = float.Parse(cena_zewnetrzna_hurt, CultureInfo.InvariantCulture);
                            float cena_zewnetrznaFloat = float.Parse(cena_zewnetrzna, CultureInfo.InvariantCulture);
                            var sum = cena_zewnetrznaFloat - cena_zewnetrzna_hurtFloat;
                            float percentage = (sum / cena_zewnetrznaFloat) * 100;
                            ws.Cell(row, column).Value = sum;
                            if (percentage <= 20) {
                                ws.Row(row).Style.Fill.BackgroundColor = XLColor.Orange;
                                isMarkupLowerThan20Percent = true;
                            }
                            break;
                    }
                    column++;
                }
                var TwoBoolConditions = isMarkupLowerThan20Percent && isPhotosNumberLowerThan2;
                if (TwoBoolConditions)
                    ws.Row(row).Style.Fill.BackgroundColor = XLColor.LightBlue;
                isMarkupLowerThan20Percent = false;
                isPhotosNumberLowerThan2 = false;
                row++;
                column = 1;
            }
            wb.SaveAs(path);
        }
        public void OpenFile(string path) {
            doc.Load(path);

        }
    }
}
