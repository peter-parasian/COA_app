using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using WpfApp1.Core.Models;
using WpfApp1.Shared.Helpers;
using WpfApp1.ViewModels;

namespace WpfApp1.Core.Services
{
    public class CoaPrintService3
    {
        private const double DPI_VAL = 96.0;
        public async Task<string> GenerateCoaExcel(IList<WireSheetModel> sheets, bool defaultUseOhmPerMeter = false)
        {
            return await Task.Run<string>(() =>
            {
                if (sheets == null || sheets.Count == 0)
                {
                    throw new ArgumentException("Tidak ada sheet yang diberikan untuk generate COA.", nameof(sheets));
                }

                bool hasAnyItems = false;
                foreach (WireSheetModel sheet in sheets)
                {
                    if (sheet.Items != null && sheet.Items.Count > 0)
                    {
                        hasAnyItems = true;
                        break;
                    }
                }

                if (!hasAnyItems)
                {
                    throw new InvalidOperationException("Semua sheet kosong. Tidak ada data untuk di-export.");
                }

                WireExportItem? firstItem = sheets.FirstOrDefault(s => s.Items.Count > 0)?.Items[0];

                if (firstItem == null)
                {
                    throw new InvalidOperationException("Gagal membaca data item pertama.");
                }

                string customerName = firstItem.CustomerName;
                string size = firstItem.Specification;

                // 3. Persiapan Folder & Nama File
                string basePath = @"C:\Users\mrrx\Documents\My Web Sites\H\COA";
                DateTime now = DateTime.Now;
                string yearFolder = now.ToString("yyyy");
                string yearPath = Path.Combine(basePath, yearFolder);

                if (!Directory.Exists(yearPath))
                {
                    Directory.CreateDirectory(yearPath);
                }

                string finalMonthFolderName = string.Empty;
                bool folderFound = false;
                try
                {
                    string[] existingDirectories = Directory.GetDirectories(yearPath);
                    foreach (string dirPath in existingDirectories)
                    {
                        string dirName = Path.GetFileName(dirPath);
                        string leadingDigits = string.Empty;
                        foreach (char c in dirName)
                        {
                            if (char.IsDigit(c)) leadingDigits += c;
                            else break;
                        }

                        if (int.TryParse(leadingDigits, out int folderNum) && folderNum == now.Month)
                        {
                            finalMonthFolderName = dirName;
                            folderFound = true;
                            break;
                        }
                    }
                }
                catch { }

                if (!folderFound)
                {
                    var cultureIndo = new System.Globalization.CultureInfo("id-ID");
                    string monthName = cultureIndo.DateTimeFormat.GetMonthName(now.Month);
                    monthName = cultureIndo.TextInfo.ToTitleCase(monthName);
                    finalMonthFolderName = $"{now.Month}. {monthName}";
                }

                string finalDirectory = Path.Combine(yearPath, finalMonthFolderName);
                if (!Directory.Exists(finalDirectory))
                {
                    Directory.CreateDirectory(finalDirectory);
                }

                string[] existingFiles = Directory.GetFiles(finalDirectory, "*.xlsx");
                int validFileCount = 0;
                foreach (string f in existingFiles)
                {
                    if (!Path.GetFileName(f).StartsWith("~$")) validFileCount++;
                }

                int nextDocNum = validFileCount + 1;
                string docNumberSeq = nextDocNum.ToString("000");

                string fullDocNumberString = $"{docNumberSeq}/IWPI/QC/{now:MM}/{now:yyyy}";
                string safeCustomerName = customerName.Replace(".", "").Replace("/", "-");
                string fileName = $"{docNumberSeq}. COA Wire {safeCustomerName} {size}.xlsx";
                string fullSavePath = Path.Combine(finalDirectory, fileName);

                // 4. Dapatkan Konfigurasi Template & Gambar
                WireTemplateConfig templateConfig = GetTemplateConfiguration(customerName, size);
                Assembly assembly = Assembly.GetExecutingAssembly();

                // 5. Proses Pembuatan Excel
                using (Stream? templateStream = assembly.GetManifestResourceStream(templateConfig.ResourceName))
                {
                    if (templateStream == null)
                    {
                        throw new FileNotFoundException($"Template Embedded Resource tidak ditemukan: '{templateConfig.ResourceName}'. Pastikan Build Action file Excel diset ke 'Embedded Resource'.");
                    }

                    using (XLWorkbook workbook = new XLWorkbook(templateStream))
                    {
                        IXLWorksheet templateSheet = workbook.Worksheet(1);

                        List<ClosedXML.Excel.Drawings.IXLPicture> existingPics = templateSheet.Pictures.ToList();
                        foreach (ClosedXML.Excel.Drawings.IXLPicture pic in existingPics)
                        {
                            pic.Delete();
                        }

                        if (sheets.Count > 0)
                        {
                            templateSheet.Name = sheets[0].SheetName;
                        }

                        for (int i = 1; i < sheets.Count; i++)
                        {
                            templateSheet.CopyTo(sheets[i].SheetName);
                        }

                        for (int i = 0; i < sheets.Count; i++)
                        {
                            IXLWorksheet ws = workbook.Worksheet(i + 1);
                            WireSheetModel model = sheets[i];

                            if (!string.IsNullOrEmpty(templateConfig.CellDocNumber))
                            {
                                ws.Cell(templateConfig.CellDocNumber).Value = fullDocNumberString;
                            }

                            if (!string.IsNullOrEmpty(templateConfig.CellDate))
                            {
                                ws.Cell(templateConfig.CellDate).Value = now.ToString("dd-MM-yyyy");
                            }

                            if (templateConfig.Images != null && templateConfig.Images.Count > 0)
                            {
                                foreach (TemplateImageConfig imgConfig in templateConfig.Images)
                                {
                                    if (string.IsNullOrEmpty(imgConfig.ResourceName)) continue;

                                    using (Stream? imgStream = assembly.GetManifestResourceStream(imgConfig.ResourceName))
                                    {
                                        if (imgStream != null)
                                        {
                                            using (MemoryStream msRaw = new MemoryStream())
                                            {
                                                imgStream.CopyTo(msRaw);

                                                byte[] resizedBytes = ImageHelper.CompressImage(msRaw.ToArray(), imgConfig.WidthCm, imgConfig.HeightCm);

                                                using (MemoryStream msResized = new MemoryStream(resizedBytes))
                                                {
                                                    ClosedXML.Excel.Drawings.IXLPicture pic = ws.AddPicture(msResized);

                                                    pic.MoveTo(ws.Cell(imgConfig.Row, imgConfig.Col), imgConfig.OffsetX, imgConfig.OffsetY);

                                                    if (imgConfig.WidthCm > 0)
                                                    {
                                                        pic.Width = (int)((imgConfig.WidthCm / 2.54) * DPI_VAL);
                                                    }

                                                    if (imgConfig.HeightCm > 0)
                                                    {
                                                        pic.Height = (int)((imgConfig.HeightCm / 2.54) * DPI_VAL);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            ProcessDataRows(ws, model.Items, templateConfig, customerName, size, defaultUseOhmPerMeter);
                        }

                        workbook.SaveAs(fullSavePath);
                    }
                }

                FileInfo fileInfo = new FileInfo(fullSavePath);
                if (!fileInfo.Exists || fileInfo.Length == 0)
                {
                    throw new IOException($"File tidak berhasil dibuat atau kosong: {fullSavePath}");
                }

                return fullSavePath;
            });
        }

        private void ProcessDataRows(IXLWorksheet ws, IList<WireExportItem> items, WireTemplateConfig config, string customer, string size, bool defaultUseOhmPerMeter)
        {
            if (items == null || items.Count == 0) return;

            ws.ConditionalFormats.RemoveAll();

            ws.ShowGridLines = false;
            int currentRow = config.StartRow;

            bool useOhmPerMeter = config.UseOhmPerMeter.HasValue ? config.UseOhmPerMeter.Value : defaultUseOhmPerMeter;

            foreach (WireExportItem item in items)
            {
                WireRecord rec = item.RecordData;

                // Basic Data
                if (!string.IsNullOrEmpty(config.ColLot))
                {
                    ws.Cell($"{config.ColLot}{currentRow}").Value = rec.Lot;
                }

                if (!string.IsNullOrEmpty(config.ColDate))
                {
                    string formattedDate = rec.Date.Replace("/", "-");
                    ws.Cell($"{config.ColDate}{currentRow}").Value = formattedDate;
                }

                // Diameter
                if (!string.IsNullOrEmpty(config.ColDiameter))
                {
                    IXLCell cell = ws.Cell($"{config.ColDiameter}{currentRow}");
                    cell.Value = Math.Round(rec.Diameter, 2);
                    if (WireValidationHelper.ShouldHighlightDiameter(customer, size, rec.Diameter))
                    {
                        SetErrorStyle(cell);
                    }
                }

                // Elongation
                if (!string.IsNullOrEmpty(config.ColElongation))
                {
                    IXLCell cell = ws.Cell($"{config.ColElongation}{currentRow}");
                    cell.Value = Math.Round(rec.Elongation, 2);
                    if (WireValidationHelper.ShouldHighlightElongation(customer, size, rec.Elongation))
                    {
                        SetErrorStyle(cell);
                    }
                }

                // Tensile Strength (Kg/mm2) 
                if (!string.IsNullOrEmpty(config.ColTensileKg))
                {
                    double tensileKg = MathHelper.CalculateTensileStrengthKgmm2(rec.Tensile);
                    IXLCell cell = ws.Cell($"{config.ColTensileKg}{currentRow}");
                    cell.Value = Math.Round(tensileKg, 2);

                    if (WireValidationHelper.ShouldHighlightTensile(customer, size, rec.Tensile))
                    {
                        SetErrorStyle(cell);
                    }
                }

                // Tensile Strength (N/mm2) 
                if (!string.IsNullOrEmpty(config.ColTensileN))
                {
                    IXLCell cell = ws.Cell($"{config.ColTensileN}{currentRow}");
                    cell.Value = Math.Round(rec.Tensile, 2);

                    if (WireValidationHelper.ShouldHighlightTensile(customer, size, rec.Tensile))
                    {
                        SetErrorStyle(cell);
                    }
                }

                // Yield Strength
                if (!string.IsNullOrEmpty(config.ColYield))
                {
                    IXLCell cell = ws.Cell($"{config.ColYield}{currentRow}");
                    cell.Value = Math.Round(rec.Yield, 2);

                    if (WireValidationHelper.ShouldHighlightYield(customer, size, rec.Yield))
                    {
                        SetErrorStyle(cell);
                    }
                }

                // Copper Content 
                if (!string.IsNullOrEmpty(config.ColCopper))
                {
                    IXLCell cell = ws.Cell($"{config.ColCopper}{currentRow}");
                    cell.Value = "";
                    cell.Style.Fill.BackgroundColor = XLColor.Red;
                }

                // Conductor Resistance Logic
                if (!string.IsNullOrEmpty(config.ColCondRes))
                {
                    double condRes = MathHelper.CalculateConductorResisten(rec.IACS, rec.Diameter);

                    if (!useOhmPerMeter)
                    {
                        condRes = condRes * 1000.0;
                    }

                    IXLCell cell = ws.Cell($"{config.ColCondRes}{currentRow}");
                    cell.Value = Math.Round(condRes, 2);

                    if (WireValidationHelper.ShouldHighlightConductorResistance(customer, size, rec.IACS, rec.Diameter))
                    {
                        SetErrorStyle(cell);
                    }
                }

                // IACS
                if (!string.IsNullOrEmpty(config.ColIACS))
                {
                    IXLCell cell = ws.Cell($"{config.ColIACS}{currentRow}");
                    cell.Value = Math.Round(rec.IACS, 2);
                    if (WireValidationHelper.ShouldHighlightIACS(customer, size, rec.IACS))
                    {
                        SetErrorStyle(cell);
                    }
                }

                // Electrical Resistivity
                if (!string.IsNullOrEmpty(config.ColElecRes))
                {
                    double elecRes = MathHelper.CalculateElectricalResistivity(rec.IACS);
                    IXLCell cell = ws.Cell($"{config.ColElecRes}{currentRow}");
                    cell.Value = Math.Round(elecRes, 5);

                    if (WireValidationHelper.ShouldHighlightElectricalResistivity(customer, size, rec.IACS))
                    {
                        SetErrorStyle(cell);
                    }
                }

                // Electrical Conductivity
                if (!string.IsNullOrEmpty(config.ColElecCond))
                {
                    double elecCond = MathHelper.CalculateElectricalConductivity(rec.IACS);
                    IXLCell cell = ws.Cell($"{config.ColElecCond}{currentRow}");
                    cell.Value = Math.Round(elecCond, 2);

                    if (WireValidationHelper.ShouldHighlightElectricalConductivity(customer, size, rec.IACS))
                    {
                        SetErrorStyle(cell);
                    }
                }

                // Remarks
                if (!string.IsNullOrEmpty(config.ColRemarks))
                {
                    ws.Cell($"{config.ColRemarks}{currentRow}").Value = "OK";
                }
                if (!string.IsNullOrEmpty(config.ColRemarks2))
                {
                    ws.Cell($"{config.ColRemarks2}{currentRow}").Value = "OK (HARD)";
                }
                if (!string.IsNullOrEmpty(config.ColRemarks3))
                {
                    ws.Cell($"{config.ColRemarks3}{currentRow}").Value = "No Scratch Surface\nNo Discoloration";
                }

                currentRow += config.RowStep;
            }
        }

        private void SetErrorStyle(IXLCell cell)
        {
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            // cell.Style.Font.Bold = true; 
            cell.Style.Font.FontColor = XLColor.Black;
        }

        private WireTemplateConfig GetTemplateConfiguration(string customer, string size)
        {
            string key = $"{customer.Trim()}|{size.Trim()}";
            var config = new WireTemplateConfig();
            //config.UseOhmPerMeter = true;

            switch (key)
            {
                case "Canning|1.20":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_120_Canning.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA-comp.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 14,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 12,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 18,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 22,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColCopper = "M";
                    config.ColCondRes = "O";
                    config.ColIACS = "Q";
                    config.ColRemarks = "S";
                    break;

                case "Indolakto|1.20":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_120_Indolakto.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 10,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 33,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;

                case "Multi Colour|1.20":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_120_Multi.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire2_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 29,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;

                case "Nestle|1.20":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_120_Nestle.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 20,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 10,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 23,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 20,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColYield = "J";
                    config.ColTensileN = "M";
                    config.ColElecCond = "S";
                    config.ColCopper = "U";
                    config.ColElecRes = "V";
                    config.ColRemarks = "X";
                    break;

                case "Canning|1.24":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_124_Canning.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 14,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 30,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 18,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 25,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColCopper = "M";
                    config.ColCondRes = "O";
                    config.ColIACS = "Q";
                    config.ColRemarks = "S";
                    break;

                case "Cometa|1.24":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_124_Cometa.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 13,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 67,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 17,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 8,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColCopper = "M";
                    config.ColElecRes = "N";
                    config.ColIACS = "O";
                    config.ColRemarks = "R";
                    break;

                case "Multi Colour|1.24":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_124_Multi.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 12,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire2_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 33,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;

                case "Almicos|1.38":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_138_Almicos.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 29,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;

                case "Avia Avian|1.38":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_138_Avia.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 43,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 7,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 43,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 28,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;

                case "Cometa|1.38":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_138_Cometa.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 14,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 17,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 17,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 29,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColCopper = "M";
                    config.ColCondRes = "N";
                    config.UseOhmPerMeter = true;
                    config.ColIACS = "O";
                    config.ColRemarks = "R";
                    break;

                case "Eka Timur|1.38":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_138_Eka.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 29,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;

                case "Energy Lautan|1.38":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_138_Lautan.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 29,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;


                case "Masami Pasifik|1.38":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_138_Masami.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 17,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 40,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;

                case "Metal Manufacturing|1.38":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_138_Metal.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 5,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 29,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;

                case "Multi Colour|1.38":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_138_Multi.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 10,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire2_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 33,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;

                case "Prisma Cable|1.38":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_138_Prisma.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 11,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 23,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 15,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 29,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "I";
                    config.ColCopper = "L";
                    config.ColIACS = "N";
                    config.ColRemarks = "P";
                    break;

                case "Cometa|1.50":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_150_Cometa.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 14,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 20,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 17,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 28,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColCopper = "M";
                    config.ColCondRes = "N";
                    config.UseOhmPerMeter = true;
                    config.ColIACS = "O";
                    config.ColRemarks = "R";
                    break;

                case "Avia Avian|1.50":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_150_Avia.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 43,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 8,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 43,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 29,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;

                case "Multi Colour|1.50":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_150_Multi.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 12,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire2_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 32,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;

                case "Indowire (Hard)|1.60":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_160_Indowire_HARD.xlsx";  
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 3,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 29,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks2 = "O";
                    break;

                case "Indowire (Soft)|1.60":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_160_Indowire_SOFT.xlsx";  
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 42,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 3,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 42,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 28,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "D";
                    config.ColElongation = "G";
                    config.ColTensileKg = "J";
                    config.ColIACS = "M";
                    config.ColRemarks2 = "O";
                    break;

                case "Magnakabel|1.60":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_160_Magnakabel.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 43,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 2,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 43,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 29,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColDiameter = "G";
                    config.ColElongation = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;

                case "Metal Manufacturing|1.60":
                    config.ResourceName = "WpfApp1.Shared.Images.Size_160_Magnakabel.xlsx";
                    config.StartRow = 15;
                    config.RowStep = 1;
                    config.CellDocNumber = "C6";
                    config.CellDate = "C7";
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.logo_COA.png",
                        Row = 1,
                        Col = 1,
                        WidthCm = 5.11,
                        HeightCm = 1.63,
                        OffsetX = 0,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.approved_IMG_v2.png",
                        Row = 43,
                        Col = 10,
                        WidthCm = 9.41,
                        HeightCm = 3.3,
                        OffsetX = 3,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.SoC_Free_HD.png",
                        Row = 43,
                        Col = 1,
                        WidthCm = 2.27,
                        HeightCm = 2.33,
                        OffsetX = 50,
                        OffsetY = 0
                    });
                    config.Images.Add(new TemplateImageConfig
                    {
                        ResourceName = "WpfApp1.Shared.Images.document_wire1_comp.png",
                        Row = 1,
                        Col = 14,
                        WidthCm = 3.58,
                        HeightCm = 1.52,
                        OffsetX = 28,
                        OffsetY = 0
                    });
                    config.ColLot = "A";
                    config.ColDate = "B";
                    config.ColRemarks3 = "D";
                    config.ColDiameter = "G";
                    config.ColElongation = "J";
                    config.ColIACS = "M";
                    config.ColRemarks = "O";
                    break;

                default:
                    return GetTemplateConfigurationFull(customer, size);
            }
            return config;
        }

        private WireTemplateConfig GetTemplateConfigurationFull(string customer, string size)
        {
            return new WireTemplateConfig();
        }

        public void ClearCache()
        {
        }

        private struct TemplateImageConfig
        {
            public string ResourceName;
            public int Row;
            public int Col;

            public double WidthCm;
            public double HeightCm;

            public int OffsetX;
            public int OffsetY;
        }

        private struct WireTemplateConfig
        {
            public string ResourceName;
            public int StartRow;
            public int RowStep;

            public string CellDocNumber;
            public string CellDate;

            public List<TemplateImageConfig> Images;

            public string ColLot;
            public string ColDate;
            public string ColDiameter;
            public string ColElongation;
            public string ColTensileKg;
            public string ColTensileN;
            public string ColYield;
            public string ColCopper;
            public string ColCondRes;
            public string ColIACS;
            public string? ColElecRes;
            public string? ColElecCond;
            public string ColRemarks;
            public string ColRemarks2;
            public string ColRemarks3;

            public bool? UseOhmPerMeter;

            public WireTemplateConfig()
            {
                ResourceName = ""; StartRow = 0; RowStep = 0;
                CellDocNumber = ""; CellDate = "";

                Images = new List<TemplateImageConfig>();

                ColLot = ""; ColDate = ""; ColDiameter = ""; ColElongation = "";
                ColTensileKg = ""; ColTensileN = ""; ColYield = "";
                ColCopper = ""; ColCondRes = ""; ColIACS = ""; ColRemarks2 = "";
                ColElecRes = null; ColElecCond = null; ColRemarks = "";
                ColRemarks3 = "";

                UseOhmPerMeter = null;
            }
        }
    }
}