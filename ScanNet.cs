using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using WIA;

/*
╔═══════════════════════════════╗
║    © JRZ Soft Mx | ScanNet    ║
╚═══════════════════════════════╝
*/

namespace ScanNet
{
    public class ScanNet
    {
        //Class Scan
        #region
        public class ScanCfg
        {
            public int? IntColor_Mode { get; set; }
            public int? IntResolution_DPI_Horizontal { get; set; }
            public int? IntResolution_DPI_Vertical { get; set; }
            public int? IntStart_Pixel_Horizontal { get; set; }
            public int? IntStart_Pixel_Vertical { get; set; }
            public int? IntScan_Size_Pixels_Horizontal { get; set; }
            public int? IntScan_Size_Pixels_Vertical { get; set; }
            public int? IntScan_Brightness_Percents { get; set; }
            public int? IntScan_Contrast_Percents { get; set; }
        }
        #endregion

        //Scan
        #region
        public static void Scan(string strScanner, string strFilePath, string strImageFormat, bool bolPdf = false, ScanCfg scConfig = null)
        {
            try
            {
                DeviceManager dmScanners = new DeviceManager();
                Device Scanner = null;
                List<string> lisScanners = new List<string>();
                for (int i = 1; i <= dmScanners.DeviceInfos.Count; i++)
                {
                    if (dmScanners.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType)
                    {
                        continue;
                    }
                    string strScan = dmScanners.DeviceInfos[i].Properties["Name"].get_Value().ToString();
                    if (strScan == strScanner)
                    {
                        Scanner = dmScanners.DeviceInfos[i].Connect();
                    }
                    else
                    {
                        lisScanners.Add(strScan);
                    }
                }
                if (string.IsNullOrEmpty(strScanner))
                {
                    throw new Exception("The scanner is not specify");

                }
                if (Scanner == null)
                {
                    string strScanners = lisScanners.Count > 0 ? " scanners aviables " + string.Join(",", lisScanners.ToArray()) : " no scanners aviables";
                    throw new Exception("The scanner is not found" + strScanners);
                }
                if (!Directory.Exists(Path.GetDirectoryName(strFilePath)))
                {
                    throw new Exception("The directory of path doesn't exists.");
                }
                if (string.IsNullOrEmpty(Path.GetFileName(strFilePath)))
                {
                    throw new Exception("The file name of path is invalid.");
                }
                CommonDialogClass dlg = new CommonDialogClass();
                var item = Scanner.Items[1];
                try
                {
                    if (scConfig != null)
                    {
                        if (scConfig.IntColor_Mode.HasValue)
                        {
                            SetItem(item.Properties, "6146", scConfig.IntColor_Mode.Value);
                        }
                        if (scConfig.IntResolution_DPI_Horizontal.HasValue)
                        {
                            SetItem(item.Properties, "6147", scConfig.IntResolution_DPI_Horizontal);
                        }
                        if (scConfig.IntResolution_DPI_Vertical.HasValue)
                        {
                            SetItem(item.Properties, "6148", scConfig.IntResolution_DPI_Vertical);
                        }
                        if (scConfig.IntStart_Pixel_Horizontal.HasValue)
                        {
                            SetItem(item.Properties, "6149", scConfig.IntStart_Pixel_Horizontal.HasValue);
                        }
                        if (scConfig.IntStart_Pixel_Vertical.HasValue)
                        {
                            SetItem(item.Properties, "6150", scConfig.IntStart_Pixel_Vertical.HasValue);
                        }
                        if (scConfig.IntScan_Size_Pixels_Horizontal.HasValue)
                        {
                            SetItem(item.Properties, "6151", scConfig.IntScan_Size_Pixels_Horizontal.HasValue);
                        }
                        if (scConfig.IntScan_Size_Pixels_Vertical.HasValue)
                        {
                            SetItem(item.Properties, "6152", scConfig.IntScan_Size_Pixels_Vertical.HasValue);
                        }
                        if (scConfig.IntScan_Brightness_Percents.HasValue)
                        {
                            SetItem(item.Properties, "6154", scConfig.IntScan_Brightness_Percents.Value);
                        }
                        if (scConfig.IntScan_Contrast_Percents.HasValue)
                        {
                            SetItem(item.Properties, "6155", scConfig.IntScan_Contrast_Percents.Value);
                        }
                    }
                    if (scConfig == null || (scConfig != null && !scConfig.IntColor_Mode.HasValue))
                    {
                        SetItem(item.Properties, "6146", 1);
                    }
                    if (scConfig == null || (scConfig != null && !scConfig.IntResolution_DPI_Horizontal.HasValue))
                    {
                        SetItem(item.Properties, "6147", 150);
                    }
                    if (scConfig == null || (scConfig != null && !scConfig.IntResolution_DPI_Vertical.HasValue))
                    {
                        SetItem(item.Properties, "6148", 150);
                    }
                    string strFormat = string.Empty;
                    if (string.IsNullOrEmpty(strImageFormat))
                    {
                        throw new Exception("Not valid image format specify bmp, gif, jpg, png, tiff");
                    }
                    switch (strImageFormat.ToLower())
                    {
                        case "bmp":
                            strFormat = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}";
                            break;
                        case "gif":
                            strFormat = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}";
                            break;
                        case "jpg":
                            strFormat = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}";
                            break;
                        case "png":
                            strFormat = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}";
                            break;
                        case "tiff":
                            strFormat = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}";
                            break;
                        default:
                            throw new Exception("Not valid image format specify bmp, gif, jpg, png, tiff");
                    }
                    string strExtension = Path.GetExtension(strFilePath).ToLower();
                    if ((bolPdf && ".pdf" != strExtension) || (!bolPdf && strExtension != "." + strImageFormat))
                    {
                        string strValid = bolPdf ? "pdf" : strExtension;
                        throw new Exception("The file extension dosen't match with format please specify. " + strValid);
                    }
                    ImageFile img = (ImageFile)dlg.ShowTransfer(item, strFormat, true);
                    byte[] imgbte = (byte[])img.FileData.get_BinaryData();
                    MemoryStream ms = new MemoryStream(imgbte);
                    Image image = Image.FromStream(ms);
                    string strImageFile = !bolPdf ? strFilePath : Path.GetTempPath() + DateTime.Now.ToString("yyyyddM_HHmmss") + "_" + Path.GetFileName(strFilePath) + "." + strImageFormat.ToLower();
                    image.Save(strImageFile, ImageFormat.Jpeg);
                    if (bolPdf)
                    {
                        PdfDocument Pdf = new PdfDocument();
                        PdfPage Page = Pdf.AddPage();
                        XGraphics Gfx = XGraphics.FromPdfPage(Page);
                        XImage Img = XImage.FromFile(strImageFile);
                        Gfx.DrawImage(Img, 0, 0, (int)Page.Width, (int)Page.Height);
                        Pdf.Save(strFilePath);
                        //File.Delete(strImageFile);
                    }
                }
                catch (COMException e)
                {
                    uint errorCode = (uint)e.ErrorCode;
                    if (errorCode == 0x80210006)
                    {
                        throw new Exception("The scanner is busy or isn't ready");
                    }
                    else if (errorCode == 0x80210064)
                    {
                        throw new Exception("The scanning process has been cancelled.");
                    }
                    else
                    {
                        throw new Exception("A non catched error occurred, check the console");
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        #endregion

        //SetItem
        #region
        private static void SetItem(IProperties properties, object propName, object propValue)
        {
            Property prop = properties.get_Item(ref propName);
            prop.set_Value(ref propValue);
        }
        #endregion

        //Scanners List
        #region
        public static List<string> ScannersList()
        {
            List<string> lisScanners = new List<string>();
            try
            {
                DeviceManager dmScanners = new DeviceManager();
                for (int i = 1; i <= dmScanners.DeviceInfos.Count; i++)
                {
                    if (dmScanners.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType)
                    {
                        continue;
                    }
                    lisScanners.Add(dmScanners.DeviceInfos[i].Properties["Name"].get_Value().ToString());
                }
            }
            catch (Exception)
            {
                throw;
            }
            return lisScanners;
        }
        #endregion
    }
}
