# ScanNet
Tool for scan from your APP save to BMP, PNG, GIF, JPEG, TIFF, and add the image into PDF

## How it works

Scan with Windows Image Acquisition (WIA), only need call Scan with scanner name, filepath and image format, addicionaly your set true for make pdf, create new instance of configuration class for set color mode, DPI, Start Pixels, Size of Scan, Brigthness and contrast.

### Save Image JPG into PDF
Example for save image to JPG into PDF

````csharp
//Configuration by defaul and example of new instance for configuration class
ScanNet.ScanNet.ScanCfg cfg = new ScanNet.ScanNet.ScanCfg();
//4 black and white 2 grayscale 1 color
cfg.IntColor_Mode = 1;
cfg.IntResolution_DPI_Horizontal = 150;
cfg.IntResolution_DPI_Vertical = 150;
//scan name, filepath, format, pdf, ScanCfg
ScanNet.ScanNet.Scan("SP 200 Series", @"C:\Users\Admin\Desktop\scan.pdf", "jpg", false, cfg);
````

### Save Image to JPG
Example for save image to JPG

````csharp
//scan name, filepath, format, pdf, ScanCfg
ScanNet.ScanNet.Scan("SP 200 Series", @"C:\Users\Admin\Desktop\scan.jpg", "jpg");
````

### Get list of available 
Get list string of available scanners

````csharp
List<string> lisScanners = ScanNet.ScanNet.ScannersList();
//fill combo with scanners for user select.
foreach(string strScanner in lisScanners)
{
    comboBox1.Items.Add(strScanner);
}
