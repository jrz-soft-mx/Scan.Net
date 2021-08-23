# ScanNet
Tool for scan from your APP save to BMP, PNG, GIF, JPEG, TIFF, and add the image into PDF
# ![MD5-Updater](Tools/logo.png)

MD5 Update is a library written in C# for easily allow easy add update functionality to desktops applications only need webserver with PHP for publish the update directory. 

## The NuGet Package

````powershell
PM> Install-Package MD5.Update
````

## How it works

Scan with Windows Image Adquisicion,

### Save Image JPG into PDF
Example for save image to JPG into PDF

````csharp
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace HelloWorld
{
    public partial class HelloWorld : Form
    {
        public HelloWorld()
        {		
            InitializeComponent();
            //Configuration by defaul
            ScanNet.ScanNet.ScanCfg cfg = new ScanNet.ScanNet.ScanCfg();
            //4 black and white 2 grayscale 1 color
            cfg.IntColor_Mode = 1;
            cfg.IntResolution_DPI_Horizontal = 150;
            cfg.IntResolution_DPI_Vertical = 150;
            //scan name, filepath, format, pdf, ScanCfg
            ScanNet.ScanNet.Scan("SP 200 Series", @"C:\Users\Admin\Desktop\scan.pdf", "jpg", false, cfg);
	      }
    }
}
````

### Save Image to JPG
Example for save image to JPG

````csharp
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace HelloWorld
{
    public partial class HelloWorld : Form
    {
        public HelloWorld()
        {		
            InitializeComponent();
            //scan name, filepath, format, pdf, ScanCfg
            ScanNet.ScanNet.Scan("SP 200 Series", @"C:\Users\Admin\Desktop\scan.jpg", "jpg");
	      }
    }
}
````

### Updt.exe (Tool for replace yourapp).
Get list string of aviable scanners
````csharp
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace HelloWorld
{
    public partial class HelloWorld : Form
    {
        public HelloWorld()
        {		
            InitializeComponent();
            List<string> lisScanners = ScanNet.ScanNet.ScannersList();
            foreach(string strScanner in lisScanners)
            {
                comboBox1.Items.Add(strScanner);
            }
	      }
    }
}
