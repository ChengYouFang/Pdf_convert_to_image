using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace PDFConvertImage
{
    public partial class PDFtoImage : Form
    {
        public PDFtoImage()
        {
            InitializeComponent();
        }

        private void button_LoadFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = getDeskPath();
            dialog.Filter = "PDF|*.pdf";
            dialog.RestoreDirectory = true;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //倍數
                short pixelZoom = 10;

                //http://www.cnblogs.com/kongxianghai/archive/2012/04/09/2438321.html
                //頁
                Acrobat.CAcroPDPage pdfPage = null;

                //頁面大小
                Acrobat.CAcroPoint pdfPoint = null;

                //矩形
                Acrobat.CAcroRect pdfRect = null;

                //http://forums.codeguru.com/showthread.php?31751-What-do-I-need-for-CreateObject(-AcroExch.PDDoc-)
                Acrobat.CAcroPDDoc pdfDoc = (Acrobat.CAcroPDDoc)Microsoft.VisualBasic.Interaction.CreateObject("AcroExch.PDDoc", "");
                if (pdfDoc.Open(dialog.FileName))
                {
                    //PDF頁數
                    int pageCount = pdfDoc.GetNumPages();
                    for (int i = 0; i < pageCount; i++)
                    {
                        //取得目前頁數
                        pdfPage = (Acrobat.CAcroPDPage)pdfDoc.AcquirePage(i);

                        //取得頁面大小
                        pdfPoint = (Acrobat.CAcroPoint)pdfPage.GetSize();

                        //http://forums.adobe.com/thread/305835
                        pdfRect = (Acrobat.CAcroRect)Microsoft.VisualBasic.Interaction.CreateObject("AcroExch.Rect", "");

                        //左邊
                        pdfRect.Left = 0;

                        //頂端
                        pdfRect.Top = 0;

                        //右邊
                        pdfRect.right = (short)(pdfPoint.x * pixelZoom);

                        //底部
                        pdfRect.bottom = (short)(pdfPoint.y * pixelZoom);

                        //取得整個頁面
                        pdfPage.CopyToClipboard(pdfRect, (short)(pdfRect.Left * pixelZoom), (short)(pdfRect.Top * pixelZoom), (short)(100 * pixelZoom));

                        ///<summary>http://msdn.microsoft.com/en-us/library/system.windows.forms.idataobject.aspx
                        ///You retrieve stored data from an IDataObject by calling the GetData method and specifying the data format in the format parameter. Set the  autoConvert parameter to false to retrieve only data that was stored in the specified format. To convert the stored data to the specified format, set autoConvert to true, or do not use autoConvert.
                        ///</summary>
                        IDataObject loClipboardData = Clipboard.GetDataObject();

                        //Get Image
                        Bitmap pdfBitmap = (Bitmap)loClipboardData.GetData(DataFormats.Bitmap);

                        //Save Image
                        pdfBitmap.Save(getDeskPath() + (i + 1) + ".jpeg", ImageFormat.Jpeg);
                    }
                    pdfDoc.Close();

                    ///<summary>http://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.releasecomobject.aspx
                    ///This method is used to explicitly control the lifetime of a COM object used from managed code. You should use this method to free the underlying COM object that holds references to resources in a timely manner or when objects must be freed in a specific order.
                    ///</summary>
                    Marshal.ReleaseComObject(pdfPage);
                    Marshal.ReleaseComObject(pdfRect);
                    Marshal.ReleaseComObject(pdfDoc);
                    MessageBox.Show("轉換成功");
                }
                else
                {
                    MessageBox.Show("檔案可能損毀了");
                }
            }
        }

        //取得桌面路徑
        private static string getDeskPath()
        {
            return System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\";
        }

        private void PDFtoImage_Load(object sender, EventArgs e)
        {
            this.MaximizeBox = false;
            this.MinimizeBox = false;
        }
    }
}