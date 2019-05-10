using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Globalization;
using ImageResizer;
using ImageResizer.ExtensionMethods;

namespace Hyperfish.ImportExport.O365
{
    class ImageHelper
    {
        public static Stream ResizeImage(Stream image, int maxWidth)
        {
            image.Seek(0, SeekOrigin.Begin);

            var originalImage = Image.FromStream(image, true, true);

            int newHeight = (maxWidth * originalImage.Height) / originalImage.Width;

            Bitmap newImage = new Bitmap(maxWidth, newHeight);

            using (Graphics gr = Graphics.FromImage(newImage))
            {
                gr.SmoothingMode = SmoothingMode.HighQuality;
                gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                gr.PixelOffsetMode = PixelOffsetMode.HighQuality;
                gr.DrawImage(originalImage, new Rectangle(0, 0, maxWidth, newHeight)); //copy to new bitmap
            }
            
            MemoryStream memStream = new MemoryStream();
            newImage.Save(memStream, ImageFormat.Jpeg);
            originalImage.Dispose();
            memStream.Seek(0, SeekOrigin.Begin);

            return memStream;
        }

        public static MemoryStream ResizeImage(Stream originalImage, long maxSizeInBytes, int quality)
        {
            long currentSize = 0;

            MemoryStream imageStream = new MemoryStream();

            // pre-process to make sure its a JPG. need to do this first or the scale factor will not be right
            ResizeAndFormatImage(originalImage, imageStream, 1, quality);

            // get the current size    
            currentSize = imageStream.Length;

            // while it is too big
            while (currentSize > maxSizeInBytes)
            {
                // calc the scale factor
                double scale = Math.Sqrt((double)maxSizeInBytes / (double)imageStream.Length);

                // min 0.9 scale factor to stop excessive scaling iterations
                if (scale > 0.9) scale = 0.9;

                using (MemoryStream scaledImage = new MemoryStream())
                {
                    imageStream.Seek(0, SeekOrigin.Begin);

                    // resize the image
                    ResizeAndFormatImage(imageStream, scaledImage, scale, quality);

                    // get the new size
                    currentSize = scaledImage.Length;

                    // resize closes stream.  need to new it up again.
                    imageStream = scaledImage.CopyToMemoryStream(true);

                    if (currentSize < maxSizeInBytes)
                    {
                        return imageStream;
                    }
                }
            }

            return imageStream;
        }

        private static void ResizeAndFormatImage(Stream imageIn, Stream imageOut, double scaleFactor, int quality)
        {
            var settings = new ResizeSettings("zoom=" + scaleFactor.ToString(CultureInfo.CreateSpecificCulture("en-US")));
            settings.Scale = ScaleMode.DownscaleOnly;
            settings.Quality = quality;
            settings.Format = "jpg";

            imageIn.Seek(0, SeekOrigin.Begin);
            imageOut.Seek(0, SeekOrigin.Begin);

            // resize the image
            ImageResizer.ImageBuilder.Current.Build(imageIn, imageOut, settings);
        }
        
    }
}
