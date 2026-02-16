using System;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace WpfApp1.Shared.Helpers
{
    public static class ImageHelper
    {
        public static byte[] CompressImage(byte[] sourceBytes, double targetWidthCm, double targetHeightCm)
        {
            if (sourceBytes == null || sourceBytes.Length == 0) return Array.Empty<byte>();
            if (targetWidthCm <= 0 || targetHeightCm <= 0) return sourceBytes;

            try
            {
                using (var msInput = new MemoryStream(sourceBytes))
                {
                    var decoder = BitmapDecoder.Create(msInput, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                    var originalFrame = decoder.Frames[0];


                    double targetPixelWidth = (targetWidthCm / 2.54) * 96.0;
                    double targetPixelHeight = (targetHeightCm / 2.54) * 96.0;

                    double scaleX = targetPixelWidth / originalFrame.PixelWidth;
                    double scaleY = targetPixelHeight / originalFrame.PixelHeight;


                    var transformedBitmap = new TransformedBitmap(
                        originalFrame,
                        new ScaleTransform(scaleX, scaleY));

                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(transformedBitmap));

                    using (var msOutput = new MemoryStream())
                    {
                        encoder.Save(msOutput);
                        return msOutput.ToArray();
                    }
                }
            }
            catch (Exception)
            {
                return sourceBytes;
            }
        }
    }
}