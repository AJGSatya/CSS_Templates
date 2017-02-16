using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace Aga.Controls
{
    public static class BitmapHelper
    {
        public static void SetAlphaChanelValue(Bitmap image, byte value)
        {
            if (image == null)
                throw new ArgumentNullException("image");
            if (image.PixelFormat != PixelFormat.Format32bppArgb)
                throw new ArgumentException("Wrong PixelFormat");

            BitmapData bitmapData = image.LockBits(new Rectangle(0, 0, image.Width, image.Height),
                                                   ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
            unsafe
            {
                var pPixel = (PixelData*) bitmapData.Scan0;
                for (int i = 0; i < bitmapData.Height; i++)
                {
                    for (int j = 0; j < bitmapData.Width; j++)
                    {
                        pPixel->A = value;
                        pPixel++;
                    }
                    pPixel += bitmapData.Stride - (bitmapData.Width*4);
                }
            }
            image.UnlockBits(bitmapData);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct PixelData
        {
            public readonly byte B;
            public readonly byte G;
            public readonly byte R;
            public byte A;
        }
    }
}