using System;
using System.Drawing;
using System.IO;


namespace JsonCrawlParser
{
    /// <summary>
    /// Utility class for converting Base64 image strings to images.
    /// </summary>
    public class ImageConverter
    {
        /// <summary>
        /// Save the base64 string to the given image location.
        /// </summary>
        /// <param name="base64">Base64 string.</param>
        /// <param name="path">Save path.</param>
        public void SaveImage(string base64, string path)
        {
            // Load the image into memory and save to the given location.
            using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(base64)))
            using (Bitmap bm2 = new Bitmap(ms))
            {
                bm2.Save(path);
            }
            
        }
    }
}
