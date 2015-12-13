using System;
using System.Net;

namespace JsonCrawlParser
{
    /// <summary>
    /// Utility to download a file from the web.
    /// </summary>
    public class FileDownload
    {
        /// <summary>
        /// Download a file stored at the given url.
        /// </summary>
        /// <param name="url">URL the file is stored at.</param>
        /// <param name="savePath">The path to save the file to.</param>
        public void Download(string url, string savePath) {
            
            try
            {
                using (WebClient client = new WebClient())
                {
                    client.DownloadFile(url, savePath);
                }
            }
            catch (Exception ex) {
                // Re-throw the exception.                
                throw ex;
            }
            
        }
    }
}
