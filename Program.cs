using JsonConverter;
using System;
using System.Collections.Generic;

namespace JsonCrawlParser
{
    /// <summary>
    /// An application used to parse the results of a Ferguson client site crawl performed by
    /// ParseHub.
    /// </summary>
    class Program
    {
        /// <summary>
        /// Main entry point.
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            /*******************************************************
            * CONSTANTS
            * ------------------------------------------------------
            * (none)
            ********************************************************/
            string resourcePath;                        // The root path to save documents/images.
            JsonConverter converter;                    // Converts json to product data.
            int imageCount;                             // Total number of images saved locally.
            string json;                                // Json string contained within the source file (raw products).
            List<JsonConverter.ProductModel> products;  // All products found in json.
            string choice;                              // Response to prompt question.
            string jsonSource;                          // Full path to json file containing crawl results.

            // Instantiate / initialize.
            converter = new JsonConverter();
            imageCount = 0;

            // Get the full path to the json file containing product crawl results.
            Console.WriteLine("Read path to .json crawl file:");
            jsonSource = Console.ReadLine();

            // Read all text from the json file and convert to products.
            Console.WriteLine("Parsing Data...");
            json = System.IO.File.ReadAllText(jsonSource);
            products = converter.Convert(json);

            // Ask the user if they would like to download the product's associated documents?
            Console.WriteLine("Download documents: y/n");
            choice = Console.ReadLine();

            // Should the product's documents be downloaded?
            if (choice == "y")
            {
                /*******************************************************
                * CONSTANTS
                * ------------------------------------------------------
                * (none)
                ********************************************************/
                FileDownload fileLoader; // Used to download the document file from the web.

                // Get the path to the folder the resources should be saved to.
                Console.WriteLine("Write path for documents: ");
                resourcePath = Console.ReadLine(); // Ex. @"C:\Users\{User Name}\Desktop\product_images\";
                
                // Instantiate.
                fileLoader = new FileDownload();

                // Iterate over all of the products.
                foreach (var item in products)
                {
                    /*******************************************************
                    * CONSTANTS
                    * ------------------------------------------------------
                    * (none)
                    ********************************************************/
                    string savePath; // Save path for the product's document.
                    
                    // Is there a document?
                    if (item.Doc1Href.Length != 0)
                    {
                        // Set the save path to the resource root. Use the product's model number
                        //   and the document's name as the file name.
                        savePath = String.Format("{0}{1}_{2}.pdf", resourcePath, item.Model, item.Doc1Name);

                        // Download the file.
                        fileLoader.Download(item.Doc1Href, savePath);
                    }

                    // Is there a second document?
                    if (item.Doc2Href.Length != 0)
                    {
                        // Set the save path to the resource root. Use the product's model number
                        //   and the document's name as the file name.
                        savePath = String.Format("{0}{1}_{2}.pdf", resourcePath, item.Model, item.Doc2Name);

                        // Download the file.
                        fileLoader.Download(item.Doc2Href, savePath);
                    }
                }
            }

            // Ask the user if they would like to download the product's images?
            Console.WriteLine("Download images: y/n");
            choice = Console.ReadLine();
            
            // Download the images?
            if (choice == "y")
            {
                /*******************************************************
                * CONSTANTS
                * ------------------------------------------------------
                * (none)
                ********************************************************/
                ImageConverter imageConv;

                // Instantiate.
                imageConv = new ImageConverter();

                // Get the path to the folder the resources should be saved to.
                Console.WriteLine("Write path for images: ");
                resourcePath = Console.ReadLine(); // Ex. @"C:\Users\{User Name}\Desktop\product_images\";

                // Update user that the save it in process.
                Console.WriteLine("Save Images...");

                // Iterate over all products.
                foreach (var item in products)
                {
                    // Is there an image to convert?
                    if (item.ImageBase64.Length != 0 && item.Model.Length != 0)
                    {
                        /*******************************************************
                        * CONSTANTS
                        * ------------------------------------------------------
                        * (none)
                        ********************************************************/
                        //"data:image/png;base64,iVBORw0KGgoA...
                        string[] imageData = item.ImageBase64.Split(',');
                        string[] typeData = imageData[0].Split(';');
                        string type = typeData[0].Replace("data:image/", "");

                        // Save the image to the resources folder.
                        imageConv.SaveImage(imageData[1], String.Format("{0}{0}.{1}", resourcePath, item.Model, type));

                        Console.Clear();
                        Console.WriteLine(String.Format("Image: {0}", ++imageCount));
                    }
                }
            }

            // Ask the user if they would like to write the products to Excel?
            Console.WriteLine("Write to Excel: y/n");
            choice = Console.ReadLine();

            // Write to Excel?
            if (choice == "y")
            {
                /*******************************************************
                * CONSTANTS
                * ------------------------------------------------------
                * (none)
                ********************************************************/
                string savePath; // The path to save the products to.
                ExcelDAO dao;    // Utility for saving the products to Excel.

                // Get save path.
                Console.WriteLine("Full path for xls file to be created: ");
                savePath = Console.ReadLine(); // Ex. @"C:\Users\{User Name}\Desktop\pipe_and_tubing_Parsed.xls";
                
                // Save.
                dao = new ExcelDAO(savePath);
                dao.Write(ref products);
            }

            // Done.
            Console.WriteLine("Done!");
            Console.Read();
        }
    }
}
