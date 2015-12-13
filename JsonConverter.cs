using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace JsonCrawlParser
{
    /// <summary>
    /// Utility for converting json strings to product data.
    /// </summary>
    public class JsonConverter
    {
        /// <summary>
        /// Convert the given json string to a list of strongly typed objects.
        /// </summary>
        /// <param name="json">JSON string.</param>
        /// <param name="max">Optional max number of results to iterate over.</param>
        /// <returns>All products found within the given json.</returns>
        public List<ProductModel> Convert(string json, int max = -1)
        {
            /*******************************************************
            * CONSTANTS
            * ------------------------------------------------------
            * (none)
            ********************************************************/
            JavaScriptSerializer serializer; // Used to de-serialize the json objects.
            int productNodeCount;            // Total number of 
            List<ProductModel> products;     // Products to be returned.

            // Instantiate the return collection.
            products = new List<ProductModel>();

            // Instantiate the serializer and configure for use.
            serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = Int32.MaxValue;

            // De-serialize the json string into dynamic objects which can be iterated.
            dynamic rootNode = serializer.Deserialize<dynamic>(json);

            // Get the products node from the root.
            //    Result is a Dictionary<string,object>.
            dynamic productsNode = rootNode["products"]; 

            // Get the total number of product notes.
            productNodeCount = productsNode.Length;

            // If the product max isn't the default then use the given max as the product count
            //   so that the iteration stops at the max.
            if (max != -1) productNodeCount = max;

            // Iterate over all the product nodes in the data source.
            for (int i = 0; i < productNodeCount; i++)
            {
                /*******************************************************
                * CONSTANTS
                * ------------------------------------------------------
                * (none)
                ********************************************************/
                ProductModel product; // The product currently looped to.

                // Instantiate a new instance of the product.
                product = new ProductModel();

                // Populate the product.
                product.Manufacturer = GetNodeProperty(productsNode[i], "manu");
                product.ImageUrl = GetNodeProperty(productsNode[i], "image_src");
                product.Count = GetNodeProperty(productsNode[i], "count", "0");
                product.UnitOfMeasure = GetNodeProperty(productsNode[i], "unit_measure");
                product.Model = GetNodeProperty(productsNode[i], "model");
                product.Short = GetNodeProperty(productsNode[i], "short");
                product.NetPrice = GetNodeProperty(productsNode[i], "net_price");
                product.ImageBase64 = GetNodeProperty(productsNode[i], "image");
                product.Warehouse = GetNodeProperty(productsNode[i], "warehouse");
                product.Doc1Name = GetNodeProperty(productsNode[i], "doc_1_name");
                product.Doc1Href = GetNodeProperty(productsNode[i], "doc_1_href");
                product.Doc2Name = GetNodeProperty(productsNode[i], "doc_2_name");
                product.Doc2Href = GetNodeProperty(productsNode[i], "doc_2_href");
                product.Vendor = GetNodeProperty(productsNode[i], "vendor");
                product.Code = GetNodeProperty(productsNode[i], "code");
                

                try
                {
                    product.Alt_FEI = productsNode[i]["add_prod_codes"][0]["FEI"];
                }
                catch (Exception) { product.Alt_FEI = ""; }

                try
                {
                    product.Alt_Manu_Code = productsNode[i]["add_prod_codes"][0]["manu_codes"];
                }
                catch (Exception) { product.Alt_Manu_Code = ""; }


                try
                {
                    product.Alt_UPC_Code = productsNode[i]["add_prod_codes"][0]["upc_codes"];
                }
                catch (Exception) { product.Alt_UPC_Code = ""; }

                // ********************************
                // Product Features.
                // ********************************
                try
                {                    
                    var featCount = productsNode[i]["feature_labels"].Length;
                    dynamic featureLabels = productsNode[i]["feature_labels"];
                    dynamic featureValues = productsNode[i]["feature_values"];

                    // Iterate over Features
                    for (var f = 0; f < featCount; f++)
                        product.Features.Add(new KeyValuePair<string, string>(featureLabels[f]["label"], featureValues[f]["value"]));
                }
                catch (Exception) { /* Ignore Errors */}

                // ********************************
                // Bullets
                // ********************************
                try
                {
                    var bulletCount = productsNode[i]["bullets"].Length;
                    dynamic bullets = productsNode[i]["bullets"];

                    for (var b = 0; b < bulletCount; b++)
                        product.Bullets.Add(bullets[b]["bullet"]);
                }
                catch (Exception) { /* Ignore Errors */}

                // Add the populated product to the return collection.
                products.Add(product);
            }

            // Return the products.
            return products; 
        }

        /// <summary>
        /// Attempt to retrieve a property's value from the given node.
        /// </summary>
        /// <param name="node">The node to get a value for.</param>
        /// <param name="key">The property's name.</param>
        /// <param name="defaultValue">The default value to be returned when a property is not found.</param>
        /// <returns>The property value or default value.</returns>
        private string GetNodeProperty(dynamic node, string key, string defaultValue = "")
        {
            try
            {
                // Attempt to get the node's property value.
                return node[key];
            }
            catch (Exception)
            {
                // Return the default value.
                return defaultValue;
            }
        }
        
        /// <summary>
        /// Object that represents a single product.
        /// </summary>
        public class ProductModel
        {
            public string Manufacturer { get; set; }
            public string ImageUrl { get; set; }
            public string Count { get; set; }
            public string UnitOfMeasure { get; set; }
            public string Model { get; set; }
            public string Short { get; set; }
            public string NetPrice { get; set; }
            public string ImageBase64 { get; set; }
            public string Warehouse { get; set; }
            public string Doc1Name { get; set; }
            public string Doc1Href { get; set; }
            public string Doc2Name { get; set; }
            public string Doc2Href { get; set; }
            public string Vendor { get; set; }
            public string Code { get; set; }
            public string Alt_FEI { get; set; }
            public string Alt_Manu_Code { get; set; }
            public string Alt_UPC_Code { get; set; }
            public List<string> Bullets { get; set; }
            public List<KeyValuePair<string, string>> Features { get; set; }

            public ProductModel()
            {
                this.Features = new List<KeyValuePair<string, string>>();
                this.Bullets = new List<string>();
            }

        }        
    }

    
    
}
