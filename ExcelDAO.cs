using System;
using System.Collections.Generic;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace JsonConverter
{
    /// <summary>
    /// Utility for saving products to Excel.
    /// </summary>
    public class ExcelDAO
    {
        /// <summary>
        /// The full path to the Excel file to be created and written to.
        /// </summary>
        string _fileName;

        /// <summary>
        /// The Excel worksheet columns.
        /// </summary>
        enum Columns
        {
            [Description("MANU")] // Descriptions are used as worksheet column headers.
            MANU = 1,
            [Description("MODEL")]
            MODEL,
            [Description("SHORT")]
            SHORT,
            [Description("NET_PRICE")]
            NET_PRICE,
            [Description("IMAGE")]
            IMAGE,
            [Description("IMAGE_URL")]
            IMAGE_URL,
            [Description("WAREHOUSE")]
            WAREHOUSE,
            [Description("DOC1NAME")]
            DOC1NAME,
            [Description("DOC1HREF")]
            DOC1HREF,
            [Description("DOC2NAME")]
            DOC2NAME,
            [Description("DOC2HREF")]
            DOC2HREF,
            [Description("VENDOR")]
            VENDOR,
            [Description("CODE")]
            CODE,
            [Description("ALT_FEI")]
            ALT_FEI ,
            [Description("ALT_MANU")]
            ALT_MANU,
            [Description("ALT_UPS")]
            ALT_UPS,
            [Description("COUNT")]
            COUNT, 
            [Description("UNIT_OF_MEASURE")]
            UNIT_OF_MEASURE,
            [Description("BULLETS")]
            BULLETS,
            [Description("FEATURES...")]
            FEATURES
        }

        /// <summary>
        /// Utility's constructor.
        /// </summary>
        /// <param name="fileName"></param>
        public ExcelDAO(string fileName)
        {
            _fileName = fileName;
        }

        /// <summary>
        /// Write the 
        /// </summary>
        /// <param name="products">The products to be written.</param>
        public void Write(ref List<JsonCrawlParser.JsonConverter.ProductModel> products)
        {
            /*******************************************************
            * CONSTANTS
            * ------------------------------------------------------
            * SHEET_NAME : The worksheet name to write data to.
            * DEF_ROW    : The default row to start with in the worksheet.
            ********************************************************/
            const string SHEET_NAME = "Data";
            const int DEF_ROW = 1;
            Excel.Application excelApp; // Excel application reference.
            Excel.Workbook wb;          // Excel workbook reference.
            Excel.Worksheet sh;         // Excel worksheet reference.
            int worksheetRow;           // The current product row.
            int columnCount;            // Known Excel column count.

            try
            {
                // Instantiate a new instance of Excel.
                excelApp = new Excel.Application();
            }
            catch (Exception ex)
            {
                throw new Exception("Error creating excel: " + ex.Message);
            }


            // Initialize.
            wb = excelApp.Workbooks.Add();
            sh = wb.Sheets.Add();
            worksheetRow = DEF_ROW;
            sh.Name = SHEET_NAME;
            
            // Get the total column count (i.e. known column count).
            columnCount = Enum.GetNames(typeof(Columns)).Length;

            
            try
            {


                // ********************************
                // Headers
                // ********************************

                // Write all header to the worksheet.
                sh.Cells[worksheetRow, (int)Columns.MANU].Value2 = EnumUtils<Columns>.GetDescription(Columns.MANU);
                sh.Cells[worksheetRow, (int)Columns.MODEL].Value2 = EnumUtils<Columns>.GetDescription(Columns.MODEL);
                sh.Cells[worksheetRow, (int)Columns.SHORT].Value2 = EnumUtils<Columns>.GetDescription(Columns.SHORT);
                sh.Cells[worksheetRow, (int)Columns.NET_PRICE].Value2 = EnumUtils<Columns>.GetDescription(Columns.NET_PRICE);
                sh.Cells[worksheetRow, (int)Columns.IMAGE].Value2 = EnumUtils<Columns>.GetDescription(Columns.IMAGE);
                sh.Cells[worksheetRow, (int)Columns.IMAGE_URL].Value2 = EnumUtils<Columns>.GetDescription(Columns.IMAGE_URL);
                sh.Cells[worksheetRow, (int)Columns.WAREHOUSE].Value2 = EnumUtils<Columns>.GetDescription(Columns.WAREHOUSE);
                sh.Cells[worksheetRow, (int)Columns.DOC1NAME].Value2 = EnumUtils<Columns>.GetDescription(Columns.DOC1NAME);
                sh.Cells[worksheetRow, (int)Columns.DOC1HREF].Value2 = EnumUtils<Columns>.GetDescription(Columns.DOC1HREF);
                sh.Cells[worksheetRow, (int)Columns.DOC2NAME].Value2 = EnumUtils<Columns>.GetDescription(Columns.DOC2NAME);
                sh.Cells[worksheetRow, (int)Columns.DOC2HREF].Value2 = EnumUtils<Columns>.GetDescription(Columns.DOC2HREF);
                sh.Cells[worksheetRow, (int)Columns.VENDOR].Value2 = EnumUtils<Columns>.GetDescription(Columns.VENDOR);
                sh.Cells[worksheetRow, (int)Columns.CODE].Value2 = EnumUtils<Columns>.GetDescription(Columns.CODE);
                sh.Cells[worksheetRow, (int)Columns.ALT_FEI].Value2 = EnumUtils<Columns>.GetDescription(Columns.ALT_FEI);
                sh.Cells[worksheetRow, (int)Columns.ALT_MANU].Value2 = EnumUtils<Columns>.GetDescription(Columns.ALT_MANU);
                sh.Cells[worksheetRow, (int)Columns.ALT_UPS].Value2 = EnumUtils<Columns>.GetDescription(Columns.ALT_UPS);
                sh.Cells[worksheetRow, (int)Columns.COUNT].Value2 = EnumUtils<Columns>.GetDescription(Columns.COUNT);
                sh.Cells[worksheetRow, (int)Columns.UNIT_OF_MEASURE].Value2 = EnumUtils<Columns>.GetDescription(Columns.UNIT_OF_MEASURE);
                sh.Cells[worksheetRow, (int)Columns.BULLETS].Value2 = EnumUtils<Columns>.GetDescription(Columns.BULLETS);
                sh.Cells[worksheetRow, (int)Columns.FEATURES].Value2 = EnumUtils<Columns>.GetDescription(Columns.FEATURES);

                // Advance to the first data row.
                worksheetRow++;

                // ********************************
                // Product Rows
                // ********************************

                // Write all product rows.
                foreach (var item in products)
                {
                    int columnIndex = columnCount;

                    sh.Cells[worksheetRow, (int)Columns.MANU].Value2 = item.Manufacturer;
                    sh.Cells[worksheetRow, (int)Columns.MODEL].Value2 = item.Model;
                    sh.Cells[worksheetRow, (int)Columns.SHORT].Value2 = item.Short;
                    sh.Cells[worksheetRow, (int)Columns.NET_PRICE].Value2 = item.NetPrice;
                    sh.Cells[worksheetRow, (int)Columns.IMAGE].Value2 = item.ImageBase64;
                    sh.Cells[worksheetRow, (int)Columns.IMAGE_URL].Value2 = item.ImageUrl;
                    sh.Cells[worksheetRow, (int)Columns.WAREHOUSE].Value2 = item.Warehouse;
                    sh.Cells[worksheetRow, (int)Columns.DOC1NAME].Value2 = item.Doc1Name;
                    sh.Cells[worksheetRow, (int)Columns.DOC1HREF].Value2 = item.Doc1Href;
                    sh.Cells[worksheetRow, (int)Columns.DOC2NAME].Value2 = item.Doc2Name;
                    sh.Cells[worksheetRow, (int)Columns.DOC2HREF].Value2 = item.Doc2Href;
                    sh.Cells[worksheetRow, (int)Columns.VENDOR].Value2 = item.Vendor;
                    sh.Cells[worksheetRow, (int)Columns.CODE].Value2 = item.Code;
                    sh.Cells[worksheetRow, (int)Columns.ALT_FEI].Value2 = item.Alt_FEI;
                    sh.Cells[worksheetRow, (int)Columns.ALT_MANU].Value2 = item.Alt_Manu_Code;
                    sh.Cells[worksheetRow, (int)Columns.ALT_UPS].Value2 = item.Alt_UPC_Code;
                    sh.Cells[worksheetRow, (int)Columns.COUNT].Value2 = item.Count;
                    sh.Cells[worksheetRow, (int)Columns.UNIT_OF_MEASURE].Value2 = item.UnitOfMeasure;

                    // Write the bullets as a comma delimited string.
                    sh.Cells[worksheetRow, (int)Columns.BULLETS].Value2 = String.Join(", ", item.Bullets);

                    // Store each feature in its own column.
                    foreach (var feature in item.Features)
                        sh.Cells[worksheetRow, columnIndex++].Value2 = String.Format("{0} : {1}", feature.Key, feature.Value);

                    // Advance to the next worksheet row.
                    worksheetRow++;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Excel write error: " + ex.Message);
            }
            finally
            {
                // Save the file and clean up.
                wb.SaveAs(_fileName);
                wb.Close();
                excelApp.Quit();
            }
        }

    }
}
