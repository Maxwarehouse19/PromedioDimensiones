using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Data.SqlClient;
using System.Data;

namespace PromedioDimensionesMaxwarehouse.Clases
{
    class ManipulaExcel
    {
        public String fromFile = ConfigurationManager.AppSettings["NombreArchivoBase"];
        public string toFile = ConfigurationManager.AppSettings["NombreCsv"];

        // Abre el Excel para manipular contenido
        // --------------------------------------
        private void AbreExcel(ref Microsoft.Office.Interop.Excel.Application app, ref Microsoft.Office.Interop.Excel.Workbook wb)
        {

            app = new Microsoft.Office.Interop.Excel.Application();
            wb = app.Workbooks.Open(fromFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }

        // realiza arreglos para convertir el archivo en csv
        // -------------------------------------------------
        private void RealizaArreglosExcel(ref Microsoft.Office.Interop.Excel.Application app, ref Microsoft.Office.Interop.Excel.Workbook wb)
        {
            app.DisplayAlerts = false;
            app.Cells.Replace(What: "\"", Replacement: " ", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, MatchByte: false, SearchFormat: false, ReplaceFormat: false);
            app.Cells.Replace(What: ",", Replacement: " ", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, MatchByte: false, SearchFormat: false, ReplaceFormat: false);
            wb.SaveAs(toFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, false, Type.Missing, Type.Missing, Type.Missing);
            wb.Close(false, Type.Missing, Type.Missing);
            app.Quit();
        }

        // crea archivo csv para cargarlo a SQL
        // ------------------------------------
        public void CreaArchivoCsv()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb = new Microsoft.Office.Interop.Excel.Workbook();
                // Abre el Excel para manipular contenido
                // --------------------------------------
                AbreExcel(ref app, ref wb);

                // realiza arreglos para convertir el archivo en csv
                // -------------------------------------------------
                RealizaArreglosExcel(ref app, ref wb);
            }
            catch (SystemException exp)
            {
                Console.Write("Error: " + exp.Message);
                throw new System.SystemException("Error en programa");
            }
        }

        // crea el Excel coincidencias entre carrier y DW
        // ----------------------------------------------
        private void CreaExcelFacturadas()
        {
            //Create a new ExcelPackage
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //Set some properties of the Excel document
                excelPackage.Workbook.Properties.Author = "EAVD";
                excelPackage.Workbook.Properties.Title = "Archivo de Coincidencias";
                excelPackage.Workbook.Properties.Subject = "Creacion de Excel";
                excelPackage.Workbook.Properties.Created = DateTime.Now;

                //Create the WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                //Crea encabezado de Reporte
                //--------------------------
                worksheet.Cells[1, 1].Value = "SalesOrderNumber";
                worksheet.Cells[1, 2].Value = "TotalSales";
                worksheet.Cells[1, 3].Value = "HoldCode";
                worksheet.Cells[1, 4].Value = "SalesSku";
                worksheet.Cells[1, 5].Value = "SalesCategoryAtTimeOfSale";
                worksheet.Cells[1, 6].Value = "UomCode";
                worksheet.Cells[1, 7].Value = "UomQuantity";
                worksheet.Cells[1, 8].Value = "SalesStatus";
                worksheet.Cells[1, 9].Value = "SalesOrderDate";
                worksheet.Cells[1, 10].Value = "SalesChannelName";
                worksheet.Cells[1, 11].Value = "FulfillmentSku";
                worksheet.Cells[1, 12].Value = "CustomerName";
                worksheet.Cells[1, 13].Value = "FulfillmentChannelName";
                worksheet.Cells[1, 14].Value = "LinkedFulfillmentChannelName";
                worksheet.Cells[1, 15].Value = "FulfillmentLocationName";
                worksheet.Cells[1, 16].Value = "FulfillmentChannelType";
                worksheet.Cells[1, 17].Value = "FulfillmentOrderNumber";
                worksheet.Cells[1, 18].Value = "Quantity";
                worksheet.Cells[1, 19].Value = "Sku";
                worksheet.Cells[1, 20].Value = "Title";
                worksheet.Cells[1, 21].Value = "TotalCost";
                worksheet.Cells[1, 22].Value = "Commission";
                worksheet.Cells[1, 23].Value = "InventoryCost";
                worksheet.Cells[1, 24].Value = "UnitCost";
                worksheet.Cells[1, 25].Value = "ServiceCost";
                worksheet.Cells[1, 26].Value = "EstimatedShippingCost";
                worksheet.Cells[1, 27].Value = "ShippingCost";
                worksheet.Cells[1, 28].Value = "ShippingPrice";
                worksheet.Cells[1, 29].Value = "OverheadCost";
                worksheet.Cells[1, 30].Value = "PackageCost";
                worksheet.Cells[1, 31].Value = "ProfitLoss";
                worksheet.Cells[1, 32].Value = "Carrier";
                worksheet.Cells[1, 33].Value = "ShippingServiceLevel";
                worksheet.Cells[1, 34].Value = "ShippedByUser";
                worksheet.Cells[1, 35].Value = "ShippingWeight";
                worksheet.Cells[1, 36].Value = "Weight";
                worksheet.Cells[1, 37].Value = "Width";
                worksheet.Cells[1, 38].Value = "Length";
                worksheet.Cells[1, 39].Value = "Height";
                worksheet.Cells[1, 40].Value = "StateRegion";
                worksheet.Cells[1, 41].Value = "TrackingNum";
                worksheet.Cells[1, 42].Value = "MfrName";
                worksheet.Cells[1, 43].Value = "PricingRule";
                worksheet.Cells[1, 44].Value = "Ground Tracking ID Prefix";
                worksheet.Cells[1, 45].Value = "Express or Ground Tracking ID";
                worksheet.Cells[1, 46].Value = "Net Charge Amount";
                worksheet.Cells[1, 47].Value = "Service Type";
                worksheet.Cells[1, 48].Value = "Ground Service";
                worksheet.Cells[1, 49].Value = "Shipment Date";
                worksheet.Cells[1, 50].Value = "POD Delivery Date";
                worksheet.Cells[1, 51].Value = "Actual Weight Amount";
                worksheet.Cells[1, 52].Value = "Rated Weight Amount";
                worksheet.Cells[1, 53].Value = "Dim Length";
                worksheet.Cells[1, 54].Value = "Dim Width";
                worksheet.Cells[1, 55].Value = "Dim Height";
                worksheet.Cells[1, 56].Value = "Dim Divisor";
                worksheet.Cells[1, 57].Value = "Shipper State";
                worksheet.Cells[1, 58].Value = "Zone Code";
                worksheet.Cells[1, 59].Value = "Tendered Date";
                worksheet.Cells[1, 60].Value = "Earned Discount";
                worksheet.Cells[1, 61].Value = "Fuel Surcharge";
                worksheet.Cells[1, 62].Value = "Performance Pricing";
                worksheet.Cells[1, 63].Value = "Delivery Area Surcharge Extended";
                worksheet.Cells[1, 64].Value = "Delivery Area Surcharge";
                worksheet.Cells[1, 65].Value = "USPS Non-Mach Surcharge";
                worksheet.Cells[1, 66].Value = "Residential";
                worksheet.Cells[1, 67].Value = "Grace Discount";
                worksheet.Cells[1, 68].Value = "Declared Value";
                worksheet.Cells[1, 69].Value = "DAS Extended Resi";
                worksheet.Cells[1, 70].Value = "Additional Handling";
                worksheet.Cells[1, 71].Value = "Parcel Re-Label Charge";
                worksheet.Cells[1, 72].Value = "Indirect Signature";
                worksheet.Cells[1, 73].Value = "DAS Resi";
                worksheet.Cells[1, 74].Value = "Address Correction";
                worksheet.Cells[1, 75].Value = "DAS Extended Comm";
                worksheet.Cells[1, 76].Value = "Oversize Charge";
                worksheet.Cells[1, 77].Value = "AHS - Dimensions";
                worksheet.Cells[1, 78].Value = "Mail Class ";
                worksheet.Cells[1, 79].Value = "Tracking Number ";
                worksheet.Cells[1, 80].Value = "Total Postage Amt ";
                worksheet.Cells[1, 81].Value = "Delivery Date ";
                worksheet.Cells[1, 82].Value = "Weight ";
                worksheet.Cells[1, 83].Value = "Zone ";
                worksheet.Cells[1, 84].Value = "Service Type ";
                worksheet.Cells[1, 85].Value = "Tracking Number ";
                worksheet.Cells[1, 86].Value = "Net Charge Amount ";

                // Realiza Lectura de la Base de Datos y obtiene las coincidencias para imprimir el reporte
                // ----------------------------------------------------------------------------------------
                SqlConnection Conn = null;
                string query = "";
                SqlCommand command = new SqlCommand(query, Conn);
                command.CommandTimeout = 600;
                Conn.Open();
                SqlDataReader reader = command.ExecuteReader();
                int Fila = 2;
                try
                {
                    while (reader.Read())
                    {
                        worksheet.Cells[Fila, 1].Value  = reader["SalesOrderNumber"].ToString();
                        worksheet.Cells[Fila, 2].Value  = reader["TotalSales"].ToString();
                        worksheet.Cells[Fila, 3].Value  = reader["HoldCode"].ToString();
                        worksheet.Cells[Fila, 4].Value  = reader["SalesSku"].ToString();
                        worksheet.Cells[Fila, 5].Value  = reader["SalesCategoryAtTimeOfSale"].ToString();
                        worksheet.Cells[Fila, 6].Value  = reader["UomCode"].ToString();
                        worksheet.Cells[Fila, 7].Value  = reader["UomQuantity"].ToString();
                        worksheet.Cells[Fila, 8].Value  = reader["SalesStatus"].ToString();
                        worksheet.Cells[Fila, 9].Value  = reader["SalesOrderDate"].ToString();
                        worksheet.Cells[Fila, 10].Value = reader["SalesChannelName"].ToString();
                        worksheet.Cells[Fila, 11].Value = reader["FulfillmentSku"].ToString();
                        worksheet.Cells[Fila, 12].Value = reader["CustomerName"].ToString();
                        worksheet.Cells[Fila, 13].Value = reader["FulfillmentChannelName"].ToString();
                        worksheet.Cells[Fila, 14].Value = reader["LinkedFulfillmentChannelName"].ToString();
                        worksheet.Cells[Fila, 15].Value = reader["FulfillmentLocationName"].ToString();
                        worksheet.Cells[Fila, 16].Value = reader["FulfillmentChannelType"].ToString();
                        worksheet.Cells[Fila, 17].Value = reader["FulfillmentOrderNumber"].ToString();
                        worksheet.Cells[Fila, 18].Value = reader["Quantity"].ToString();
                        worksheet.Cells[Fila, 19].Value = reader["Sku"].ToString();
                        worksheet.Cells[Fila, 20].Value = reader["Title"].ToString();
                        worksheet.Cells[Fila, 21].Value = reader["TotalCost"].ToString();
                        worksheet.Cells[Fila, 22].Value = reader["Commission"].ToString();
                        worksheet.Cells[Fila, 23].Value = reader["InventoryCost"].ToString();
                        worksheet.Cells[Fila, 24].Value = reader["UnitCost"].ToString();
                        worksheet.Cells[Fila, 25].Value = reader["ServiceCost"].ToString();
                        worksheet.Cells[Fila, 26].Value = reader["EstimatedShippingCost"].ToString();
                        worksheet.Cells[Fila, 27].Value = reader["ShippingCost"].ToString();
                        worksheet.Cells[Fila, 28].Value = reader["ShippingPrice"].ToString();
                        worksheet.Cells[Fila, 29].Value = reader["OverheadCost"].ToString();
                        worksheet.Cells[Fila, 30].Value = reader["PackageCost"].ToString();
                        worksheet.Cells[Fila, 31].Value = reader["ProfitLoss"].ToString();
                        worksheet.Cells[Fila, 32].Value = reader["Carrier"].ToString();
                        worksheet.Cells[Fila, 33].Value = reader["ShippingServiceLevel"].ToString();
                        worksheet.Cells[Fila, 34].Value = reader["ShippedByUser"].ToString();
                        worksheet.Cells[Fila, 35].Value = reader["ShippingWeight"].ToString();
                        worksheet.Cells[Fila, 36].Value = reader["Weight"].ToString();
                        worksheet.Cells[Fila, 37].Value = reader["Width"].ToString();
                        worksheet.Cells[Fila, 38].Value = reader["Length"].ToString();
                        worksheet.Cells[Fila, 39].Value = reader["Height"].ToString();
                        worksheet.Cells[Fila, 40].Value = reader["StateRegion"].ToString();
                        worksheet.Cells[Fila, 41].Value = reader["TrackingNum"].ToString();
                        worksheet.Cells[Fila, 42].Value = reader["MfrName"].ToString();
                        worksheet.Cells[Fila, 43].Value = reader["PricingRule"].ToString();
                        worksheet.Cells[Fila, 44].Value = reader["Ground Tracking ID Prefix"].ToString();
                        worksheet.Cells[Fila, 45].Value = reader["Express or Ground Tracking ID"].ToString();
                        worksheet.Cells[Fila, 46].Value = reader["Net Charge Amount"].ToString();
                        worksheet.Cells[Fila, 47].Value = reader["Service Type"].ToString();
                        worksheet.Cells[Fila, 48].Value = reader["Ground Service"].ToString();
                        worksheet.Cells[Fila, 49].Value = reader["Shipment Date"].ToString();
                        worksheet.Cells[Fila, 50].Value = reader["POD Delivery Date"].ToString();
                        worksheet.Cells[Fila, 51].Value = reader["Actual Weight Amount"].ToString();
                        worksheet.Cells[Fila, 52].Value = reader["Rated Weight Amount"].ToString();
                        worksheet.Cells[Fila, 53].Value = reader["Dim Length"].ToString();
                        worksheet.Cells[Fila, 54].Value = reader["Dim Width"].ToString();
                        worksheet.Cells[Fila, 55].Value = reader["Dim Height"].ToString();
                        worksheet.Cells[Fila, 56].Value = reader["Dim Divisor"].ToString();
                        worksheet.Cells[Fila, 57].Value = reader["Shipper State"].ToString();
                        worksheet.Cells[Fila, 58].Value = reader["Zone Code"].ToString();
                        worksheet.Cells[Fila, 59].Value = reader["Tendered Date"].ToString();
                        worksheet.Cells[Fila, 60].Value = reader["Earned Discount"].ToString();
                        worksheet.Cells[Fila, 61].Value = reader["Fuel Surcharge"].ToString();
                        worksheet.Cells[Fila, 62].Value = reader["Performance Pricing"].ToString();
                        worksheet.Cells[Fila, 63].Value = reader["Delivery Area Surcharge Extended"].ToString();
                        worksheet.Cells[Fila, 64].Value = reader["Delivery Area Surcharge"].ToString();
                        worksheet.Cells[Fila, 65].Value = reader["USPS Non-Mach Surcharge"].ToString();
                        worksheet.Cells[Fila, 66].Value = reader["Residential"].ToString();
                        worksheet.Cells[Fila, 67].Value = reader["Grace Discount"].ToString();
                        worksheet.Cells[Fila, 68].Value = reader["Declared Value"].ToString();
                        worksheet.Cells[Fila, 69].Value = reader["DAS Extended Resi"].ToString();
                        worksheet.Cells[Fila, 70].Value = reader["Additional Handling"].ToString();
                        worksheet.Cells[Fila, 71].Value = reader["Parcel Re-Label Charge"].ToString();
                        worksheet.Cells[Fila, 72].Value = reader["Indirect Signature"].ToString();
                        worksheet.Cells[Fila, 73].Value = reader["DAS Resi"].ToString();
                        worksheet.Cells[Fila, 74].Value = reader["Address Correction"].ToString();
                        worksheet.Cells[Fila, 75].Value = reader["DAS Extended Comm"].ToString();
                        worksheet.Cells[Fila, 76].Value = reader["Oversize Charge"].ToString();
                        worksheet.Cells[Fila, 77].Value = reader["AHS - Dimensions"].ToString();
                        worksheet.Cells[Fila, 78].Value = reader["Mail Class "].ToString();
                        worksheet.Cells[Fila, 79].Value = reader["Tracking Number "].ToString();
                        worksheet.Cells[Fila, 80].Value = reader["Total Postage Amt "].ToString();
                        worksheet.Cells[Fila, 81].Value = reader["Delivery Date "].ToString();
                        worksheet.Cells[Fila, 82].Value = reader["Weight "].ToString();
                        worksheet.Cells[Fila, 83].Value = reader["Zone "].ToString();
                        worksheet.Cells[Fila, 84].Value = reader["Service Type "].ToString();
                        worksheet.Cells[Fila, 85].Value = reader["Tracking Number "].ToString();
                        worksheet.Cells[Fila, 86].Value = reader["Net Charge Amount "].ToString();

                        // incrementa la fila a procesar
                        // -----------------------------
                        Fila++;
                    }
                }
                catch (Exception e)
                {
                    //Logguer.Log("Exception while creating ReportData:" + e.Message + "\t" + e.GetType());
                }
                finally
                {
                    reader.Close();
                    Conn.Close();
                }

                //Save your file
                FileInfo fi = new FileInfo(@"Path\To\Your\File.xlsx");
                excelPackage.SaveAs(fi);
            }
        }
    }
}
