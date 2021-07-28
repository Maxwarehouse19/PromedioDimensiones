using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

namespace PromedioDimensionesMaxwarehouse.Clases
{
    class ManejoBD
    {

        public void EliminaRegistroDW(SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "EliminaRegistrosDW";
            SqlCommand EliminaDW = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            EliminaDW.CommandType = CommandType.StoredProcedure;
            Conexion.Open();
            EliminaDW.ExecuteNonQuery();
            Conexion.Close();
        }

        public void EliminaRegistroFEDEX(SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "EliminaRegistrosFEDEX";
            SqlCommand EliminaDW = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            EliminaDW.CommandType = CommandType.StoredProcedure;
            Conexion.Open();
            EliminaDW.ExecuteNonQuery();
            Conexion.Close();
        }

        public void EliminaRegistroUSPS(SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "EliminaRegistrosUSPS";
            SqlCommand EliminaDW = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            EliminaDW.CommandType = CommandType.StoredProcedure;
            Conexion.Open();
            EliminaDW.ExecuteNonQuery();
            Conexion.Close();
        }

        public void EliminaRegistroUPS(SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "EliminaRegistrosUPS";
            SqlCommand EliminaDW = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            EliminaDW.CommandType = CommandType.StoredProcedure;
            Conexion.Open();
            EliminaDW.ExecuteNonQuery();
            Conexion.Close();
        }

        public void EliminaRegistroAmazon(SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "EliminaRegistrosAmazon";
            SqlCommand EliminaDW = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            EliminaDW.CommandType = CommandType.StoredProcedure;
            Conexion.Open();
            EliminaDW.ExecuteNonQuery();
            Conexion.Close();
        }

        public void EliminaRegistroBOX(SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "EliminaRegistrosBOX";
            SqlCommand EliminaDW = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            EliminaDW.CommandType = CommandType.StoredProcedure;
            Conexion.Open();
            EliminaDW.ExecuteNonQuery();
            Conexion.Close();
        }

        public void EliminaRegistroEJDDimensions(SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "EliminaRegistrosEJDDimensions";
            SqlCommand EliminaDW = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            EliminaDW.CommandType = CommandType.StoredProcedure;
            Conexion.Open();
            EliminaDW.ExecuteNonQuery();
            Conexion.Close();
        }

        public void EliminaRegistroJensenDimensions(SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "EliminaRegistrosJensenDimensions";
            SqlCommand EliminaDW = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            EliminaDW.CommandType = CommandType.StoredProcedure;
            Conexion.Open();
            EliminaDW.ExecuteNonQuery();
            Conexion.Close();
        }

        public void EliminaRegistroM15(SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "EliminaRegistrosMI15";
            SqlCommand EliminaDW = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            EliminaDW.CommandType = CommandType.StoredProcedure;
            Conexion.Open();
            EliminaDW.ExecuteNonQuery();
            Conexion.Close();
        }

        public void InsertaBDDW(PromedioDimensionesMaxwarehouse.Clases.ProcesarArchivoDW clsArchivoDW, SqlConnection Conexion)
        {

            string sqlAInsertaRegistroDW = "InsertaRegistroDW";
            SqlCommand GrabaDW = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            GrabaDW.CommandType = CommandType.StoredProcedure;
            GrabaDW.Parameters.Add("@SalesOrderNumber", SqlDbType.VarChar).Value = clsArchivoDW.SalesOrderNumber;
            GrabaDW.Parameters.Add("@HoldCode", SqlDbType.VarChar).Value = clsArchivoDW.HoldCode;
            GrabaDW.Parameters.Add("@TotalSales", SqlDbType.VarChar).Value = clsArchivoDW.TotalSales;
            GrabaDW.Parameters.Add("@SalesSku", SqlDbType.VarChar).Value = clsArchivoDW.SalesSku;
            GrabaDW.Parameters.Add("@SalesCategoryAtTimeOfSale", SqlDbType.VarChar).Value = clsArchivoDW.SalesCategoryAtTimeOfSale;
            GrabaDW.Parameters.Add("@UomCode", SqlDbType.VarChar).Value = clsArchivoDW.UomCode;
            GrabaDW.Parameters.Add("@UomQuantity", SqlDbType.VarChar).Value = clsArchivoDW.UomQuantity;
            GrabaDW.Parameters.Add("@SalesStatus", SqlDbType.VarChar).Value = clsArchivoDW.SalesStatus;
            GrabaDW.Parameters.Add("@SalesOrderDate", SqlDbType.VarChar).Value = clsArchivoDW.SalesOrderDate;
            GrabaDW.Parameters.Add("@SalesChannelName", SqlDbType.VarChar).Value = clsArchivoDW.SalesChannelName;
            GrabaDW.Parameters.Add("@CustomerName", SqlDbType.VarChar).Value = clsArchivoDW.CustomerName;
            GrabaDW.Parameters.Add("@FulfillmentSku", SqlDbType.VarChar).Value = clsArchivoDW.FulfillmentSku;
            GrabaDW.Parameters.Add("@FulfillmentChannelName", SqlDbType.VarChar).Value = clsArchivoDW.FulfillmentChannelName;
            GrabaDW.Parameters.Add("@FulfillmentChannelType", SqlDbType.VarChar).Value = clsArchivoDW.FulfillmentChannelType;
            GrabaDW.Parameters.Add("@LinkedFulfillmentChannelName", SqlDbType.VarChar).Value = clsArchivoDW.LinkedFulfillmentChannelName;
            GrabaDW.Parameters.Add("@FulfillmentLocationName", SqlDbType.VarChar).Value = clsArchivoDW.FulfillmentLocationName;
            GrabaDW.Parameters.Add("@FulfillmentOrderNumber", SqlDbType.VarChar).Value = clsArchivoDW.FulfillmentOrderNumber;
            GrabaDW.Parameters.Add("@Quantity", SqlDbType.VarChar).Value = clsArchivoDW.Quantity;
            GrabaDW.Parameters.Add("@Sku", SqlDbType.VarChar).Value = clsArchivoDW.Sku;
            GrabaDW.Parameters.Add("@Title", SqlDbType.VarChar).Value = clsArchivoDW.Title;
            GrabaDW.Parameters.Add("@TotalCost", SqlDbType.VarChar).Value = clsArchivoDW.TotalCost;
            GrabaDW.Parameters.Add("@Commission", SqlDbType.VarChar).Value = clsArchivoDW.Commission;
            GrabaDW.Parameters.Add("@InventoryCost", SqlDbType.VarChar).Value = clsArchivoDW.InventoryCost;
            GrabaDW.Parameters.Add("@UnitCost", SqlDbType.VarChar).Value = clsArchivoDW.UnitCost;
            GrabaDW.Parameters.Add("@ServiceCost", SqlDbType.VarChar).Value = clsArchivoDW.ServiceCost;
            GrabaDW.Parameters.Add("@EstimatedShippingCost", SqlDbType.VarChar).Value = clsArchivoDW.EstimatedShippingCost;
            GrabaDW.Parameters.Add("@ShippingCost", SqlDbType.VarChar).Value = clsArchivoDW.ShippingCost;
            GrabaDW.Parameters.Add("@ShippingPrice", SqlDbType.VarChar).Value = clsArchivoDW.ShippingPrice;
            GrabaDW.Parameters.Add("@OverheadCost", SqlDbType.VarChar).Value = clsArchivoDW.OverheadCost;
            GrabaDW.Parameters.Add("@PackageCost", SqlDbType.VarChar).Value = clsArchivoDW.PackageCost;
            GrabaDW.Parameters.Add("@ProfitLoss", SqlDbType.VarChar).Value = clsArchivoDW.ProfitLoss;
            GrabaDW.Parameters.Add("@Carrier", SqlDbType.VarChar).Value = clsArchivoDW.Carrier;
            GrabaDW.Parameters.Add("@ShippingServiceLevel", SqlDbType.VarChar).Value = clsArchivoDW.ShippingServiceLevel;
            GrabaDW.Parameters.Add("@ShippedByUser", SqlDbType.VarChar).Value = clsArchivoDW.ShippedByUser;
            GrabaDW.Parameters.Add("@ShippingWeight", SqlDbType.VarChar).Value = clsArchivoDW.ShippingWeight;
            GrabaDW.Parameters.Add("@Length", SqlDbType.VarChar).Value = clsArchivoDW.Length;
            GrabaDW.Parameters.Add("@Width", SqlDbType.VarChar).Value = clsArchivoDW.varWidth;
            GrabaDW.Parameters.Add("@Height", SqlDbType.VarChar).Value = clsArchivoDW.varHeight;
            GrabaDW.Parameters.Add("@Weight", SqlDbType.VarChar).Value = clsArchivoDW.Weight;
            GrabaDW.Parameters.Add("@StateRegion", SqlDbType.VarChar).Value = clsArchivoDW.StateRegion;
            GrabaDW.Parameters.Add("@TrackingNum", SqlDbType.VarChar).Value = clsArchivoDW.TrackingNum;
            GrabaDW.Parameters.Add("@MfrName", SqlDbType.VarChar).Value = clsArchivoDW.MfrName;
            GrabaDW.Parameters.Add("@PricingRule", SqlDbType.VarChar).Value = clsArchivoDW.PricingRule;
            GrabaDW.Parameters.Add("@ActualShippingCost", SqlDbType.VarChar).Value = clsArchivoDW.ActualShippingCost;
            GrabaDW.Parameters.Add("@ActualShipping", SqlDbType.VarChar).Value = clsArchivoDW.ActualShipping;
            GrabaDW.Parameters.Add("@ShippingCostDifference", SqlDbType.VarChar).Value = clsArchivoDW.ShippingCostDifference;

            Conexion.Open();
            GrabaDW.ExecuteNonQuery();
            Conexion.Close();
        }

        public void InsertaBDFEDEX(PromedioDimensiones.PedidoFedex clsPedido, SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "InsertaRegistroFEDEX";
            SqlCommand commandObtieneCuentasFedex = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            commandObtieneCuentasFedex.CommandType = CommandType.StoredProcedure;
            commandObtieneCuentasFedex.Parameters.Add("@BilltoAccountNumber", SqlDbType.VarChar).Value = clsPedido.BilltoAccountNumber;
            commandObtieneCuentasFedex.Parameters.Add("@InvoiceDate", SqlDbType.VarChar).Value = clsPedido.InvoiceDate;
            commandObtieneCuentasFedex.Parameters.Add("@InvoiceNumber", SqlDbType.VarChar).Value = clsPedido.InvoiceNumber;
            commandObtieneCuentasFedex.Parameters.Add("@StoreID", SqlDbType.VarChar).Value = clsPedido.StoreID;
            commandObtieneCuentasFedex.Parameters.Add("@OriginalAmountDue", SqlDbType.VarChar).Value = clsPedido.OriginalAmountDue;
            commandObtieneCuentasFedex.Parameters.Add("@CurrentBalance", SqlDbType.VarChar).Value = clsPedido.CurrentBalance;
            commandObtieneCuentasFedex.Parameters.Add("@Payor", SqlDbType.VarChar).Value = clsPedido.Payor;
            commandObtieneCuentasFedex.Parameters.Add("@GroundTrackingIDPrefix", SqlDbType.VarChar).Value = clsPedido.GroundTrackingIDPrefix;
            commandObtieneCuentasFedex.Parameters.Add("@ExpressorGroundTrackingID", SqlDbType.VarChar).Value = clsPedido.ExpressorGroundTrackingID;
            commandObtieneCuentasFedex.Parameters.Add("@TransportationChargeAmount", SqlDbType.VarChar).Value = clsPedido.TransportationChargeAmount;
            commandObtieneCuentasFedex.Parameters.Add("@NetChargeAmount", SqlDbType.VarChar).Value = clsPedido.NetChargeAmount;
            commandObtieneCuentasFedex.Parameters.Add("@ServiceType", SqlDbType.VarChar).Value = clsPedido.ServiceType;
            commandObtieneCuentasFedex.Parameters.Add("@GroundService", SqlDbType.VarChar).Value = clsPedido.GroundService;
            commandObtieneCuentasFedex.Parameters.Add("@ShipmentDate", SqlDbType.VarChar).Value = clsPedido.ShipmentDate;
            commandObtieneCuentasFedex.Parameters.Add("@PODDeliveryDate", SqlDbType.VarChar).Value = clsPedido.PODDeliveryDate;
            commandObtieneCuentasFedex.Parameters.Add("@PODDeliveryTime", SqlDbType.VarChar).Value = clsPedido.PODDeliveryTime;
            commandObtieneCuentasFedex.Parameters.Add("@PODServiceAreaCode", SqlDbType.VarChar).Value = clsPedido.PODServiceAreaCode;
            commandObtieneCuentasFedex.Parameters.Add("@PODSignatureDescription", SqlDbType.VarChar).Value = clsPedido.PODSignatureDescription;
            commandObtieneCuentasFedex.Parameters.Add("@ActualWeightAmount", SqlDbType.VarChar).Value = clsPedido.ActualWeightAmount;
            commandObtieneCuentasFedex.Parameters.Add("@ActualWeightUnits", SqlDbType.VarChar).Value = clsPedido.ActualWeightUnits;
            commandObtieneCuentasFedex.Parameters.Add("@RatedWeightAmount", SqlDbType.VarChar).Value = clsPedido.RatedWeightAmount;
            commandObtieneCuentasFedex.Parameters.Add("@RatedWeightUnits", SqlDbType.VarChar).Value = clsPedido.RatedWeightUnits;
            commandObtieneCuentasFedex.Parameters.Add("@NumberofPieces", SqlDbType.VarChar).Value = clsPedido.NumberofPieces;
            commandObtieneCuentasFedex.Parameters.Add("@BundleNumber", SqlDbType.VarChar).Value = clsPedido.BundleNumber;
            commandObtieneCuentasFedex.Parameters.Add("@MeterNumber", SqlDbType.VarChar).Value = clsPedido.MeterNumber;
            commandObtieneCuentasFedex.Parameters.Add("@TDMasterTrackingID", SqlDbType.VarChar).Value = clsPedido.TDMasterTrackingID;
            commandObtieneCuentasFedex.Parameters.Add("@ServicePackaging", SqlDbType.VarChar).Value = clsPedido.ServicePackaging;
            commandObtieneCuentasFedex.Parameters.Add("@DimLength", SqlDbType.VarChar).Value = clsPedido.DimLength;
            commandObtieneCuentasFedex.Parameters.Add("@DimWidth", SqlDbType.VarChar).Value = clsPedido.DimWidth;
            commandObtieneCuentasFedex.Parameters.Add("@DimHeight", SqlDbType.VarChar).Value = clsPedido.DimHeight;
            commandObtieneCuentasFedex.Parameters.Add("@DimDivisor", SqlDbType.VarChar).Value = clsPedido.DimDivisor;
            commandObtieneCuentasFedex.Parameters.Add("@DimUnit", SqlDbType.VarChar).Value = clsPedido.DimUnit;
            commandObtieneCuentasFedex.Parameters.Add("@RecipientName", SqlDbType.VarChar).Value = clsPedido.RecipientName;
            commandObtieneCuentasFedex.Parameters.Add("@RecipientCompany", SqlDbType.VarChar).Value = clsPedido.RecipientCompany;
            commandObtieneCuentasFedex.Parameters.Add("@RecipientAddressLine1", SqlDbType.VarChar).Value = clsPedido.RecipientAddressLine1;
            commandObtieneCuentasFedex.Parameters.Add("@RecipientAddressLine2", SqlDbType.VarChar).Value = clsPedido.RecipientAddressLine2;
            commandObtieneCuentasFedex.Parameters.Add("@RecipientCity", SqlDbType.VarChar).Value = clsPedido.RecipientCity;
            commandObtieneCuentasFedex.Parameters.Add("@RecipientState", SqlDbType.VarChar).Value = clsPedido.RecipientState;
            commandObtieneCuentasFedex.Parameters.Add("@RecipientZipCode", SqlDbType.VarChar).Value = clsPedido.RecipientZipCode;
            commandObtieneCuentasFedex.Parameters.Add("@RecipientCountryTerritory", SqlDbType.VarChar).Value = clsPedido.RecipientCountryTerritory;
            commandObtieneCuentasFedex.Parameters.Add("@ShipperCompany", SqlDbType.VarChar).Value = clsPedido.ShipperCompany;
            commandObtieneCuentasFedex.Parameters.Add("@ShipperName", SqlDbType.VarChar).Value = clsPedido.ShipperName;
            commandObtieneCuentasFedex.Parameters.Add("@ShipperAddressLine1", SqlDbType.VarChar).Value = clsPedido.ShipperAddressLine1;
            commandObtieneCuentasFedex.Parameters.Add("@ShipperAddressLine2", SqlDbType.VarChar).Value = clsPedido.ShipperAddressLine2;
            commandObtieneCuentasFedex.Parameters.Add("@ShipperCity", SqlDbType.VarChar).Value = clsPedido.ShipperCity;
            commandObtieneCuentasFedex.Parameters.Add("@ShipperState", SqlDbType.VarChar).Value = clsPedido.ShipperState;
            commandObtieneCuentasFedex.Parameters.Add("@ShipperZipCode", SqlDbType.VarChar).Value = clsPedido.ShipperZipCode;
            commandObtieneCuentasFedex.Parameters.Add("@ShipperCountryTerritory", SqlDbType.VarChar).Value = clsPedido.ShipperCountryTerritory;
            commandObtieneCuentasFedex.Parameters.Add("@OriginalCustomerReference", SqlDbType.VarChar).Value = clsPedido.OriginalCustomerReference;
            commandObtieneCuentasFedex.Parameters.Add("@OriginalRef2", SqlDbType.VarChar).Value = clsPedido.OriginalRef2;
            commandObtieneCuentasFedex.Parameters.Add("@OriginalRef3PONumber", SqlDbType.VarChar).Value = clsPedido.OriginalRef3PONumber;
            commandObtieneCuentasFedex.Parameters.Add("@OriginalDepartmentReferenceDescription", SqlDbType.VarChar).Value = clsPedido.OriginalDepartmentReferenceDescription;
            commandObtieneCuentasFedex.Parameters.Add("@UpdatedCustomerReference", SqlDbType.VarChar).Value = clsPedido.UpdatedCustomerReference;
            commandObtieneCuentasFedex.Parameters.Add("@UpdatedRef2", SqlDbType.VarChar).Value = clsPedido.UpdatedRef2;
            commandObtieneCuentasFedex.Parameters.Add("@UpdatedRef3PONumber", SqlDbType.VarChar).Value = clsPedido.UpdatedRef3PONumber;
            commandObtieneCuentasFedex.Parameters.Add("@UpdatedDepartmentReferenceDescription", SqlDbType.VarChar).Value = clsPedido.UpdatedDepartmentReferenceDescription;
            commandObtieneCuentasFedex.Parameters.Add("@RMA", SqlDbType.VarChar).Value = clsPedido.RMA;
            commandObtieneCuentasFedex.Parameters.Add("@OriginalRecipientAddressLine1", SqlDbType.VarChar).Value = clsPedido.OriginalRecipientAddressLine1;
            commandObtieneCuentasFedex.Parameters.Add("@OriginalRecipientAddressLine2", SqlDbType.VarChar).Value = clsPedido.OriginalRecipientAddressLine2;
            commandObtieneCuentasFedex.Parameters.Add("@OriginalRecipientCity", SqlDbType.VarChar).Value = clsPedido.OriginalRecipientCity;
            commandObtieneCuentasFedex.Parameters.Add("@OriginalRecipientState", SqlDbType.VarChar).Value = clsPedido.OriginalRecipientState;
            commandObtieneCuentasFedex.Parameters.Add("@OriginalRecipientZipCode", SqlDbType.VarChar).Value = clsPedido.OriginalRecipientZipCode;
            commandObtieneCuentasFedex.Parameters.Add("@OriginalRecipientCountryTerritory", SqlDbType.VarChar).Value = clsPedido.OriginalRecipientCountryTerritory;
            commandObtieneCuentasFedex.Parameters.Add("@ZoneCode", SqlDbType.VarChar).Value = clsPedido.ZoneCode;
            commandObtieneCuentasFedex.Parameters.Add("@CostAllocation", SqlDbType.VarChar).Value = clsPedido.CostAllocation;
            commandObtieneCuentasFedex.Parameters.Add("@AlternateAddressLine1", SqlDbType.VarChar).Value = clsPedido.AlternateAddressLine1;
            commandObtieneCuentasFedex.Parameters.Add("@AlternateAddressLine2", SqlDbType.VarChar).Value = clsPedido.AlternateAddressLine2;
            commandObtieneCuentasFedex.Parameters.Add("@AlternateCity", SqlDbType.VarChar).Value = clsPedido.AlternateCity;
            commandObtieneCuentasFedex.Parameters.Add("@AlternateStateProvince", SqlDbType.VarChar).Value = clsPedido.AlternateStateProvince;
            commandObtieneCuentasFedex.Parameters.Add("@AlternateZipCode", SqlDbType.VarChar).Value = clsPedido.AlternateZipCode;
            commandObtieneCuentasFedex.Parameters.Add("@AlternateCountryTerritoryCode", SqlDbType.VarChar).Value = clsPedido.AlternateCountryTerritoryCode;
            commandObtieneCuentasFedex.Parameters.Add("@CrossRefTrackingIDPrefix", SqlDbType.VarChar).Value = clsPedido.CrossRefTrackingIDPrefix;
            commandObtieneCuentasFedex.Parameters.Add("@CrossRefTrackingID", SqlDbType.VarChar).Value = clsPedido.CrossRefTrackingID;
            commandObtieneCuentasFedex.Parameters.Add("@EntryDate", SqlDbType.VarChar).Value = clsPedido.EntryDate;
            commandObtieneCuentasFedex.Parameters.Add("@EntryNumber", SqlDbType.VarChar).Value = clsPedido.EntryNumber;
            commandObtieneCuentasFedex.Parameters.Add("@CustomsValue", SqlDbType.VarChar).Value = clsPedido.CustomsValue;
            commandObtieneCuentasFedex.Parameters.Add("@CustomsValueCurrencyCode", SqlDbType.VarChar).Value = clsPedido.CustomsValueCurrencyCode;
            commandObtieneCuentasFedex.Parameters.Add("@DeclaredValue", SqlDbType.VarChar).Value = clsPedido.DeclaredValue;
            commandObtieneCuentasFedex.Parameters.Add("@DeclaredValueCurrencyCode", SqlDbType.VarChar).Value = clsPedido.DeclaredValueCurrencyCode;
            commandObtieneCuentasFedex.Parameters.Add("@CommodityDescription", SqlDbType.VarChar).Value = clsPedido.CommodityDescription;
            commandObtieneCuentasFedex.Parameters.Add("@CommodityCountryTerritoryCode", SqlDbType.VarChar).Value = clsPedido.CommodityCountryTerritoryCode;
            commandObtieneCuentasFedex.Parameters.Add("@CommodityDescription1", SqlDbType.VarChar).Value = clsPedido.CommodityDescription1;
            commandObtieneCuentasFedex.Parameters.Add("@CommodityCountryTerritoryCode1", SqlDbType.VarChar).Value = clsPedido.CommodityCountryTerritoryCode1;
            commandObtieneCuentasFedex.Parameters.Add("@CommodityDescription2", SqlDbType.VarChar).Value = clsPedido.CommodityDescription2;
            commandObtieneCuentasFedex.Parameters.Add("@CommodityCountryTerritoryCode2", SqlDbType.VarChar).Value = clsPedido.CommodityCountryTerritoryCode2;
            commandObtieneCuentasFedex.Parameters.Add("@CommodityDescription3", SqlDbType.VarChar).Value = clsPedido.CommodityDescription3;
            commandObtieneCuentasFedex.Parameters.Add("@CommodityCountryTerritoryCode3", SqlDbType.VarChar).Value = clsPedido.CommodityCountryTerritoryCode3;
            commandObtieneCuentasFedex.Parameters.Add("@CurrencyConversionDate", SqlDbType.VarChar).Value = clsPedido.CurrencyConversionDate;
            commandObtieneCuentasFedex.Parameters.Add("@CurrencyConversionRate", SqlDbType.VarChar).Value = clsPedido.CurrencyConversionRate;
            commandObtieneCuentasFedex.Parameters.Add("@MultiweightNumber", SqlDbType.VarChar).Value = clsPedido.MultiweightNumber;
            commandObtieneCuentasFedex.Parameters.Add("@MultiweightTotalMultiweightUnits", SqlDbType.VarChar).Value = clsPedido.MultiweightTotalMultiweightUnits;
            commandObtieneCuentasFedex.Parameters.Add("@MultiweightTotalMultiweightWeight", SqlDbType.VarChar).Value = clsPedido.MultiweightTotalMultiweightWeight;
            commandObtieneCuentasFedex.Parameters.Add("@MultiweightTotalShipmentChargeAmount", SqlDbType.VarChar).Value = clsPedido.MultiweightTotalShipmentChargeAmount;
            commandObtieneCuentasFedex.Parameters.Add("@MultiweightTotalShipmentWeight", SqlDbType.VarChar).Value = clsPedido.MultiweightTotalShipmentWeight;
            commandObtieneCuentasFedex.Parameters.Add("@GroundTrackingIDAddressCorrectionDiscountChargeAmount", SqlDbType.VarChar).Value = clsPedido.GroundTrackingIDAddressCorrectionDiscountChargeAmount;
            commandObtieneCuentasFedex.Parameters.Add("@GroundTrackingIDAddressCorrectionGrossChargeAmount", SqlDbType.VarChar).Value = clsPedido.GroundTrackingIDAddressCorrectionGrossChargeAmount;
            commandObtieneCuentasFedex.Parameters.Add("@RatedMethod", SqlDbType.VarChar).Value = clsPedido.RatedMethod;
            commandObtieneCuentasFedex.Parameters.Add("@SortHub", SqlDbType.VarChar).Value = clsPedido.SortHub;
            commandObtieneCuentasFedex.Parameters.Add("@EstimatedWeight", SqlDbType.VarChar).Value = clsPedido.EstimatedWeight;
            commandObtieneCuentasFedex.Parameters.Add("@EstimatedWeightUnit", SqlDbType.VarChar).Value = clsPedido.EstimatedWeightUnit;
            commandObtieneCuentasFedex.Parameters.Add("@PostalClass", SqlDbType.VarChar).Value = clsPedido.PostalClass;
            commandObtieneCuentasFedex.Parameters.Add("@ProcessCategory", SqlDbType.VarChar).Value = clsPedido.ProcessCategory;
            commandObtieneCuentasFedex.Parameters.Add("@PackageSize", SqlDbType.VarChar).Value = clsPedido.PackageSize;
            commandObtieneCuentasFedex.Parameters.Add("@DeliveryConfirmation", SqlDbType.VarChar).Value = clsPedido.DeliveryConfirmation;
            commandObtieneCuentasFedex.Parameters.Add("@TenderedDate", SqlDbType.VarChar).Value = clsPedido.TenderedDate;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription1", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription1;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount1", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount1;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription2", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription2;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount2", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount2;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription3", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription3;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount3", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount3;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription4", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription4;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount4", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount4;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription5", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription5;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount5", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount5;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription6", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription6;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount6", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount6;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription7", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription7;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount7", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount7;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription8", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription8;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount8", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount8;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription9", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription9;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount9", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount9;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription10", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription10;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount10", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount10;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription11", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription11;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount11", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount11;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription12", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription12;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount12", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount12;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription13", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription13;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount13", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount13;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription14", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription14;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount14", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount14;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription15", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription15;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount15", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount15;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription16", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription16;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount16", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount16;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription17", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription17;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount17", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount17;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription18", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription18;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount18", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount18;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription19", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription19;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount19", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount19;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription20", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription20;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount20", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount20;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription21", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription21;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount21", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount21;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription22", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription22;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount22", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount22;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription23", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription23;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount23", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount23;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeDescription24", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeDescription24;
            commandObtieneCuentasFedex.Parameters.Add("@TrackingIDChargeAmount24", SqlDbType.VarChar).Value = clsPedido.TrackingIDChargeAmount24;
            commandObtieneCuentasFedex.Parameters.Add("@ShipmentNotes", SqlDbType.VarChar).Value = clsPedido.ShipmentNotes;

            Conexion.Open();
            commandObtieneCuentasFedex.ExecuteNonQuery();
            Conexion.Close();
        }

        public void InsertaBDUSPS(PromedioDimensiones.PedidoUSPS clsPedido, SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "InsertaRegistroUSPS";
            SqlCommand commandObtieneUSPS = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            commandObtieneUSPS.CommandType = CommandType.StoredProcedure;
            commandObtieneUSPS.Parameters.Add("@AccountNumber", SqlDbType.VarChar).Value = clsPedido.AccountNumber;
            commandObtieneUSPS.Parameters.Add("@ID", SqlDbType.VarChar).Value = clsPedido.ID;
            commandObtieneUSPS.Parameters.Add("@DateTime", SqlDbType.VarChar).Value = clsPedido.DateTime;
            commandObtieneUSPS.Parameters.Add("@Postmark", SqlDbType.VarChar).Value = clsPedido.Postmark;
            commandObtieneUSPS.Parameters.Add("@Origin", SqlDbType.VarChar).Value = clsPedido.Origin;
            commandObtieneUSPS.Parameters.Add("@Destination", SqlDbType.VarChar).Value = clsPedido.Destination;
            commandObtieneUSPS.Parameters.Add("@Type", SqlDbType.VarChar).Value = clsPedido.Type;
            commandObtieneUSPS.Parameters.Add("@MailClass", SqlDbType.VarChar).Value = clsPedido.MailClass;
            commandObtieneUSPS.Parameters.Add("@TrackingNumber", SqlDbType.VarChar).Value = clsPedido.TrackingNumber;
            commandObtieneUSPS.Parameters.Add("@DeclaredValue", SqlDbType.VarChar).Value = clsPedido.DeclaredValue;
            commandObtieneUSPS.Parameters.Add("@TotalPostageAmt", SqlDbType.VarChar).Value = clsPedido.TotalPostageAmt;
            commandObtieneUSPS.Parameters.Add("@Balance", SqlDbType.VarChar).Value = clsPedido.Balance;
            commandObtieneUSPS.Parameters.Add("@RefundStatus", SqlDbType.VarChar).Value = clsPedido.RefundStatus;
            commandObtieneUSPS.Parameters.Add("@GroupCode", SqlDbType.VarChar).Value = clsPedido.GroupCode;
            commandObtieneUSPS.Parameters.Add("@ReferenceID", SqlDbType.VarChar).Value = clsPedido.ReferenceID;
            commandObtieneUSPS.Parameters.Add("@DeliveryDate", SqlDbType.VarChar).Value = clsPedido.DeliveryDate;
            commandObtieneUSPS.Parameters.Add("@StatusCode", SqlDbType.VarChar).Value = clsPedido.StatusCode;
            commandObtieneUSPS.Parameters.Add("@StatusDescription", SqlDbType.VarChar).Value = clsPedido.StatusDescription;
            commandObtieneUSPS.Parameters.Add("@Weight", SqlDbType.VarChar).Value = clsPedido.Weight;
            commandObtieneUSPS.Parameters.Add("@OptionalServices", SqlDbType.VarChar).Value = clsPedido.OptionalServices;
            commandObtieneUSPS.Parameters.Add("@DestinationName", SqlDbType.VarChar).Value = clsPedido.DestinationName;
            commandObtieneUSPS.Parameters.Add("@DestinationCompanyName", SqlDbType.VarChar).Value = clsPedido.DestinationCompanyName;
            commandObtieneUSPS.Parameters.Add("@DestinationAddress", SqlDbType.VarChar).Value = clsPedido.DestinationAddress;
            commandObtieneUSPS.Parameters.Add("@DestinationCity", SqlDbType.VarChar).Value = clsPedido.DestinationCity;
            commandObtieneUSPS.Parameters.Add("@DestinationState", SqlDbType.VarChar).Value = clsPedido.DestinationState;
            commandObtieneUSPS.Parameters.Add("@DestinationZip", SqlDbType.VarChar).Value = clsPedido.DestinationZip;
            commandObtieneUSPS.Parameters.Add("@DestinationCountry", SqlDbType.VarChar).Value = clsPedido.DestinationCountry;
            commandObtieneUSPS.Parameters.Add("@Phone", SqlDbType.VarChar).Value = clsPedido.Phone;
            commandObtieneUSPS.Parameters.Add("@Email", SqlDbType.VarChar).Value = clsPedido.Email;
            commandObtieneUSPS.Parameters.Add("@Reference2", SqlDbType.VarChar).Value = clsPedido.Reference2;
            commandObtieneUSPS.Parameters.Add("@Reference3", SqlDbType.VarChar).Value = clsPedido.Reference3;
            commandObtieneUSPS.Parameters.Add("@Reference4", SqlDbType.VarChar).Value = clsPedido.Reference4;
            commandObtieneUSPS.Parameters.Add("@PackageDescription", SqlDbType.VarChar).Value = clsPedido.PackageDescription;
            commandObtieneUSPS.Parameters.Add("@Zone", SqlDbType.VarChar).Value = clsPedido.Zone;
            commandObtieneUSPS.Parameters.Add("@IsCubic", SqlDbType.VarChar).Value = clsPedido.IsCubic;
            commandObtieneUSPS.Parameters.Add("@CubicValue", SqlDbType.VarChar).Value = clsPedido.CubicValue;
            commandObtieneUSPS.Parameters.Add("@AdjWeight", SqlDbType.VarChar).Value = clsPedido.AdjWeight;
            commandObtieneUSPS.Parameters.Add("@AdjDimensions", SqlDbType.VarChar).Value = clsPedido.AdjDimensions;
            commandObtieneUSPS.Parameters.Add("@AdjFromZIP", SqlDbType.VarChar).Value = clsPedido.AdjFromZIP;
            commandObtieneUSPS.Parameters.Add("@AdjToZIP", SqlDbType.VarChar).Value = clsPedido.AdjToZIP;
            commandObtieneUSPS.Parameters.Add("@AdjMailClass", SqlDbType.VarChar).Value = clsPedido.AdjMailClass;

            Conexion.Open();
            commandObtieneUSPS.ExecuteNonQuery();
            Conexion.Close();
        }

        public void InsertaBDUPS(PromedioDimensiones.PedidoUPS clsPedido, SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "InsertaRegistroUPS";
            SqlCommand commandObtieneUPS = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            commandObtieneUPS.CommandType = CommandType.StoredProcedure;
            commandObtieneUPS.Parameters.Add("@F1", SqlDbType.VarChar).Value = clsPedido.Campo1;
            commandObtieneUPS.Parameters.Add("@F2", SqlDbType.VarChar).Value = clsPedido.Campo2;
            commandObtieneUPS.Parameters.Add("@F3", SqlDbType.VarChar).Value = clsPedido.Campo3;
            commandObtieneUPS.Parameters.Add("@F4", SqlDbType.VarChar).Value = clsPedido.Campo4;
            commandObtieneUPS.Parameters.Add("@F5", SqlDbType.VarChar).Value = clsPedido.Campo5;
            commandObtieneUPS.Parameters.Add("@F6", SqlDbType.VarChar).Value = clsPedido.Campo6;
            commandObtieneUPS.Parameters.Add("@F7", SqlDbType.VarChar).Value = clsPedido.Campo7;
            commandObtieneUPS.Parameters.Add("@F8", SqlDbType.VarChar).Value = clsPedido.Campo8;
            commandObtieneUPS.Parameters.Add("@F9", SqlDbType.VarChar).Value = clsPedido.Campo9;
            commandObtieneUPS.Parameters.Add("@F10", SqlDbType.VarChar).Value = clsPedido.Campo10;
            commandObtieneUPS.Parameters.Add("@F11", SqlDbType.VarChar).Value = clsPedido.Campo11;
            commandObtieneUPS.Parameters.Add("@F12", SqlDbType.VarChar).Value = clsPedido.Campo12;
            commandObtieneUPS.Parameters.Add("@F13", SqlDbType.VarChar).Value = clsPedido.Campo13;
            commandObtieneUPS.Parameters.Add("@F14", SqlDbType.VarChar).Value = clsPedido.Campo14;
            commandObtieneUPS.Parameters.Add("@F15", SqlDbType.VarChar).Value = clsPedido.Campo15;
            commandObtieneUPS.Parameters.Add("@F16", SqlDbType.VarChar).Value = clsPedido.Campo16;
            commandObtieneUPS.Parameters.Add("@F17", SqlDbType.VarChar).Value = clsPedido.Campo17;
            commandObtieneUPS.Parameters.Add("@F18", SqlDbType.VarChar).Value = clsPedido.Campo18;
            commandObtieneUPS.Parameters.Add("@F19", SqlDbType.VarChar).Value = clsPedido.Campo19;
            commandObtieneUPS.Parameters.Add("@F20", SqlDbType.VarChar).Value = clsPedido.Campo20;
            commandObtieneUPS.Parameters.Add("@F21", SqlDbType.VarChar).Value = clsPedido.Campo21;
            commandObtieneUPS.Parameters.Add("@F22", SqlDbType.VarChar).Value = clsPedido.Campo22;
            commandObtieneUPS.Parameters.Add("@F23", SqlDbType.VarChar).Value = clsPedido.Campo23;
            commandObtieneUPS.Parameters.Add("@F24", SqlDbType.VarChar).Value = clsPedido.Campo24;
            commandObtieneUPS.Parameters.Add("@F25", SqlDbType.VarChar).Value = clsPedido.Campo25;
            commandObtieneUPS.Parameters.Add("@F26", SqlDbType.VarChar).Value = clsPedido.Campo26;
            commandObtieneUPS.Parameters.Add("@F27", SqlDbType.VarChar).Value = clsPedido.Campo27;
            commandObtieneUPS.Parameters.Add("@F28", SqlDbType.VarChar).Value = clsPedido.Campo28;
            commandObtieneUPS.Parameters.Add("@F29", SqlDbType.VarChar).Value = clsPedido.Campo29;
            commandObtieneUPS.Parameters.Add("@F30", SqlDbType.VarChar).Value = clsPedido.Campo30;
            commandObtieneUPS.Parameters.Add("@F31", SqlDbType.VarChar).Value = clsPedido.Campo31;
            commandObtieneUPS.Parameters.Add("@F32", SqlDbType.VarChar).Value = clsPedido.Campo32;
            commandObtieneUPS.Parameters.Add("@F33", SqlDbType.VarChar).Value = clsPedido.Campo33;
            commandObtieneUPS.Parameters.Add("@F34", SqlDbType.VarChar).Value = clsPedido.Campo34;
            commandObtieneUPS.Parameters.Add("@F35", SqlDbType.VarChar).Value = clsPedido.Campo35;
            commandObtieneUPS.Parameters.Add("@F36", SqlDbType.VarChar).Value = clsPedido.Campo36;
            commandObtieneUPS.Parameters.Add("@F37", SqlDbType.VarChar).Value = clsPedido.Campo37;
            commandObtieneUPS.Parameters.Add("@F38", SqlDbType.VarChar).Value = clsPedido.Campo38;
            commandObtieneUPS.Parameters.Add("@F39", SqlDbType.VarChar).Value = clsPedido.Campo39;
            commandObtieneUPS.Parameters.Add("@F40", SqlDbType.VarChar).Value = clsPedido.Campo40;
            commandObtieneUPS.Parameters.Add("@F41", SqlDbType.VarChar).Value = clsPedido.Campo41;
            commandObtieneUPS.Parameters.Add("@F42", SqlDbType.VarChar).Value = clsPedido.Campo42;

            Conexion.Open();
            commandObtieneUPS.ExecuteNonQuery();
            Conexion.Close();
        }

        public void InsertaBDAMAZON(PromedioDimensiones.PedidoAmazon clsPedido, SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "InsertaRegistroAMAZON";
            SqlCommand commandObtieneAmazon = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            commandObtieneAmazon.CommandType = CommandType.StoredProcedure;

            commandObtieneAmazon.Parameters.Add("@datetime", SqlDbType.VarChar).Value = clsPedido.datetime;
            commandObtieneAmazon.Parameters.Add("@settlementid", SqlDbType.VarChar).Value = clsPedido.settlementid;
            commandObtieneAmazon.Parameters.Add("@type", SqlDbType.VarChar).Value = clsPedido.type;
            commandObtieneAmazon.Parameters.Add("@orderid", SqlDbType.VarChar).Value = clsPedido.orderid;
            commandObtieneAmazon.Parameters.Add("@sku", SqlDbType.VarChar).Value = clsPedido.sku;
            commandObtieneAmazon.Parameters.Add("@description", SqlDbType.VarChar).Value = clsPedido.description;
            commandObtieneAmazon.Parameters.Add("@quantity", SqlDbType.VarChar).Value = clsPedido.quantity;
            commandObtieneAmazon.Parameters.Add("@marketplace", SqlDbType.VarChar).Value = clsPedido.marketplace;
            commandObtieneAmazon.Parameters.Add("@fulfillment", SqlDbType.VarChar).Value = clsPedido.fulfillment;
            commandObtieneAmazon.Parameters.Add("@ordercity", SqlDbType.VarChar).Value = clsPedido.ordercity;
            commandObtieneAmazon.Parameters.Add("@orderstate", SqlDbType.VarChar).Value = clsPedido.orderstate;
            commandObtieneAmazon.Parameters.Add("@orderpostal", SqlDbType.VarChar).Value = clsPedido.orderpostal;
            commandObtieneAmazon.Parameters.Add("@taxcollectionmodel", SqlDbType.VarChar).Value = clsPedido.taxcollectionmodel;
            commandObtieneAmazon.Parameters.Add("@productsales", SqlDbType.VarChar).Value = clsPedido.productsales;
            commandObtieneAmazon.Parameters.Add("@productsalestax", SqlDbType.VarChar).Value = clsPedido.productsalestax;
            commandObtieneAmazon.Parameters.Add("@shippingcredits", SqlDbType.VarChar).Value = clsPedido.shippingcredits;
            commandObtieneAmazon.Parameters.Add("@shippingcreditstax", SqlDbType.VarChar).Value = clsPedido.shippingcreditstax;
            commandObtieneAmazon.Parameters.Add("@giftwrapcredits", SqlDbType.VarChar).Value = clsPedido.giftwrapcredits;
            commandObtieneAmazon.Parameters.Add("@giftwrapcreditstax", SqlDbType.VarChar).Value = clsPedido.giftwrapcreditstax;
            commandObtieneAmazon.Parameters.Add("@promotionalrebates", SqlDbType.VarChar).Value = clsPedido.promotionalrebates;
            commandObtieneAmazon.Parameters.Add("@promotionalrebatestax", SqlDbType.VarChar).Value = clsPedido.promotionalrebatestax;
            commandObtieneAmazon.Parameters.Add("@marketplacewithheldtax", SqlDbType.VarChar).Value = clsPedido.marketplacewithheldtax;
            commandObtieneAmazon.Parameters.Add("@sellingfees", SqlDbType.VarChar).Value = clsPedido.sellingfees;
            commandObtieneAmazon.Parameters.Add("@fbafees", SqlDbType.VarChar).Value = clsPedido.fbafees;
            commandObtieneAmazon.Parameters.Add("@othertransactionfees", SqlDbType.VarChar).Value = clsPedido.othertransactionfees;
            commandObtieneAmazon.Parameters.Add("@other", SqlDbType.VarChar).Value = clsPedido.other;
            commandObtieneAmazon.Parameters.Add("@total", SqlDbType.VarChar).Value = clsPedido.total;

            Conexion.Open();
            commandObtieneAmazon.ExecuteNonQuery();
            Conexion.Close();
        }

        public void InsertaBOX(PesosDimensiones.Box clsPedido, SqlConnection Conexion)
        {
            string sqlAInsertaRegistroDW = "InsertaRegistroBOX";
            SqlCommand commandObtieneBOX = new SqlCommand(sqlAInsertaRegistroDW, Conexion);

            commandObtieneBOX.CommandType = CommandType.StoredProcedure;
            commandObtieneBOX.Parameters.Add("@STATE", SqlDbType.VarChar).Value = clsPedido.STATE;
            commandObtieneBOX.Parameters.Add("@POSTALCODE", SqlDbType.VarChar).Value = clsPedido.POSTALCODE;
            commandObtieneBOX.Parameters.Add("@SHIPPER", SqlDbType.VarChar).Value = clsPedido.SHIPPER;
            commandObtieneBOX.Parameters.Add("@PROSHIP_SHIPDATE", SqlDbType.VarChar).Value = clsPedido.PROSHIP_SHIPDATE;
            commandObtieneBOX.Parameters.Add("@PACKAGING_PLAINTEXT", SqlDbType.VarChar).Value = clsPedido.PACKAGING_PLAINTEXT;
            commandObtieneBOX.Parameters.Add("@WEIGHT", SqlDbType.VarChar).Value = clsPedido.WEIGHT;
            commandObtieneBOX.Parameters.Add("@DIMENSIONS", SqlDbType.VarChar).Value = clsPedido.DIMENSIONS;
            commandObtieneBOX.Parameters.Add("@TRACKING_NUMBER", SqlDbType.VarChar).Value = clsPedido.TRACKING_NUMBER;
            commandObtieneBOX.Parameters.Add("@CCN_SAP_ORDER_NUMBER", SqlDbType.VarChar).Value = clsPedido.CCN_SAP_ORDER_NUMBER;
            commandObtieneBOX.Parameters.Add("@CCN_ORDER_NUMBER", SqlDbType.VarChar).Value = clsPedido.CCN_ORDER_NUMBER;
            commandObtieneBOX.Parameters.Add("@CCN_COMPANY_CODE", SqlDbType.VarChar).Value = clsPedido.CCN_COMPANY_CODE;
            commandObtieneBOX.Parameters.Add("@CCN_STR_NUM", SqlDbType.VarChar).Value = clsPedido.CCN_STR_NUM;
            commandObtieneBOX.Parameters.Add("@CCN_DELIVERY_NUMBER", SqlDbType.VarChar).Value = clsPedido.CCN_DELIVERY_NUMBER;
            commandObtieneBOX.Parameters.Add("@SHIPPER_SYMBOL", SqlDbType.VarChar).Value = clsPedido.SHIPPER_SYMBOL;
            commandObtieneBOX.Parameters.Add("@OrderDate", SqlDbType.VarChar).Value = clsPedido.OrderDate;
            commandObtieneBOX.Parameters.Add("@PROSHIP_SERVICE_PLAINTEXT", SqlDbType.VarChar).Value = clsPedido.PROSHIP_SERVICE_PLAINTEXT;
            commandObtieneBOX.Parameters.Add("@CCN_SHIP_TEXT", SqlDbType.VarChar).Value = clsPedido.CCN_SHIP_TEXT;

            Conexion.Open();
            commandObtieneBOX.ExecuteNonQuery();
            Conexion.Close();
        }

        public void InsertaEJDDimensions(PesosDimensiones.EJDDimensions clsPedido, SqlConnection Conexion)
        {
            string sqlAInsertaRegistroEJDDimensions = "InsertaRegistroEJDDimensions";
            SqlCommand commandObtieneEJDDimensions = new SqlCommand(sqlAInsertaRegistroEJDDimensions, Conexion);

            commandObtieneEJDDimensions.CommandType = CommandType.StoredProcedure;
            commandObtieneEJDDimensions.Parameters.Add("@EvpSku", SqlDbType.VarChar).Value = clsPedido.EvpSku;
            commandObtieneEJDDimensions.Parameters.Add("@Title", SqlDbType.VarChar).Value = clsPedido.Title;
            commandObtieneEJDDimensions.Parameters.Add("@EJDSku", SqlDbType.VarChar).Value = clsPedido.EJDSku;
            commandObtieneEJDDimensions.Parameters.Add("@EJDUomCode", SqlDbType.VarChar).Value = clsPedido.EJDUomCode;
            commandObtieneEJDDimensions.Parameters.Add("@EJDUomQuantity", SqlDbType.VarChar).Value = clsPedido.EJDUomQuantity;
            commandObtieneEJDDimensions.Parameters.Add("@Length", SqlDbType.VarChar).Value = clsPedido.Length;
            commandObtieneEJDDimensions.Parameters.Add("@Height", SqlDbType.VarChar).Value = clsPedido.Height;
            commandObtieneEJDDimensions.Parameters.Add("@Width", SqlDbType.VarChar).Value = clsPedido.Width;
            commandObtieneEJDDimensions.Parameters.Add("@Weight", SqlDbType.VarChar).Value = clsPedido.Weight;
           

            Conexion.Open();
            commandObtieneEJDDimensions.ExecuteNonQuery();
            Conexion.Close();
        }

        public void InsertaJensenDimensions(PesosDimensiones.JensenDimensions clsPedido, SqlConnection Conexion)
        {
            string sqlAInsertaRegistroJensenDimensions = "InsertaRegistroJensenDimensions";
            SqlCommand commandObtieneJensenDimensions = new SqlCommand(sqlAInsertaRegistroJensenDimensions, Conexion);

            commandObtieneJensenDimensions.CommandType = CommandType.StoredProcedure;
            commandObtieneJensenDimensions.Parameters.Add("@EvpSku", SqlDbType.VarChar).Value = clsPedido.EvpSku;
            commandObtieneJensenDimensions.Parameters.Add("@Title", SqlDbType.VarChar).Value = clsPedido.Title;
            commandObtieneJensenDimensions.Parameters.Add("@JensenSku", SqlDbType.VarChar).Value = clsPedido.JensenSku;
            commandObtieneJensenDimensions.Parameters.Add("@UomCode", SqlDbType.VarChar).Value = clsPedido.UomCode;
            commandObtieneJensenDimensions.Parameters.Add("@UomQuantity", SqlDbType.VarChar).Value = clsPedido.UomQuantity;
            commandObtieneJensenDimensions.Parameters.Add("@Length", SqlDbType.VarChar).Value = clsPedido.Length;
            commandObtieneJensenDimensions.Parameters.Add("@Height", SqlDbType.VarChar).Value = clsPedido.Height;
            commandObtieneJensenDimensions.Parameters.Add("@Width", SqlDbType.VarChar).Value = clsPedido.Width;
            commandObtieneJensenDimensions.Parameters.Add("@Weight", SqlDbType.VarChar).Value = clsPedido.Weight;


            Conexion.Open();
            commandObtieneJensenDimensions.ExecuteNonQuery();
            Conexion.Close();
        }

        public void InsertaMI15(PesosDimensiones.MI15 clsPedido, SqlConnection Conexion)
        {
            string sqlAInsertaRegistroJensenDimensions = "InsertaRegistroMI15";
            SqlCommand commandObtieneJensenDimensions = new SqlCommand(sqlAInsertaRegistroJensenDimensions, Conexion);

            commandObtieneJensenDimensions.CommandType = CommandType.StoredProcedure;
            commandObtieneJensenDimensions.Parameters.Add("@SHIPPINGDATE", SqlDbType.VarChar).Value = clsPedido.SHIPPINGDATE;
            commandObtieneJensenDimensions.Parameters.Add("@MANIFESTDATE", SqlDbType.VarChar).Value = clsPedido.MANIFESTDATE;
            commandObtieneJensenDimensions.Parameters.Add("@PACKAGEID", SqlDbType.VarChar).Value = clsPedido.PACKAGEID;
            commandObtieneJensenDimensions.Parameters.Add("@USPSTRACKINGNUMBER", SqlDbType.VarChar).Value = clsPedido.USPSTRACKINGNUMBER;
            commandObtieneJensenDimensions.Parameters.Add("@SEQUENCE", SqlDbType.VarChar).Value = clsPedido.SEQUENCE;
            commandObtieneJensenDimensions.Parameters.Add("@COSTCENTER1", SqlDbType.VarChar).Value = clsPedido.COSTCENTER1;
            commandObtieneJensenDimensions.Parameters.Add("@COSTCENTER2", SqlDbType.VarChar).Value = clsPedido.COSTCENTER2;
            commandObtieneJensenDimensions.Parameters.Add("@COSTCENTER3", SqlDbType.VarChar).Value = clsPedido.COSTCENTER3;
            commandObtieneJensenDimensions.Parameters.Add("@BILLEDWEIGHT", SqlDbType.VarChar).Value = clsPedido.BILLEDWEIGHT;

            commandObtieneJensenDimensions.Parameters.Add("@WEIGHTTYPE", SqlDbType.VarChar).Value = clsPedido.WEIGHTTYPE;
            commandObtieneJensenDimensions.Parameters.Add("@ZIP", SqlDbType.VarChar).Value = clsPedido.ZIP;
            commandObtieneJensenDimensions.Parameters.Add("@ZONE", SqlDbType.VarChar).Value = clsPedido.ZONE;
            commandObtieneJensenDimensions.Parameters.Add("@SERVICE", SqlDbType.VarChar).Value = clsPedido.SERVICE;
            commandObtieneJensenDimensions.Parameters.Add("@UPSMI", SqlDbType.VarChar).Value = clsPedido.UPSMI;
            commandObtieneJensenDimensions.Parameters.Add("@SAVINGS", SqlDbType.VarChar).Value = clsPedido.SAVINGS;
            commandObtieneJensenDimensions.Parameters.Add("@OVERLABELEDUSPSTRACKING", SqlDbType.VarChar).Value = clsPedido.OVERLABELEDUSPSTRACKING;
            commandObtieneJensenDimensions.Parameters.Add("@ERRORREASON", SqlDbType.VarChar).Value = clsPedido.ERRORREASON;

            Conexion.Open();
            commandObtieneJensenDimensions.ExecuteNonQuery();
            Conexion.Close();
        }
    }
}
