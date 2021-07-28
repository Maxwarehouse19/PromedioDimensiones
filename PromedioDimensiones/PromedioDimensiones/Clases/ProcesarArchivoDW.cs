using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PromedioDimensionesMaxwarehouse.Clases
{
    class ProcesarArchivoDW
    {
        public string SalesOrderNumber              ;
        public string HoldCode                      ;
        public string TotalSales                    ;
        public string SalesSku                      ;
        public string SalesCategoryAtTimeOfSale     ;
        public string UomCode                       ;
        public string UomQuantity                   ;
        public string SalesStatus                   ;
        public string SalesOrderDate                ;
        public string SalesChannelName              ;
        public string CustomerName                  ;
        public string FulfillmentSku                ;
        public string FulfillmentChannelName        ;
        public string FulfillmentChannelType        ;
        public string LinkedFulfillmentChannelName  ;
        public string FulfillmentLocationName       ;
        public string FulfillmentOrderNumber        ;
        public string Quantity                      ;
        public string Sku                           ;
        public string Title                         ;
        public string TotalCost                     ;
        public string Commission                    ;
        public string InventoryCost                 ;
        public string UnitCost                      ;
        public string ServiceCost                   ;
        public string EstimatedShippingCost         ;
        public string ShippingCost                  ;
        public string ShippingPrice                 ;
        public string OverheadCost                  ;
        public string PackageCost                   ;
        public string ProfitLoss                    ;
        public string Carrier                       ;
        public string ShippingServiceLevel          ;
        public string ShippedByUser                 ;
        public string ShippingWeight                ;
        public string Length                        ;
        public string varWidth                      ;
        public string varHeight                     ;
        public string Weight                        ;
        public string StateRegion                   ;
        public string TrackingNum                   ;
        public string MfrName                       ;
        public string PricingRule                   ;
        public string ActualShippingCost            ;
        public string ActualShipping                ;
        public string ShippingCostDifference        ;

        // obtiene el valor del registro actual
        // ------------------------------------
        public void ObtieneValorRegistro(string[] valor)
        {
            SalesOrderNumber             = valor[0];;
            HoldCode                     = valor[1];;
            TotalSales                   = valor[2];;
            SalesSku                     = valor[3];;
            SalesCategoryAtTimeOfSale    = valor[4];;
            UomCode                      = valor[5];;
            UomQuantity                  = valor[6];;
            SalesStatus                  = valor[7];;
            SalesOrderDate               = valor[8];;
            SalesChannelName             = valor[9];;
            CustomerName                 = valor[10];
            FulfillmentSku               = valor[11];
            FulfillmentChannelName       = valor[12];
            FulfillmentChannelType       = valor[13];
            LinkedFulfillmentChannelName = valor[14];
            FulfillmentLocationName      = valor[15];
            FulfillmentOrderNumber       = valor[16];
            Quantity                     = valor[17];
            Sku                          = valor[18];
            Title                        = valor[19];
            TotalCost                    = valor[20];
            Commission                   = valor[21];
            InventoryCost                = valor[22];
            UnitCost                     = valor[23];
            ServiceCost                  = valor[24];
            EstimatedShippingCost        = valor[25];
            ShippingCost                 = valor[26];
            ShippingPrice                = valor[27];
            OverheadCost                 = valor[28];
            PackageCost                  = valor[29];
            ProfitLoss                   = valor[30];
            Carrier                      = valor[31];
            ShippingServiceLevel         = valor[32];
            ShippedByUser                = valor[33];
            ShippingWeight               = valor[34];
            Length                       = valor[35];
            varWidth                     = valor[36];
            varHeight                    = valor[37];
            Weight                       = valor[38];
            StateRegion                  = valor[39];
            TrackingNum                  = valor[40];
            MfrName                      = valor[41];
            PricingRule                  = valor[42];
            //ActualShippingCost           = valor[43];
            //ActualShipping               = valor[44];
            //ShippingCostDifference       = valor[45];
        }

    }
}
