using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PromedioDimensiones
{
    class PedidoUSPS
    {
        public string AccountNumber { get; set; }
        public string ID { get; set; }
        public string DateTime { get; set; }
        public string Postmark { get; set; }
        public string Origin { get; set; }
        public string Destination { get; set; }
        public string Type { get; set; }
        public string MailClass { get; set; }
        public string TrackingNumber { get; set; }
        public string DeclaredValue { get; set; }
        public string TotalPostageAmt { get; set; }
        public string Balance { get; set; }
        public string RefundStatus { get; set; }
        public string GroupCode { get; set; }
        public string ReferenceID { get; set; }
        public string DeliveryDate { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
        public string Weight { get; set; }
        public string OptionalServices { get; set; }
        public string DestinationName { get; set; }
        public string DestinationCompanyName { get; set; }
        public string DestinationAddress { get; set; }
        public string DestinationCity { get; set; }
        public string DestinationState { get; set; }
        public string DestinationZip { get; set; }
        public string DestinationCountry { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public string Reference2 { get; set; }
        public string Reference3 { get; set; }
        public string Reference4 { get; set; }
        public string PackageDescription { get; set; }
        public string Zone { get; set; }
        public string IsCubic { get; set; }
        public string CubicValue { get; set; }
        public string AdjWeight { get; set; }
        public string AdjDimensions { get; set; }
        public string AdjFromZIP { get; set; }
        public string AdjToZIP { get; set; }
        public string AdjMailClass { get; set; }


    }

}
