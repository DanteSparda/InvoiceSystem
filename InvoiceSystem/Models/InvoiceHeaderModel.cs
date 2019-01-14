using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceSystem.Models
{
  public class InvoiceHeaderModel
  {
    #region HeadingData
    public string InvoiceNumber { get; set; }

    public string Date { get; set; }

    public string PaymentDate { get; set; }

    public string ShipToName { get; set; }

    public string CarrierReference { get; set; }

    public string TaxNumber { get; set; }

    public string ShipToCity { get; set; }

    public string ShipToAddress { get; set; }

    public string TAV { get; set; }

    public string ShipToPhone { get; set; }

    public string TotalAmountIncludingTaxes { get; set; }

    public string TotalAmountIncludingTaxesWithWords { get; set; }
    public List<string> Emails { get; set; }

    #endregion

    #region FooterData
    public string DateOfTaxEvent { get; set; }

    public string DealReason { get; set; }

    public string DealDescription { get; set; }

    public string DealLocation { get; set; }

    public string Receiver { get; set; }
    #endregion

  }
}
