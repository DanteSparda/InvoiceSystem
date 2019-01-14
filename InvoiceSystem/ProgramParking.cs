using ExcelDataReader;
using InvoiceSystem.Helpers;
using InvoiceSystem.Models;
using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
namespace InvoiceSystem
{
  class ProgramParking
  {
    private const string DATEFORMAT = "dd.MM.yyyy";
    private const string INVOICE_LOCATION = "София";
    private static List<string> Months = new List<string> { "Януари", "Февруари", "Март", "Април", "Май", "Юни", "Юли", "Август", "Септември", "Октомври", "Ноември", "Декември" };
    public static void MainParking()
    {
      Warning[] warnings;
      string[] streamids;
      string mimeType;
      string encoding;
      string filenameExtension;

      var path = $@"{Environment.CurrentDirectory}\For Software Metro City 2018.xlsx";
      var invoiceNumber = int.Parse(ConfigurationManager.AppSettings["InvoiceNumber"]);
      var invoiceLength = ConfigurationManager.AppSettings["InvoiceLength"];

      using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
      {
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
          var result = reader.AsDataSet();
          var invoiceData = result.Tables[0];
          var headingRows = invoiceData.Rows[0];
          var propertyNameColIndex = GetIndexOfHeading(headingRows, "Имот");
          var quantityColIndex = GetIndexOfHeading(headingRows, "Брой паркоместа");
          var shipToNameColIndex = GetIndexOfHeading(headingRows, "Име (фирма или физическо)");
          var carrierReferenceColIndex = GetIndexOfHeading(headingRows, "ЕИК/ ЕГН");
          var shipToAddressColIndex = GetIndexOfHeading(headingRows, "Адрес");
          var tavColIndex = GetIndexOfHeading(headingRows, "МОЛ");
          var taxNumber = GetIndexOfHeading(headingRows, "ДДС");
          var emails = GetIndexOfHeading(headingRows, "Електронна поща");

          var expencesHeadingRowForInvoice = GetDataRowFromTableBasedOnRowValue(result.Tables[1].Rows, "Имот");
          var totalPriceIndex = GetIndexOfHeading(expencesHeadingRowForInvoice, "Сума с ДДС");
          //var constantCostsFirstIndex = GetIndexOfHeading(expencesHeadingRowForInvoice, "Постоянни разходи");
          //var electricCostsFirstIndex = GetIndexOfHeading(expencesHeadingRowForInvoice, "Електрическа енергия (средна стойност)");
          //var accidentalCostsFirstIndex = GetIndexOfHeading(expencesHeadingRowForInvoice, "Инцидентни разходи");
          //var addinitionalReserveFirstIndex = GetIndexOfHeading(expencesHeadingRowForInvoice, "Фонд допълнителни резервни средства");
          //var bankTaxesFirstIndex = GetIndexOfHeading(expencesHeadingRowForInvoice, "Банкови такси");

          invoiceData.Rows.RemoveAt(0);
          var invoices = new List<Invoice>();
          foreach (DataRow row in invoiceData.Rows)
          {
            var propertyName = GetContent(row, propertyNameColIndex);
            if (string.IsNullOrEmpty(propertyName))
            {
              continue;
            }


            var invoice = new Invoice()
            {
              Header = new InvoiceHeaderModel()
              {
                ShipToName = GetContent(row, shipToNameColIndex),
                CarrierReference = GetContent(row, carrierReferenceColIndex),
                TAV = GetContent(row, tavColIndex),
                ShipToAddress = GetContent(row, shipToAddressColIndex),
                Receiver = GetContent(row, tavColIndex),
                Date = DateTime.Now.ToString(DATEFORMAT),
                PaymentDate = new DateTime(2019, 01, 07).ToString(DATEFORMAT),
                DateOfTaxEvent = DateTime.Now.ToString(DATEFORMAT),
                DealLocation = INVOICE_LOCATION,
                InvoiceNumber = invoiceNumber.ToString($"D{invoiceLength}"),
                ShipToCity = INVOICE_LOCATION,
                TaxNumber = $"BG{GetContent(row, carrierReferenceColIndex)}",
                Emails = GetContent(row, emails).Trim().Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList()
              },
              Name = $"{propertyName}-{GetContent(row, shipToNameColIndex)}"
            };

            if (GetContent(row, taxNumber) != "ДА")
            {
              invoice.Header.TaxNumber = string.Empty;
            }
            //var totalAmountWithoutTaxColIndex = GetIndexOfHeading()

            var expencesValuesRowForInvoice = GetDataRowFromTableBasedOnRowValue(result.Tables[1].Rows, propertyName, invoice.Header.ShipToName);
            var totalPrice = decimal.Parse(GetContent(expencesValuesRowForInvoice, totalPriceIndex));
            var quantity = decimal.Parse(GetContent(row, quantityColIndex));

            var totalPriceAsDecimal = decimal.Round(totalPrice, 2);
            var constantCosts = string.Empty;
            var electricBill = string.Empty;
            var accidentalCosts = string.Empty;
            var addinitionalReserve = string.Empty;
            var bankTaxes = string.Empty;

            invoice.Lines.Add(new InvoiceLineModel
            {
              Description = GetLineDescription(propertyName, Months[DateTime.Now.Month - 1], DateTime.Now.Year.ToString(), constantCosts.ToString(), electricBill, accidentalCosts, addinitionalReserve, bankTaxes),
              Quantity = quantity.ToString(),
              Code = string.Empty,
              TotalPrice = totalPriceAsDecimal,
              UoM = "брой.",
              UnitPrice = decimal.Round(totalPrice / quantity, 3)
            });

            invoices.Add(invoice);

            invoiceNumber++;
          }

          CalculateTotalForInvoices(invoices);

          var report = new LocalReport();
          report.ReportPath = @"D:\Projects\RDLCCreator\RDLCCreator\Reports\Invoice-bg - Copy.rdlc";
          report.EnableExternalImages = true;
          foreach (var invoice in invoices)
          {
            report.DataSources.Add(new ReportDataSource("Header", new List<InvoiceHeaderModel> { invoice.Header }));
            report.DataSources.Add(new ReportDataSource("Lines", invoice.Lines));
            byte[] bytes = report.Render("PDF", null, out mimeType, out encoding, out filenameExtension, out streamids, out warnings);

            using (FileStream fs = new FileStream($"Invoices/{invoice.Name}.pdf", FileMode.Create))
            {
              fs.Write(bytes, 0, bytes.Length);
            }
            if (invoice.Header.Emails.Any())
            {
              SmtpClient client = new SmtpClient("smtp.office365.com", 587);
              client.EnableSsl = true;
              client.Credentials = new System.Net.NetworkCredential("bc@metrocity.center", "Metr0citY2018_dr");
              MailAddress from = new MailAddress("bc@metrocity.center", String.Empty, System.Text.Encoding.UTF8);
              MailAddress to = new MailAddress(invoice.Header.Emails.First().Trim());
              MailMessage message = new MailMessage(from, to);
              foreach(var recipient in invoice.Header.Emails.Skip(1))
              {
                message.To.Add(recipient.Trim());
              }

              message.Body = @"Здравейте,

Приложено изпращаме фактура за управление и поддръжка на общи части за паркоместа, с които членувате в ЕС Складове и Паркоместа Метро Сити, гр. София 1712, бул. Александър Малинов № 51, Булстат 177119281.

Настоящето електронно писмо и приложените документи са генерирани по електронен път, чрез софтуер „Метро Сити 2018“.

УС на ЕС Складове и Паркоместа";

              message.BodyEncoding = System.Text.Encoding.UTF8;
              message.Attachments.Add(new Attachment(File.OpenRead($"Invoices/{invoice.Name}.pdf"), $"{invoice.Name}.pdf"));
              message.Subject = "Фактура за управление и поддръжка на общи части за паркоместа";
              message.SubjectEncoding = System.Text.Encoding.UTF8;

              client.Send(message);
            }


            report.DataSources.Clear();
          }


          Console.WriteLine("done");
        }
      }
    }


    private static void CalculateTotalForInvoices(List<Invoice> invoices)
    {
      foreach (var invoice in invoices)
      {
        var totalPriceWithoutTax = invoice.Lines.Sum(x => x.TotalPrice);
        var totalPriceWithTax = (totalPriceWithoutTax).ToString("0.00");
        invoice.Header.TotalAmountIncludingTaxes = totalPriceWithTax;
        invoice.Header.TotalAmountIncludingTaxesWithWords = PriceConverter.ConvertToPriceString(totalPriceWithTax);
      }
    }

    public static DataRow GetDataRowFromTableBasedOnRowValue(DataRowCollection rowCollection, string rowValue, string identifier = "")
    {
      var headingRowIndex = -1;
      foreach (var row in rowCollection)
      {
        headingRowIndex = GetIndexOfHeading((DataRow)row, rowValue);
        if (headingRowIndex > -1)
        {
          break;
        }
      }

      if (headingRowIndex < 0)
      {
        return null;
      }

      foreach (DataRow row in rowCollection)
      {
        if (row.ItemArray[headingRowIndex].ToString() == rowValue && (string.IsNullOrEmpty(identifier) || (!string.IsNullOrEmpty(identifier) && row.ItemArray[3].ToString() == identifier)))
        {
          return row;
        }
      }

      return null;
    }

    public static int GetIndexOfHeading(DataRow dataRow, string headingContent)
    {
      return dataRow.ItemArray.ToList().IndexOf(headingContent);
    }

    public static string GetContent(DataRow datarow, int index)
    {
      var data = datarow.ItemArray[index];
      return data.ToString();
    }

    public static decimal ParseAndSumDecimals(string firstNumber, string secondNumber)
    {
      var firstNumberAsDecimal = decimal.Parse(firstNumber);
      var secondNumberAsDecimal = decimal.Parse(secondNumber);
      var result = decimal.Round(firstNumberAsDecimal + secondNumberAsDecimal, 2);
      return result;
    }

    /// <summary>
    /// Gets the index cell and the next cell and sums them up
    /// </summary>
    /// <param name="dataRow">Row to work on</param>
    /// <param name="index">Index of first cell</param>
    /// <returns>Decimal sum</returns>
    public static decimal GetSplitColSum(DataRow dataRow, int index)
    {
      return ParseAndSumDecimals(GetContent(dataRow, index), GetContent(dataRow, index + 1));
    }

    public static string GetLineDescription(string office, string month, string year, string constantCosts, string electricBill, string accidentalCosts, string addinitionalReserve, string bankTaxes)
    {
      return $@"Разходи за управление и поддръжка на общите части в Складове и паркоместа находящ се на адрес гр. София 1712, бул. Александър Малинов No 51, {office}, за Ноември и Декември 2018 година";
    }
  }
}
