using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace InvoiceSystem.Helpers
{
  public static class PriceConverter
  {
    public static string[] ones = new string[] { string.Empty, "един", "два", "три", "четири", "пет", "шест", "седем", "осем", "девет" };
    public static string[] thousandsOnes = new string[] { string.Empty, "един", "две", "три", "четири", "пет", "шест", "седем", "осем", "девет" };
    public static string[] tensStartingWithOne = new string[] { "единайсет", "дванайсет", "тринайсет", "четиринайсет", "петнайсет", "шестнайсет", "седемнайсет", "осемнайсет", "деветнайсет" };
    public static string[] hundrets = new string[] { string.Empty, "сто", "двеста", "триста", "четиристотин", "петстотин", "шестстотин", "седемстотин", "осемстотин", "деветстотин" };

    public static string ConvertToPriceString(string price)
    {
      var result = string.Empty;
      var resultPlaceholder = "{0} лв. и {1} ст.";
      var parts = price.ToString().Split('.');
      var amount = parts[0];
      var coins = parts[1];

      if (amount.Length <= 2)
      {
        result = string.Format(resultPlaceholder, ParseTens(amount, false), coins);
      }
      else if (amount.Length == 3)
      {
        result = string.Format(resultPlaceholder, $"{hundrets[int.Parse(amount[0].ToString())]} {ParseTens(amount.Substring(1, 2), true)}", coins);
      }
      else if (amount.Length == 4)
      {
        var thousandsWord = amount[0] == '1' ? "хиляда" : $"{thousandsOnes[int.Parse(amount[0].ToString())]} хиляди";
        var parsedTens = ParseTens(amount.Substring(2, 2), true);
        var thousandsAnd = parsedTens.Length == 0 ? " и " : string.Empty;
        result = string.Format(resultPlaceholder, $"{thousandsWord} {thousandsAnd} {hundrets[int.Parse(amount[1].ToString())]} {parsedTens}", coins);
      }
      result = Regex.Replace(result, @"\s+", " ");
      result = char.ToUpper(result[0]) + result.Substring(1);
      return result;


    }

    private static string ParseTens(string tens, bool hasHundrets)
    {
      var indexes = tens.ToCharArray().Select(x => int.Parse(x.ToString())).ToList();
      var result = hasHundrets ? " и " : string.Empty;
      if (tens.Length == 0 || !indexes.Any(x => x != 0))
      {
        return string.Empty;
      }

      if (tens == "10")
      {
        return result + "десет";
      }

      if (indexes.Count == 1 || indexes[0] == 0)
      {
        return result + ones[indexes.Last()];
      }
      else if (indexes[0] == 1)
      {
        return result + tensStartingWithOne[indexes[1] - 1];
      }
      else
      {
        var onesCalculation = indexes[1] == 0 ? string.Empty : $" и {ones[indexes[1]]}";

        var hundredsPrefix = onesCalculation == string.Empty ? result : string.Empty;

        return hundredsPrefix + ones[indexes[0]] + "десет" + $"{onesCalculation}";
      }
    }
  }
}
