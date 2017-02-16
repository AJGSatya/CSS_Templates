using System.Text;

namespace OAMPS.Office.Word.Helpers
{
    public static class FrameworkExtensions
    {
        public static string ToText(this int num)
        {
            StringBuilder result;

            if (num < 0)
            {
                return string.Format("Minus {0}", ToText(-num));
            }

            if (num == 0)
            {
                return "Zero";
            }

            if (num <= 19)
            {
                var oneToNineteen = new[]
                    {
                        "One",
                        "Two",
                        "Three",
                        "Four",
                        "Five",
                        "Six",
                        "Seven",
                        "Eight",
                        "Nine",
                        "Ten",
                        "Eleven",
                        "Twelve",
                        "Thirteen",
                        "Fourteen",
                        "Fifteen",
                        "Sixteen",
                        "Seventeen",
                        "Eighteen",
                        "Nineteen"
                    };

                return oneToNineteen[num - 1];
            }

            if (num <= 99)
            {
                result = new StringBuilder();

                var multiplesOfTen = new[]
                    {
                        "Twenty",
                        "Thirty",
                        "Forty",
                        "Fifty",
                        "Sixty",
                        "Seventy",
                        "Eighty",
                        "Ninety"
                    };

                result.Append(multiplesOfTen[(num/10) - 2]);

                if (num%10 != 0)
                {
                    result.Append(" ");
                    result.Append(ToText(num%10));
                }

                return result.ToString();
            }

            if (num == 100)
            {
                return "One Hundred";
            }

            if (num <= 199)
            {
                return string.Format("One Hundred and {0}", ToText(num%100));
            }

            if (num <= 999)
            {
                result = new StringBuilder((num/100).ToText());
                result.Append(" Hundred");
                if (num%100 != 0)
                {
                    result.Append(" and ");
                    result.Append((num%100).ToText());
                }

                return result.ToString();
            }

            if (num <= 999999)
            {
                result = new StringBuilder((num/1000).ToText());
                result.Append(" Thousand");
                if (num%1000 != 0)
                {
                    switch ((num%1000) < 100)
                    {
                        case true:
                            result.Append(" and ");
                            break;
                        case false:
                            result.Append(", ");
                            break;
                    }

                    result.Append((num%1000).ToText());
                }

                return result.ToString();
            }

            if (num <= 999999999)
            {
                result = new StringBuilder((num/1000000).ToText());
                result.Append(" Million");
                if (num%1000000 != 0)
                {
                    switch ((num%1000000) < 100)
                    {
                        case true:
                            result.Append(" and ");
                            break;
                        case false:
                            result.Append(", ");
                            break;
                    }

                    result.Append((num%1000000).ToText());
                }

                return result.ToString();
            }

            result = new StringBuilder((num/1000000000).ToText());
            result.Append(" Billion");
            if (num%1000000000 != 0)
            {
                switch ((num%1000000000) < 100)
                {
                    case true:
                        result.Append(" and ");
                        break;
                    case false:
                        result.Append(", ");
                        break;
                }

                result.Append((num%1000000000).ToText());
            }

            return result.ToString();
        }
    }
}