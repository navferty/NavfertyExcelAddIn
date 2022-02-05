using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml;

using Navferty.ExcelAddIn.Web.Localization;

using Newtonsoft.Json;

using NLog;

namespace Navferty.ExcelAddIn.Web.CurrencyExchangeRates.Providers
{
    internal class ECBProvider : ExchangeRatesDataProviderBaase
    {
        private const string C_EURO_ISO = "EUR";
        private const char C_EURO = '€';



        private static readonly Lazy<CultureInfo> ci = new Lazy<CultureInfo>(() =>
        {
            CultureInfo ciNew = (CultureInfo)CultureInfo.GetCultureInfo("en-GB").Clone();
            ciNew.NumberFormat.CurrencySymbol = C_EURO.ToString();
            return ciNew;
        });

        public override string Title => UIStrings.CurrencyExchangeRates_Sources_ECB;

        private readonly ILogger logger = LogManager.GetCurrentClassLogger();
        public override ILogger Logger => logger;

        private string rawXML = String.Empty;
        private ECB.ECBExchangeRatesRecord[] rawRows = Array.Empty<ECB.ECBExchangeRatesRecord>();
        private HttpClient web = new();

        protected override async Task<WebResultRow[]> DownloadWebResultRowsForDate(DateTime dt)
        {
            rawXML = String.Empty;
            rawRows = Array.Empty<ECB.ECBExchangeRatesRecord>();

            //https://sdw-wsrest.ecb.europa.eu/service/data/EXR?startPeriod=2022-02-01&endPeriod=2022-02-01
            string sDate = dt.ToString("yyyy-MM-dd");
            var urlECBExchangeForDate = @$"https://sdw-wsrest.ecb.europa.eu/service/data/EXR?startPeriod={sDate}&endPeriod={sDate}";
            logger.Debug($"ECB: ExchangeRates query url: {urlECBExchangeForDate}");

            rawXML = await (await web.GetAsync(urlECBExchangeForDate)).
                EnsureSuccessStatusCode().
                Content.ReadAsStringAsync();

            if (string.IsNullOrWhiteSpace(rawXML))
            {
                logger.Error("ECB: web service answer xml = null!");
                throw new Exception(UIStrings.CurrencyExchangeRates_Error_Network);
            }

            try
            {
                rawRows = ParseECBXml(rawXML);
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"ECB: Failed to parse xml:\n{rawXML}");
                throw new Exception(UIStrings.CurrencyExchangeRates_Error_ParseError);
            }

            try
            {
                return rawRows.Select(row => new WebResultRow(row)).ToArray();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "ECB: Failed to convert 'ECBExchangeRatesRecord' to 'WebResultRow'!");

                throw new Exception(UIStrings.CurrencyExchangeRates_Error_ParseError);

            }
        }

        private ECB.ECBExchangeRatesRecord[] ParseECBXml(string xmlText)
        {
            var doc = new XmlDocument();
            doc.LoadXml(xmlText);

            //Get all xml nodes with <generic:Series> tag.
            //This tag is about some asset exchange rate. Not only money!
            var seriesTags = doc.GetElementsByTagName("generic:Series").Cast<XmlElement>().ToArray();
            var rawRows = seriesTags
                .Select(tagSeries => new ECB.ECBExchangeRatesRecord(tagSeries))
                .Where(x => x.CurrencyDenom == C_EURO_ISO)//Select only rates related to Euro ('EUR'). Full list contains many other asset exchange rates.
                .ToArray();

            return rawRows;
        }

        public override CultureInfo Culture => ci.Value;
    }
}
