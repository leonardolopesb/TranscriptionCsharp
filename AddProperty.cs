using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Net.Http;
using System.Net.Mail;
using System.Net.Sockets;
using System.Runtime.Remoting;
using System.Security.Policy;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using Newtonsoft.Json;
using OtpNet;

namespace TesteConsole
{
    public class AddProperty
    {
        public static readonly HttpClient client = new HttpClient();

        public static async Task Main(string[] args)
        {
            string url = "https://api_request_from_tjpe.com.br/...";

            string login = "my-login";
            string secretKey = "secret-key";

            short day = x;
            short month = y;
            short year = z;

            string folder = $"C:\\Users\\file_path";

            string[] fullPath = Directory.GetFiles(folder);

            string[] custom2FA = new string[fullPath.Length + 1];

            short n = 1;

            for (short i = n; i <= fullPath.Length; i++)
            {
                string file = fullPath[i - 1];
                string id = Path.GetFileNameWithoutExtension(file);

                custom2FA[i] = GenerateOTP(secretKey);

                Console.WriteLine("OTP: " + custom2FA[i] + "\n");

                if (i == n)
                {
                    client.DefaultRequestHeaders.Add("custom-2FA", custom2FA[i]);
                    client.DefaultRequestHeaders.Add("login", login);
                }

                if (i > n)
                {
                    client.DefaultRequestHeaders.Clear();

                    if (custom2FA[i] != custom2FA[i - 1])
                    {
                        custom2FA[i] = GenerateOTP(secretKey);
                    }
                    client.DefaultRequestHeaders.Add("custom-2FA", custom2FA[i]);
                    client.DefaultRequestHeaders.Add("login", login);
                }

                var link = url + id;

                HttpResponseMessage response = await client.GetAsync(link);

                Console.WriteLine(response.StatusCode);
                Console.WriteLine(response.RequestMessage);

                response.EnsureSuccessStatusCode();

                var jsonResponse = await response.Content.ReadAsStringAsync();

                var json = JsonConvert.DeserializeObject<List<Chamado>>(jsonResponse);

                var oferta = json[0].Offer;

                DateTime data = new DateTime(year, month, (day + 1));

                string propertyName_offer = "Oferta";
                object propertyValue_offer = oferta;
                PropertyTypes propertyType_offer = PropertyTypes.Text;

                string propertyName_date = "DataLigacao";
                DateTime propertyValue_date = data;
                PropertyTypes propertyType_date = PropertyTypes.DateTime;

                try
                {
                    string returnDate = SetCustomProperty(file, propertyName_date, propertyValue_date, propertyType_date);
                    string returnOffer = SetCustomProperty(file, propertyName_offer, propertyValue_offer, propertyType_offer);

                    Console.WriteLine("\nConcluído com sucesso!");
                }

                catch (Exception ex)
                {
                    Console.WriteLine("\nMensagem de erro: " + ex.Message);
                }

                Console.WriteLine($"\nID do {i}º arquivo: {id}\nOferta: {oferta}\n\n");
            }
        }

        private static string GenerateOTP(string secret)
        {
            var secretBytes = Base32Encoding.ToBytes(secret);
            var totp = new Totp(secretBytes);

            return totp.ComputeTotp();
        }

        private static string SetCustomProperty(string fileName, string propertyName, object propertyValue, PropertyTypes propertyType)
        {
            string returnValue = string.Empty;

            var newProp = new CustomDocumentProperty();
            bool propSet = false;

            string propertyValueString = propertyValue.ToString() ?? throw new ArgumentNullException("null");

            if (propertyValueString == null)
            {
                propertyValueString = string.Empty;
            }

            switch (propertyType)
            {
                case PropertyTypes.DateTime:
                    if (propertyValue is DateTime)
                    {
                        newProp.VTFileTime = new VTFileTime(string.Format("{0:s}Z", Convert.ToDateTime(propertyValue)));
                        propSet = true;
                    }
                    break;

                default:
                    newProp.VTLPWSTR = new VTLPWSTR(propertyValueString);
                    propSet = true;
                    break;
            }

            if (!propSet)
            {
                throw new InvalidOperationException("Invalid property type.");
            }

            newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            newProp.Name = propertyName;

            using (var document = WordprocessingDocument.Open(fileName, true))
            {
                var customProps = document.CustomFilePropertiesPart;
                if (customProps == null)
                {
                    customProps = document.AddCustomFilePropertiesPart();
                    customProps.Properties = new Properties();
                }

                var props = customProps.Properties;
                if (props != null)
                {
                    var prop = props.FirstOrDefault(p => ((CustomDocumentProperty)p).Name.Value == propertyName);

                    if (prop != null)
                    {
                        returnValue = prop.InnerText;
                        prop.Remove();
                    }

                    props.AppendChild(newProp);
                    int pid = 2;
                    foreach (CustomDocumentProperty item in props)
                    {
                        item.PropertyId = pid++;
                    }
                    props.Save();
                }
            }

            return returnValue;
        }
    }

        public class Chamado
        {
            public string ID { get; set; }
            public string Offer { get; set; }
        }

        enum PropertyTypes : int
        {
            DateTime,
            Text
        }
    }
