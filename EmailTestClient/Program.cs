namespace EmailTestClient
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    using System.Timers;
    using Microsoft.Exchange.WebServices.Data;

    internal class Program
    {
        private static string ServerFqdn = "https://mail.haywardandgreen.com/EWS/Exchange.asmx";
        private static ExchangeService service;
        private static string Username = "tomasgilmartin";
        private static string Password = "Airwolf96";
        internal static List<PullSubscription> Subscriptions = new List<PullSubscription>();

        private static void Main(string[] args)
        {
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            Console.WriteLine("Creating service at {0}", ServerFqdn);
            try
            {
                service = new ExchangeService(ExchangeVersion.Exchange2007_SP1)
                {
                    Credentials = new WebCredentials(Username, Password, "haywardandgreen.com"),
                    TraceListener = new TextTraceListener(),
                    TraceFlags = TraceFlags.All,
                    TraceEnabled = true,
                    Url = new Uri(ServerFqdn)
                };
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exiting...Error creating service: {0}", exception.Message);
                Console.ReadLine();
                return;
            }
            Console.WriteLine("Service created!");

            try
            {
                PullSubscription subscription = service.SubscribeToPullNotifications(new FolderId[] { WellKnownFolderName.Inbox }, 1440, "", EventType.Created);
                Subscriptions.Add(subscription);
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exiting...Error creating subscription: {0}", exception.Message);
                Console.ReadLine();
                return;
            }

            Console.WriteLine("Beginning polling at 30 sec interval");

            var timer = new Timer(30000);
            timer.AutoReset = true;
            timer.Elapsed += (s, e) =>
                {
                    Console.WriteLine("Begin Polling for 'Quote Request' in the subject line");
                    foreach (Item item in Subscriptions.SelectMany(pullSubscription => pullSubscription.GetEvents().ItemEvents.Select(itemEvent => Item.Bind(service, itemEvent.ItemId)).Where(item => item.Subject.Contains("Quote Request"))))
                    {
                        Console.WriteLine(item.Body.Text);
                    }
                    Console.WriteLine("End Polling");
                };
            timer.Enabled = true;
            Console.WriteLine("\r\n");
            Console.WriteLine("Press or select Enter to quit polling...");
            Console.Read();
        }
        
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        private static bool CertificateValidationCallBack(
            object sender,
            X509Certificate certificate,
            X509Chain chain,
            SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                            (status.Status == X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }
    }
}