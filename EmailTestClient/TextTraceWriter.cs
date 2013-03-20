namespace EmailTestClient
{
    using System;
    using System.IO;
    using Microsoft.Exchange.WebServices.Data;

    public class TextTraceListener : ITraceListener
    {
        public void Trace(string traceType, string traceMessage)
        {
            CreateXMLTextFile(traceType, traceMessage);
        }

        private void CreateXMLTextFile(string fileName, string traceContent)
        {
            try
            {
                if (!Directory.Exists(@"..\\TraceOutput"))
                {
                    Directory.CreateDirectory(@"..\\TraceOutput");
                }

                File.WriteAllText(@"..\\TraceOutput\\" + fileName + DateTime.Now.Ticks + ".txt", traceContent);
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}