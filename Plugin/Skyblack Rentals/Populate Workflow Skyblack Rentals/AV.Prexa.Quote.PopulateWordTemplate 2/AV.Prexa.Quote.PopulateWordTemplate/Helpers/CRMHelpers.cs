/*using System;

using System.Net;

using Microsoft.Crm.Sdk.Messages;

using Microsoft.PowerPlatform.Dataverse.Client;

namespace AV.Prexa.Quote.PopulateWordTemplate.Helpers
{
    internal class CRMHelper
    {
        public ServiceClient service { get; set; } = null;

        public CRMHelper()

        {

            EstablishCRMConnection();

        }


        /// <summary>
        /// Establishes the connection to D365.         
        /// </summary>       
        private void EstablishCRMConnection()
        {
            string functionName = "EstablishCRMConnection: ";

            //< add key = "clientId" value = "c4c61e7c-f63d-46eb-bed0-da8eb39a292d" />

            // < add key = "clientSecret" value = "lmr8Q~6ybA6NHk2XEnxW9e35mUbix9r8cfoRma2L" />

            //< add key = "crmURL" value = "https://prexa365devnew.crm3.dynamics.com" />


            try
            {
                string crmEnvironmentURL = string.Empty, clientId = string.Empty, clientSecret = string.Empty, callerId = string.Empty;
                crmEnvironmentURL = "prexa365devnew.crm3";
                clientSecret = "lmr8Q~6ybA6NHk2XEnxW9e35mUbix9r8cfoRma2L";
                clientId = "c4c61e7c-f63d-46eb-bed0-da8eb39a292d";
                callerId = "";


                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                string _connectionString = @$"Url=https://{crmEnvironmentURL}.dynamics.com;AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";

                ServiceClient _service = new(_connectionString);

                if (_service != null && _service.IsReady)

                {

                    if (!string.IsNullOrEmpty(callerId))

                    {

                        _service.CallerId = new Guid(callerId);

                    }

                    service = _service;

                    string orgId = WhoAmI();

                }

                else

                {

                    throw new Exception("Connection is not ready. Error:" + _service.LastError);

                }

            }

            catch (Exception ex)

            {

                throw new Exception(functionName + ex.Message);

            }

        }

        public string WhoAmI()

        {

            try

            {

                Guid orgId = ((WhoAmIResponse)service.Execute(new WhoAmIRequest())).OrganizationId;

                return orgId.ToString();

            }

            catch

            {

                throw;

            }

        }

    }

}*/