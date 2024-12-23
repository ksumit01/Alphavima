using DocumentFormat.OpenXml.Packaging;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Activities.Statements;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json.Serialization;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Office.Drawing;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Newtonsoft.Json;

namespace AV.Prexa.TransportTicket.PopulateWordTemplate
{
    public class PopulateWordTemplate : CodeActivity
    {
        [Input("DocTemplateId")]
        [ArgumentRequired]
        public InArgument<string> DocTemplateId { get; set; }


        [Input("RecordId")]
        [ArgumentRequired]
        public InArgument<string> RecordId { get; set; }

        [Output("Base64Pdf")]
        public OutArgument<string> Base64Pdf { get; set; }

        [Output("ErrorMessage")]
        public OutArgument<string> ErrorMessage { get; set; }
        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            tracingService.Trace("Execute method started.");

            try
            {

                string docTemplateID = DocTemplateId.Get(executionContext);
                tracingService.Trace($"DocTemplateId: {docTemplateID}");

                string recordId = RecordId.Get(executionContext);
                tracingService.Trace($"RecordId: {recordId}");

                if (string.IsNullOrEmpty(docTemplateID))
                {
                    tracingService.Trace("DocTemplateId is null or empty.");
                    throw new ArgumentNullException(nameof(docTemplateID), "DocTemplateId cannot be null or empty.");
                }

                string errorMessage = string.Empty;
                string base64Pdf = RetrieveEntityColumnData(service, new Guid(docTemplateID), recordId, tracingService, context).Result;

                tracingService.Trace("Data retrieved successfully.");

                Base64Pdf.Set(executionContext, base64Pdf);
                ErrorMessage.Set(executionContext, errorMessage);
            }
            catch (InvalidWorkflowException ex)
            {
                tracingService.Trace("InvalidWorkflowException: Execute >> " + ex.ToString());
                throw;
            }
            catch (Exception ex1)
            {
                tracingService.Trace("Exception: Execute >> " + ex1.ToString());
                throw;
            }

            tracingService.Trace("Execute method ended.");
        }

        private async Task<string> RetrieveEntityColumnData(IOrganizationService service, Guid docTemplateID, string recordID, ITracingService tracingService, IWorkflowContext context)
        {
            tracingService.Trace("RetrieveEntityColumnData started.");
            tracingService.Trace($"DocTemplateID: {docTemplateID}, RecordID: {recordID}");

            if (service == null)
            {
                tracingService.Trace("Service is null.");
                throw new ArgumentNullException(nameof(service), "Service cannot be null.");
            }

            string entityName = "avpx_relatedconfigdata";
            string fileColumnNames = "avpx_file";
            string powerAutomateUrl = null;

            try
            {
                tracingService.Trace("Retrieving Power Automate URL dynamically.");
                powerAutomateUrl = FetchWordToPDFPowerAutomateURL(service, tracingService);
                tracingService.Trace($"Retrieved Power Automate URL: {powerAutomateUrl}");

                if (string.IsNullOrEmpty(powerAutomateUrl))
                {
                    tracingService.Trace("Power Automate URL is null or empty.");
                    throw new InvalidOperationException("Power Automate URL cannot be null or empty.");
                }

                tracingService.Trace($"Retrieving entity '{entityName}' with DocTemplate ID '{docTemplateID}' and column '{fileColumnNames}'.");
                Entity entity = service.Retrieve(entityName, docTemplateID, new ColumnSet(fileColumnNames));

                if (entity == null)
                {
                    tracingService.Trace("Entity not found.");
                    throw new InvalidOperationException("Entity not found.");
                }

                tracingService.Trace("Entity retrieved successfully.");

                byte[] fileData = null;

                if (entity.Contains("avpx_file"))
                {
                    tracingService.Trace($"Downloading file data for entity '{entityName}' with ID '{docTemplateID}'.");
                    fileData = DownloadFileData(service, entityName, docTemplateID, fileColumnNames);
                    tracingService.Trace("File data downloaded successfully.");
                }
                else
                {
                    tracingService.Trace("Entity does not contain 'avpx_file' attribute.");
                }

                if (fileData != null)
                {
                    tracingService.Trace("File data found.");

                    // Extract placeholders from file data
                    // var placeholders = ExtractPlaceholders(fileData, tracingService);

                    // Populate placeholders with data from FetchXML
                    // var placeholderData = PopulatePlaceholderData(service, placeholders, tracingService);

                    // foreach (var kvp in placeholderData)
                    // {
                    //     tracingService.Trace($"Key: {kvp.Key}, Value: {kvp.Value}");
                    // }

                    tracingService.Trace("Placeholder data processed successfully.");

                    var placeholder = extracttransportTicketData(recordID, service, tracingService, context);

                    tracingService.Trace("Starting Fetch Line Item");
                    var data = RetrieveAndReturnRecords(service, tracingService, recordID);
                    tracingService.Trace("End Fetch Line Item");

                    tracingService.Trace("Processing Word template.");
                    byte[] updatedWordData = WordTemplate(fileData, null, placeholder, data, tracingService);
                    tracingService.Trace("Word template processed successfully.");

                    // Convert updated Word data to Base64 string
                    string updatedWordDataString = Convert.ToBase64String(updatedWordData);

                    // Post processed data to Power Automate
                    tracingService.Trace($"Posting processed Word data to Power Automate URL: {powerAutomateUrl}");
                    return await PostBase64ToPowerAutomate(updatedWordDataString, powerAutomateUrl);
                }
                else
                {
                    tracingService.Trace("No file data found.");
                }
            }
            catch (ArgumentNullException ex)
            {
                tracingService.Trace($"ArgumentNullException: {ex.Message}");
                throw;
            }
            catch (InvalidOperationException ex)
            {
                tracingService.Trace($"InvalidOperationException: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                tracingService.Trace($"Exception in RetrieveEntityColumnData: {ex.Message}");
                throw;
            }

            return null;
        }

        private string FetchWordToPDFPowerAutomateURL(IOrganizationService service, ITracingService tracingService)
        {
            string fetchXML = @"
    <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        <entity name='avpx_prexa365configuration'>
            <attribute name='avpx_wordtopdfpowerautomateurl' />
            <order attribute='avpx_name' descending='false' />
            <filter type='and'>
                <condition attribute='statuscode' operator='eq' value='1' />
                <condition attribute='avpx_name' operator='eq' value='Default Settings' />
            </filter>
        </entity>
    </fetch>";

            tracingService.Trace("Executing FetchXML query to retrieve Power Automate URL.");
            //tracingService.Trace($"FetchXML: {fetchXML}");

            EntityCollection results;
            try
            {
                results = service.RetrieveMultiple(new FetchExpression(fetchXML));
                tracingService.Trace($"RetrieveMultiple executed successfully. Number of records retrieved: {results.Entities.Count}");
            }
            catch (Exception ex)
            {
                tracingService.Trace($"Error executing RetrieveMultiple: {ex.Message}");
                throw new InvalidPluginExecutionException($"An error occurred while retrieving the Power Automate URL: {ex.Message}", ex);
            }

            if (results.Entities.Count > 0)
            {
                tracingService.Trace("At least one record found. Extracting the Power Automate URL.");
                Entity entity = results.Entities[0];
                if (entity.Contains("avpx_wordtopdfpowerautomateurl"))
                {
                    string powerAutomateUrl = entity["avpx_wordtopdfpowerautomateurl"].ToString();
                    tracingService.Trace($"Power Automate URL found: {powerAutomateUrl}");
                    return powerAutomateUrl;
                }
                else
                {
                    tracingService.Trace("The attribute 'avpx_wordtopdfpowerautomateurl' was not found in the retrieved entity.");
                }
            }
            else
            {
                tracingService.Trace("No records found matching the FetchXML query.");
            }

            tracingService.Trace("Returning empty string as no valid URL was found.");
            return string.Empty;
        }

        public int GetUserTimezone(Guid userid, IOrganizationService service, ITracingService tracingService)
        {
            try
            {
                tracingService.Trace($"GetUserTimezone: Retrieving timezone for user ID: {userid}");
                RetrieveUserSettingsSystemUserRequest userSettingsSystemUserRequest = new RetrieveUserSettingsSystemUserRequest
                {
                    ColumnSet = new ColumnSet(new string[] { "timezonecode" }),
                    EntityId = userid
                };
                RetrieveUserSettingsSystemUserResponse userSettings = (RetrieveUserSettingsSystemUserResponse)service.Execute(userSettingsSystemUserRequest);

                // Get the time zone of the user
                return Convert.ToInt16(userSettings.Entity["timezonecode"].ToString());
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error retrieving user timezone: " + ex.ToString());
                return 255; // Default to UTC if unable to retrieve timezone
            }
        }

        public DateTime GetLocalTime(DateTime utcTime, int timezoneCode, IOrganizationService service, ITracingService tracingService)
        {
            try
            {
                tracingService.Trace($"GetLocalTime: Converting UTC time {utcTime} to local time for timezone code {timezoneCode}");

                // Ensure the UTC time is in the correct format
                var request = new LocalTimeFromUtcTimeRequest
                {
                    TimeZoneCode = timezoneCode,
                    UtcTime = DateTime.SpecifyKind(utcTime, DateTimeKind.Utc) // Ensure the DateTime is specified as UTC
                };

                var response = (LocalTimeFromUtcTimeResponse)service.Execute(request);
                tracingService.Trace($"Local time: {response.LocalTime}");
                return response.LocalTime;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error converting UTC to local time: " + ex.ToString());
                throw;
            }
        }

        public string GetTransportTicketFetchXml(Guid transportTicketId, int ticketType)
        {
            string dynamicLinkEntity = "";

           /* switch (ticketType)
            {
                case 783090000: // Inhouse
                    dynamicLinkEntity = @"
        <link-entity name='contact' from='contactid' to='avpx_driver' link-type='inner' alias='ab'>
            <attribute name='fullname' />
        </link-entity>";
                    break;
                case 783090001: // Outsourced
                    dynamicLinkEntity = @"
        <attribute name='avpx_driverdetails' />";
                    break;
                // Add more cases here if needed
                default:
                    throw new ArgumentException("Invalid ticket type");
            }
*/

            string fetchXml = string.Format(@"
<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
    <entity name='avpx_ticket'>
        <attribute name='avpx_name' />
        <attribute name='avpx_ticketid' />
        <attribute name='avpx_branchaddress' />
        <attribute name='avpx_transporttype' />
        <attribute name='avpx_dispatchdate' />
        <attribute name='avpx_pickupdate' />
        <attribute name='avpx_tickettype' />
        <attribute name='avpx_driverdetails' />
        <attribute name='avpx_purpose' />
        <order attribute='createdon' descending='true' />
        <filter type='and'>
            <condition attribute='statuscode' operator='eq' value='1' />
            <condition attribute='avpx_ticketid' operator='eq' value='{0}' />
        </filter>
<link-entity name='avpx_rentalreservation' from='avpx_rentalreservationid' to='avpx_rentalorder' link-type='inner' alias='ai' >
            <attribute name='avpx_jobsitestreetaddress' />
            <attribute name='avpx_jobsitecity' />
            <attribute name='avpx_jobsitestateorprovince' />
            <attribute name='avpx_jobsitepostalcode' />
            <attribute name='avpx_jobsitecountry' />
            <attribute name='avpx_brl_streetaddress' />
            <attribute name='avpx_brl_city' />
            <attribute name='avpx_brl_stateorprovince' />
            <attribute name='avpx_brl_postalcode' />
            <attribute name='avpx_brl_country' />
        </link-entity>
        {1}
    </entity>
</fetch>", transportTicketId, dynamicLinkEntity);

            return fetchXml;


        }

        private Dictionary<string, string> extracttransportTicketData(string transportTicketId, IOrganizationService service, ITracingService tracingService, IWorkflowContext context)
        {
            var placeholders = new Dictionary<string, string>
    {
        {"pxStreetAddress", ""},
        {"pxCity", ""},
        {"pxState", ""},
        {"pxPostalCode", ""},
        {"pxCountry", ""},
        {"pxprintedOn", "" },
        {"pxCustomerName", ""},
        {"pxCellNo", "N/A"},
        {"pxcusAdd", ""},
        {"pxcustomerCit", ""},
        {"pxcusS", ""},
        {"pxcusP", ""},
        {"pxcusCo", ""},
        {"pxtransportdate","" },
        {"pxdriverName","" },
        {"pxTransportType",""},
        {"pxNotes","" }


    };

            try
            {
                // Initial fetch to get the order type
                string initialFetch = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' >
    <entity name='avpx_ticket' >
        <attribute name='avpx_tickettype' />
        <attribute name='avpx_transporttype' />
        <order attribute='createdon' descending='true' />
        <filter type='and' >
            <condition attribute='statuscode' operator='eq' value='1' />
            <condition attribute='avpx_ticketid' operator='eq' value='{0}' />
        </filter>


    </entity>
</fetch>", transportTicketId);

                tracingService.Trace("Initial Fetch Expression: " + initialFetch);
                EntityCollection initialEntityCollection;
                try
                {
                    initialEntityCollection = service.RetrieveMultiple(new FetchExpression(initialFetch));
                }
                catch (Exception ex)
                {
                    tracingService.Trace("Error retrieving initial transportTicket data: " + ex.ToString());
                    throw new InvalidPluginExecutionException("Failed to retrieve initial transportTicket data. " + ex.Message);
                }

                if (initialEntityCollection == null || initialEntityCollection.Entities == null || initialEntityCollection.Entities.Count == 0)
                {
                    tracingService.Trace("No transportTickets found with the given ID.");
                    return placeholders;
                }

                Entity initiaTransportTicket = initialEntityCollection.Entities[0];
                int? ticketType = initiaTransportTicket.Contains("avpx_tickettype") ? ((OptionSetValue)initiaTransportTicket["avpx_tickettype"]).Value : (int?)null;
                int? transportType = initiaTransportTicket.Contains("avpx_transporttype") ? ((OptionSetValue)initiaTransportTicket["avpx_transporttype"]).Value : (int?)null;

                tracingService.Trace("Ticket Type:" + (ticketType.HasValue ? ticketType.Value.ToString() : "null"));
                tracingService.Trace("Transport Type:" + (transportType.HasValue ? transportType.Value.ToString() : "null"));

                // Generate the dynamic FetchXML based on the order type
                string dynamicFetchXml = GetTransportTicketFetchXml(new Guid(transportTicketId), ticketType.HasValue ? ticketType.Value : 0);

                tracingService.Trace("Dynamic Fetch Expression: " + dynamicFetchXml);
                EntityCollection transportTicketEntityCollection;
                try
                {
                    transportTicketEntityCollection = service.RetrieveMultiple(new FetchExpression(dynamicFetchXml));
                }
                catch (Exception ex)
                {
                    tracingService.Trace("Error retrieving transportTicket data: " + ex.ToString());
                    throw new InvalidPluginExecutionException("Failed to retrieve transportTicket data. " + ex.Message);
                }

                if (transportTicketEntityCollection == null || transportTicketEntityCollection.Entities == null || transportTicketEntityCollection.Entities.Count == 0)
                {
                    tracingService.Trace("No transportTickets found with the given ID.");
                    return placeholders;
                }

                Entity transportTicket = transportTicketEntityCollection.Entities[0];

                tracingService.Trace("transportTicket Retrieved Successfully: " + transportTicket);

                // Helper methods to get values from the entity
                string GetValueOrDefault(string attribute, string defaultValue = "")
                {
                    try
                    {
                        if (transportTicket.Contains(attribute))
                        {
                            var value = transportTicket[attribute];
                            if (value is EntityReference entityRef)
                            {
                                return entityRef.Name ?? defaultValue;
                            }
                            if (value is Money moneyValue)
                            {
                                return moneyValue.Value.ToString("N2"); // Format money values with two decimal places
                            }
                            if (value is DateTime dateValue)
                            {
                                return dateValue.ToString("MM/dd/yyyy"); // Format dates as MM/dd/yyyy
                            }
                            return value?.ToString() ?? defaultValue;
                        }
                        return defaultValue;
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace("Error retrieving value for attribute " + attribute + ": " + ex.ToString());
                        return defaultValue;
                    }
                }

                string GetFormattedValueOrDefault(string attribute, string defaultValue = "")
                {
                    try
                    {
                        if (transportTicket.FormattedValues.Contains(attribute))
                        {
                            var value = transportTicket.FormattedValues[attribute];
                            if (decimal.TryParse(value, out decimal decimalValue))
                            {
                                return decimalValue.ToString("N2"); // Format decimal values with two decimal places
                            }
                            return value;
                        }
                        return defaultValue;
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace("Error retrieving formatted value for attribute " + attribute + ": " + ex.ToString());
                        return defaultValue;
                    }
                }

                string GetAliasedValueOrDefault(string attribute, string defaultValue = "")
                {
                    try
                    {
                        if (transportTicket.Contains(attribute) && transportTicket[attribute] is AliasedValue aliasedValue)
                        {
                            if (aliasedValue.Value is Money moneyValue)
                            {
                                return moneyValue.Value.ToString("N2"); // Format money values with two decimal places
                            }
                            if (aliasedValue.Value is DateTime dateValue)
                            {
                                tracingService.Trace(attribute + ": " + dateValue.ToString());
                                return dateValue.ToString("MM/dd/yyyy"); // Format dates as MM/dd/yyyy
                            }
                            return aliasedValue.Value?.ToString() ?? defaultValue;
                        }
                        return defaultValue;
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace("Error retrieving aliased value for attribute " + attribute + ": " + ex.ToString());
                        return defaultValue;
                    }
                }

                // Populate placeholders


                int userTimeZone = GetUserTimezone(context.UserId, service, tracingService);
                tracingService.Trace($"User TimeZone: {userTimeZone}");


                DateTime printedOnUtc = DateTime.Parse(DateTime.Now.ToString("o"));
                DateTime dispatchDateOnUtc = DateTime.Parse(GetValueOrDefault("avpx_dispatchdate", DateTime.UtcNow.ToString("o")));
                DateTime pickUpDateOnUtc = DateTime.Parse(GetValueOrDefault("avpx_pickupdate", DateTime.UtcNow.ToString("o")));

                tracingService.Trace($"Printed On Date (UTC): {printedOnUtc}");


                DateTime printedOnUtcLocal = GetLocalTime(printedOnUtc, userTimeZone, service, tracingService);
                DateTime dispatchDateOnUtcLocal = GetLocalTime(dispatchDateOnUtc, userTimeZone, service, tracingService);
                DateTime pickUpDateOnUtcLocal = GetLocalTime(pickUpDateOnUtc, userTimeZone, service, tracingService);


                tracingService.Trace($"Printed On Date (Local): {printedOnUtcLocal}");
                tracingService.Trace($"Dispatch Date (Local): {dispatchDateOnUtcLocal}");
                tracingService.Trace($"Pickup Date (Local): {pickUpDateOnUtcLocal}");



                placeholders["pxTransportType"] = GetValueAsString(transportTicket, "avpx_transporttype", service, tracingService);
                placeholders["pxprintedOn"] = printedOnUtcLocal.ToString("MM/dd/yyyy");

                /* if (ticketType.HasValue)
                 {
                     if (ticketType == 783090000) //Inhouse
                     {
                         placeholders["pxdriverName"] = GetAliasedValueOrDefault("ab.fullname");
                     }
                     else if (ticketType == 783090001)
                     {
                         placeholders["pxdriverName"] = GetValueOrDefault("avpx_driverdetails");
                     }
                 }
                 else
                 {
                     placeholders["driverName"] = "N/A";
                     tracingService.Trace("Ticket Type is null or not defined, Not Available.");
                 }
 */
                placeholders["pxdriverName"] = GetValueAsString(transportTicket, "avpx_driverdetails", service, tracingService);
                if (transportType == 783090001) //Pickup
                {
                    placeholders["pxStreetAddress"] = GetAliasedValueOrDefault("ai.avpx_brl_streetaddress");
                    placeholders["pxCity"] = GetAliasedValueOrDefault("ai.avpx_brl_city");
                    placeholders["pxState"] = string.IsNullOrEmpty(GetAliasedValueOrDefault("ai.avpx_brl_stateorprovince")) ? "" : GetAliasedValueOrDefault("ai.avpx_brl_stateorprovince") + ",";
                    placeholders["pxPostalCode"] = GetAliasedValueOrDefault("ai.avpx_brl_postalcode");
                    placeholders["pxCountry"] = GetAliasedValueOrDefault("ai.avpx_brl_country");

                    placeholders["pxcusAdd"] = GetAliasedValueOrDefault("ai.avpx_jobsitestreetaddress");
                    placeholders["pxcustomerCit"] = GetAliasedValueOrDefault("ai.avpx_jobsitecity");
                    placeholders["pxcusS"] = string.IsNullOrEmpty(GetAliasedValueOrDefault("ai.avpx_jobsitestateorprovince")) ? "" : GetAliasedValueOrDefault("ai.avpx_jobsitepostalcode") + ",";
                    placeholders["pxcusP"] = GetAliasedValueOrDefault("ai.avpx_jobsitepostalcode");
                    placeholders["pxcusCo"] = GetAliasedValueOrDefault("ai.avpx_jobsitecountry");
                    placeholders["pxtransportDate"] = dispatchDateOnUtcLocal.ToString("MM/dd/yyyy");
                    placeholders["pxNotes"] = GetValueOrDefault("avpx_purpose");
                    
                }
                else if (transportType == 783090000) //Dispatch
                {
                    placeholders["pxStreetAddress"] = GetAliasedValueOrDefault("ai.avpx_jobsitestreetaddress");
                    placeholders["pxCity"] = GetAliasedValueOrDefault("ai.avpx_jobsitecity");
                    placeholders["pxState"] = string.IsNullOrEmpty(GetAliasedValueOrDefault("ai.avpx_jobsitestateorprovince")) ? "" : GetAliasedValueOrDefault("ai.avpx_jobsitestateorprovince") + ",";
                    placeholders["pxPostalCode"] = GetAliasedValueOrDefault("ai.avpx_jobsitepostalcode");
                    placeholders["pxCountry"] = GetAliasedValueOrDefault("ai.avpx_jobsitecountry");
                    placeholders["pxtransportDate"] = pickUpDateOnUtcLocal.ToString("MM/dd/yyyy");
                    placeholders["pxcusAdd"] = GetAliasedValueOrDefault("ai.avpx_brl_streetaddress");
                    placeholders["pxcustomerCit"] = GetAliasedValueOrDefault("ai.avpx_brl_city");
                    placeholders["pxcusS"] = String.IsNullOrEmpty(GetValueOrDefault("ai.avpx_brl_stateorprovince")) ? "" : GetAliasedValueOrDefault("ai.avpx_brl_stateorprovince") + ",";
                    placeholders["pxcusP"] = GetAliasedValueOrDefault("ai.avpx_brl_postalcode");
                    placeholders["pxcusCo"] = GetAliasedValueOrDefault("ai.avpx_brl_country");
                    placeholders["pxNotes"] = GetValueOrDefault("avpx_purpose");
                    //placeholders["transportTicketNumber"] = transportTicketNo;
                }
                //placeholders["OffNumber"] = GetAliasedValueOrDefault("ac.telephone1", "N/A");
                //placeholders["printedOn"] = DateTime.Now.ToString("MM/dd/yyyy");
                // placeholders["CustomerName"] = GetValueOrDefault("avpx_customer");
                //placeholders["CellNo"] = GetAliasedValueOrDefault("ac.telephone3", "N/A");

                //placeholders["BillFrom"] = GetValueOrDefault("avpx_billfromdate");
                //placeholders["billto"] = GetValueOrDefault("avpx_billto");

                foreach (var item in placeholders)
                {
                    tracingService.Trace($"Key: {item.Key} Value: {item.Value}");
                }

                return placeholders;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error in extracting transportTicket data: " + ex.ToString());
                throw new InvalidPluginExecutionException("Failed to extract transportTicket data. " + ex.Message);
            }
        }

        private static byte[] DownloadFileData(IOrganizationService service, string entityName, Guid recordGuid, string fileAttributeName)
        {
            {
                var initializeFileBlocksDownloadRequest = new InitializeFileBlocksDownloadRequest
                {
                    Target = new EntityReference(entityName, recordGuid),
                    FileAttributeName = fileAttributeName
                };

                var initializeFileBlocksDownloadResponse = (InitializeFileBlocksDownloadResponse)
                    service.Execute(initializeFileBlocksDownloadRequest);

                byte[] fileData = new byte[0];
                DownloadBlockRequest downloadBlockRequest = new DownloadBlockRequest
                {
                    FileContinuationToken = initializeFileBlocksDownloadResponse.FileContinuationToken
                };

                var downloadBlockResponse = (DownloadBlockResponse)service.Execute(downloadBlockRequest);

                // Creates a new file, writes the specified byte array to the file,
                // and then closes the file. If the target file already exists, it is overwritten.

                /*File.WriteAllBytes(filePath +
                    initializeFileBlocksDownloadResponse.FileName,
                    downloadBlockResponse.Data);*/

                return downloadBlockResponse.Data;
            }
        }

        public async Task<string> PostBase64ToPowerAutomate(string base64String, string powerAutomateUrl)
        {
            string fileName = GenerateFileName();

            var payload = new
            {
                fileName = fileName,
                fileContent = base64String
            };

            using (var client = new HttpClient())
            {
                var content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");
                HttpResponseMessage response = await client.PostAsync(powerAutomateUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    string responseBody = await response.Content.ReadAsStringAsync();
                    return responseBody;
                }
                else
                {
                    throw new Exception($"Error: {response.StatusCode}");
                }
            }
        }

        private string GenerateFileName()
        {
            string timestamp = DateTime.UtcNow.ToString("ddMMyyyy");
            Random random = new Random();
            string randomNumber = random.Next(1000, 10000).ToString();
            return $"{timestamp}_{randomNumber}.docx";
        }

        static byte[] WordTemplate(byte[] fileData, byte[] imageData, Dictionary<string, string> placeholders, List<string[]> combinedData, ITracingService tracingService)
        {
            tracingService.Trace("Starting WordTemplate method.");

            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    memoryStream.Write(fileData, 0, fileData.Length);
                    memoryStream.Position = 0;

                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(memoryStream, true))
                    {
                        if (wordDoc.MainDocumentPart == null)
                        {
                            throw new InvalidOperationException("The document does not contain a main document part.");
                        }

                        // Replace placeholders in the main document part
                        string docText;
                        using (StreamReader reader = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                        {
                            docText = reader.ReadToEnd();
                        }

                        tracingService.Trace("Main document part read successfully.");
                        docText = docText.Replace("{{", "")
                                        .Replace("}}", "")
                                        .Replace("}},", ",")
                                        .Replace("},{{", ",")
                                        .Replace("}},", "")
                                        .Replace(",{{", "")
                                        .Replace("}},", "")
                                        .Replace(",", "");

                        foreach (var placeholder in placeholders)
                        {
                            tracingService.Trace($"Replacing placeholder: {placeholder.Key} with {placeholder.Value}");
                            if (placeholder.Key == "cusS")
                            {
                                tracingService.Trace(placeholder.Value);
                            }

                            //tracingService.Trace(docText);

                            docText = docText.Replace(placeholder.Key, System.Net.WebUtility.HtmlEncode(placeholder.Value));
                        }

                        using (StreamWriter writer = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                        {
                            writer.Write(docText);
                            tracingService.Trace("Modified document text written.");
                        }

                        // Split combined data into serialized and non-serialized sections
                        var serializedData = new List<string[]>();
                        var nonSerializedData = new List<string[]>();
                        bool isSerializedSection = false;

                        tracingService.Trace("Splitting combined data into serialized and non-serialized sections.");

                        foreach (var record in combinedData)
                        {
                            if (record.Length == 1 && record[0] == "---- End of Serialized Items ----")
                            {
                                isSerializedSection = true;
                                continue;
                            }

                            if (isSerializedSection)
                            {
                                nonSerializedData.Add(record);
                            }
                            else
                            {
                                serializedData.Add(record);
                            }
                        }

                        tracingService.Trace($"Data split completed. Serialized count: {serializedData.Count}, Non-Serialized count: {nonSerializedData.Count}");

                        // Find tables by caption
                        var tables = wordDoc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList();
                        bool serializedTableFound = false;
                        bool nonSerializedTableFound = false;

                        foreach (var table in tables)
                        {
                            var tableProperties = table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableProperties>().FirstOrDefault();
                            var tableCaption = tableProperties?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCaption>().FirstOrDefault();

                            if (tableCaption != null)
                            {
                                var altText = tableCaption.Val.Value;
                                tracingService.Trace($"Found table with alt text: {altText}");

                                if (altText.Contains("AssetGroupSer"))
                                {
                                    tracingService.Trace("Serialized table identified by alt text.");
                                    PopulateTableWithData(table, serializedData, imageData, wordDoc, tracingService);
                                    serializedTableFound = true;
                                }
                                else if (altText.Contains("AssetGroupNS"))
                                {
                                    tracingService.Trace("Non-Serialized table identified by alt text.");
                                    PopulateTableWithData(table, nonSerializedData, imageData, wordDoc, tracingService);
                                    nonSerializedTableFound = true;
                                }
                            }
                        }
                        wordDoc.MainDocumentPart.Document.Save();
                        tracingService.Trace("Document saved successfully.");
                    }

                    memoryStream.Position = 0;
                    return memoryStream.ToArray();
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace($"An error occurred while populating the template: {ex.Message}");
                tracingService.Trace($"Stack Trace: {ex.StackTrace}");
                throw;
            }
        }

        static void PopulateTableWithData(DocumentFormat.OpenXml.Wordprocessing.Table table, List<string[]> tableData, byte[] imageData, WordprocessingDocument wordDoc, ITracingService tracingService)
        {
            tracingService.Trace("Starting PopulateTableWithData.");

            foreach (var rowData in tableData)
            {
                tracingService.Trace("Adding new row.");

                DocumentFormat.OpenXml.Wordprocessing.TableRow newRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();

                foreach (var cellData in rowData)
                {
                    tracingService.Trace($"Adding new cell with data: {cellData}");

                    DocumentFormat.OpenXml.Wordprocessing.TableCell newCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                        new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                            new DocumentFormat.OpenXml.Wordprocessing.Run(
                                new DocumentFormat.OpenXml.Wordprocessing.Text(cellData)
                            )
                        )
                    );
                    newRow.Append(newCell);
                }

                if (imageData != null)
                {
                    tracingService.Trace("Adding image to row.");

                    DocumentFormat.OpenXml.Wordprocessing.TableCell imageCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                    AddImageToCell(wordDoc, imageCell, imageData, tracingService);
                    newRow.Append(imageCell);
                }

                table.Append(newRow);
            }

            tracingService.Trace("Completed PopulateTableWithData.");
        }

        static void AddImageToCell(WordprocessingDocument wordDoc, DocumentFormat.OpenXml.Wordprocessing.TableCell cell, byte[] imageData, ITracingService tracingService)
        {
            try
            {
                tracingService.Trace("AddImageToCell method started.");

                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                // Add the image part to the main document part
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                // Write the image data to the image part
                using (MemoryStream stream = new MemoryStream(imageData))
                {
                    imagePart.FeedData(stream);
                }

                // Get the image part ID
                string imagePartId = mainPart.GetIdOfPart(imagePart);

                // Construct the Drawing object for the image
                var drawing = new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = 990000L, Cy = 792000L },  // Set the size of the image
                        new DW.EffectExtent()
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties()
                        {
                            Id = 1U,
                            Name = "Picture 1"
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                new A.GraphicData(
                new Pic.Picture(
                                    new Pic.NonVisualPictureProperties(
                                        new Pic.NonVisualDrawingProperties()
                                        {
                                            Id = 0U,
                                            Name = "New Bitmap Image.jpg"
                                        },
                                        new Pic.NonVisualPictureDrawingProperties()),
                                    new Pic.BlipFill(
                                        new A.Blip()
                                        {
                                            Embed = imagePartId,
                                            CompressionState = A.BlipCompressionValues.Print
                                        },
                                        new A.Stretch(
                                            new A.FillRectangle())),
                                    new Pic.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset() { X = 0L, Y = 0L },
                                            new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                        new A.PresetGeometry(
                                            new A.AdjustValueList())
                                        { Preset = A.ShapeTypeValues.Rectangle }
                                    )
                                )
                            )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                        )
                    )
                    {
                        DistanceFromTop = 50000U,
                        DistanceFromBottom = 50000U,
                        DistanceFromLeft = 0U,
                        DistanceFromRight = 0U,
                        EditId = "50D07946"
                    });

                // Create a paragraph and append the drawing to it
                DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(drawing));

                // Append the paragraph to the cell
                cell.Append(paragraph);

                tracingService.Trace("Image added to cell successfully.");
            }
            catch (Exception ex)
            {
                tracingService.Trace($"An error occurred while adding image to cell: {ex.Message}");
                throw;
            }
            finally
            {
                tracingService.Trace("AddImageToCell method ended.");
            }
        }

        public static string GenerateFetchXml(string entityName, Dictionary<string, string> placeholders, string id, ITracingService tracingService)
        {


            StringBuilder xmlBuilder = new StringBuilder();
            tracingService.Trace("Generating Fetch XML Started");
            xmlBuilder.AppendLine("<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>");
            xmlBuilder.AppendLine($"    <entity name='{entityName}'>");

            foreach (var placeholder in placeholders)
            {
                xmlBuilder.AppendLine($"        <attribute name='{placeholder.Value}' />");
            }

            xmlBuilder.AppendLine("    <filter type='and'>");
            xmlBuilder.AppendLine($"        <condition attribute='{entityName}id' operator='eq' value='{id}' />");
            xmlBuilder.AppendLine("    </filter>");
            xmlBuilder.AppendLine("    </entity>");
            xmlBuilder.AppendLine("</fetch>");
            tracingService.Trace($" the fetch XML  = {xmlBuilder.ToString()}");
            tracingService.Trace("Returning the fetch XML Result");
            return xmlBuilder.ToString();
        }

        public List<string[]> RetrieveAndReturnRecords(IOrganizationService service, ITracingService tracingService, string recordId)
        {
            var serializedItems = new List<string[]>();
            var nonSerializedItems = new List<string[]>();

            try
            {
                // Define the FetchXML for serialized items
                string fetchXmlSerialized = $@"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='avpx_ticketitem'>
                    <attribute name='avpx_asset' />
                    <attribute name='avpx_description' />
                    <filter type='and'>
                        <condition attribute='statuscode' operator='eq' value='1' />
                        <condition attribute='avpx_transportticket' operator='eq' uitype='avpx_ticket' value='{recordId}' />
                        <condition attribute='avpx_isserialized' operator='eq' value='1' />
                    </filter>
                </entity>
            </fetch>";

                // Retrieve serialized items
                tracingService.Trace("Fetch XML for serialized items: " + fetchXmlSerialized);
                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXmlSerialized));

                // Process serialized items
                if (entityCollection.Entities.Count > 0)
                {
                    tracingService.Trace("Serialized records retrieved successfully. Count: {0}", entityCollection.Entities.Count);

                    foreach (Entity entity in entityCollection.Entities)
                    {
                        string avpxMake = "N/A";
                        string avpxModel = "N/A";
                        string avpxSerialNumber = "N/A";
                        string assetName = "N/A";
                        string description = GetValueAsString(entity, "avpx_description", service, tracingService);
                        Guid assetId = entity.GetAttributeValue<EntityReference>("avpx_asset")?.Id ?? Guid.Empty;
                        tracingService.Trace("Asset ID" + assetId);
                        if (assetId != Guid.Empty)
                        {
                            Entity assetEntity = service.Retrieve("avpx_asset", assetId, new ColumnSet("avpx_make", "avpx_model", "avpx_serialnumbervin", "avpx_name"));

                            Guid makeId = assetEntity.GetAttributeValue<EntityReference>("avpx_make")?.Id ?? Guid.Empty;
                            if (makeId != Guid.Empty)
                            {
                                Entity makeEntity = service.Retrieve("avpx_make", makeId, new ColumnSet("avpx_name"));
                                avpxMake = makeEntity.Contains("avpx_name") ? makeEntity["avpx_name"].ToString() : string.Empty;
                                tracingService.Trace("Make" + avpxMake);
                            }

                            Guid modelId = assetEntity.GetAttributeValue<EntityReference>("avpx_model")?.Id ?? Guid.Empty;
                            if (modelId != Guid.Empty)
                            {
                                Entity modelEntity = service.Retrieve("avpx_model", modelId, new ColumnSet("avpx_name"));
                                avpxModel = modelEntity.Contains("avpx_name") ? modelEntity["avpx_name"].ToString() : string.Empty;
                                tracingService.Trace("Model" + avpxModel);
                            }

                            avpxSerialNumber = GetValueOrDefault(assetEntity, "avpx_serialnumbervin", "N/A");
                            assetName = GetValueOrDefault(assetEntity, "avpx_name", "N/A");
                        }
                        string displayName = (!string.IsNullOrEmpty(description) && description != "N/A") ? description : assetName;
                        string[] recordData = new string[] { avpxSerialNumber, displayName, avpxMake, avpxModel };
                        serializedItems.Add(recordData);
                    }
                }
                else
                {
                    tracingService.Trace("No serialized records retrieved.");
                }

                // Define the FetchXML for non-serialized items
                string fetchXmlNonSerialized = $@"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='avpx_ticketitem'>
                    <attribute name='avpx_assetgroup' />
                    <attribute name='avpx_quantity' />
                     <attribute name='avpx_description' />
                    <filter type='and'>
                        <condition attribute='statuscode' operator='eq' value='1' />
                        <condition attribute='avpx_transportticket' operator='eq' uitype='avpx_ticket' value='{recordId}' />
                        <condition attribute='avpx_isserialized' operator='eq' value='0' />
                    </filter>
                </entity>
            </fetch>";

                // Retrieve non-serialized items
                tracingService.Trace("Fetch XML for non-serialized items: " + fetchXmlNonSerialized);
                entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXmlNonSerialized));

                // Process non-serialized items
                if (entityCollection.Entities.Count > 0)
                {
                    tracingService.Trace("Non-serialized records retrieved successfully. Count: {0}", entityCollection.Entities.Count);

                    foreach (Entity entity in entityCollection.Entities)
                    {
                        string avpxAssetGroupName = "N/A";
                        string avpxQuantity = "N/A";
                        string description = GetValueAsString(entity, "avpx_description", service, tracingService);
                        string displayName = (!string.IsNullOrEmpty(description) && description != "N/A") ? description : avpxAssetGroupName;
                        Guid assetGroupId = entity.GetAttributeValue<EntityReference>("avpx_assetgroup")?.Id ?? Guid.Empty;
                        if (assetGroupId != Guid.Empty)
                        {
                            Entity assetGroupEntity = service.Retrieve("avpx_device", assetGroupId, new ColumnSet("avpx_name"));
                            avpxAssetGroupName = assetGroupEntity.Contains("avpx_name") ? assetGroupEntity["avpx_name"].ToString() : string.Empty;
                            tracingService.Trace(avpxAssetGroupName);
                        }

                        avpxQuantity = GetFormattedValueOrDefault(entity, "avpx_quantity", tracingService, "N/A");

                        string[] recordData = new string[] { displayName, avpxQuantity };
                        nonSerializedItems.Add(recordData);
                    }
                }
                else
                {
                    tracingService.Trace("No non-serialized records retrieved.");
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("An error occurred: {0}", ex.Message);
            }

            // Combine lists
            var combinedResult = new List<string[]>();
            combinedResult.AddRange(serializedItems);
            combinedResult.Add(new string[] { "---- End of Serialized Items ----" });
            combinedResult.AddRange(nonSerializedItems);


            return combinedResult;
        }

        // Helper method to get value or default
        private string GetValueOrDefault(Entity entity, string attributeName, string defaultValue)
        {
            return entity.Contains(attributeName) ? entity[attributeName].ToString() : defaultValue;
        }


        private string GetValueAsString(Entity entity, string attributeName, IOrganizationService service, ITracingService tracingService)
        {
            if (entity.Contains(attributeName))
            {
                var value = entity[attributeName];
                if (value != null)
                {
                    switch (value)
                    {
                        case Money money:
                            return money.Value.ToString("N2");
                        case DateTime dateTime:
                            return dateTime.ToString("MM-dd-yyyy"); // Format as needed
                        case OptionSetValue optionSet:
                            return GetOptionSetValueLabel(service, entity.LogicalName, attributeName, optionSet.Value);
                        case EntityReference entityRef:
                            return entityRef.Name; // You may want to use entityRef.Id.ToString() or both
                        default:
                            return value.ToString();
                    }
                }
            }
            return "N/A";
        }

        private string RoundOffValue(Entity entity, string attributeName)
        {
            if (entity.Contains(attributeName))
            {
                var value = entity[attributeName];
                if (value != null)
                {
                    if (value is decimal decimalValue)
                    {
                        return decimalValue.ToString("F2");
                    }
                    else if (value is double doubleValue)
                    {
                        return doubleValue.ToString("F2");
                    }
                    else if (decimal.TryParse(value.ToString(), out decimal parsedDecimalValue))
                    {
                        return parsedDecimalValue.ToString("F2");
                    }
                }
            }
            return "0.00";
        }

        private string GetFormattedValueOrDefault(Entity entity, string attributeName, ITracingService tracingService, string defaultValue = "N/A")
        {
            try
            {
                if (entity.FormattedValues.Contains(attributeName))
                {
                    var value = entity.FormattedValues[attributeName];
                    if (decimal.TryParse(value, out decimal decimalValue))
                    {
                        return decimalValue.ToString("N2"); // Format decimal values with two decimal places
                    }
                    return value;
                }
                return defaultValue;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error retrieving formatted value for attribute " + attributeName + ": " + ex.ToString());
                return defaultValue;
            }
        }

        private string GetOptionSetValueLabel(IOrganizationService service, string entityLogicalName, string attributeName, int optionSetValue)
        {
            RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityLogicalName,
                LogicalName = attributeName,
                RetrieveAsIfPublished = true
            };

            RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
            PicklistAttributeMetadata picklistMetadata = (PicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;

            foreach (var option in picklistMetadata.OptionSet.Options)
            {
                if (option.Value == optionSetValue)
                {
                    return option.Label.UserLocalizedLabel.Label;
                }
            }

            return optionSetValue.ToString(); // Fallback to value if label not found
        }
    }
}
