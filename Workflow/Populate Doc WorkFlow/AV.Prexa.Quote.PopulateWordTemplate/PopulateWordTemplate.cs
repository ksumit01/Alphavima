using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using System.Net.Http.Headers;
using System.Net.Http;
using Newtonsoft.Json;
using System.IO;

using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Office.Drawing;
using Microsoft.Xrm.Tooling.Connector;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace AV.Prexa.Quote.PopulateWordTemplate
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

        [Output("Base64Word")]
        public OutArgument<string> Base64Word { get; set; } // New output for the Word byte array

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
                (string base64Pdf, string base64Word) = RetrieveEntityColumnData(service, new Guid(docTemplateID), recordId, tracingService, context).Result;

                tracingService.Trace("Data retrieved successfully.");

                Base64Pdf.Set(executionContext, base64Pdf);
                Base64Word.Set(executionContext, base64Word); // Set the Word byte array output
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

        private async Task<(string base64Pdf, string base64Word)> RetrieveEntityColumnData(IOrganizationService service, Guid docTemplateID, string recordID, ITracingService tracingService, IWorkflowContext context)
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

                    var placeholder = extractQuoteData(recordID, service, tracingService, context);

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
                    string base64Pdf = await PostBase64ToPowerAutomate(updatedWordDataString, powerAutomateUrl, tracingService);

                    return (base64Pdf, updatedWordDataString); // Return both the PDF and Word data
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

            return (null, null);
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

        private Dictionary<string, string> extractQuoteData(string quoteId, IOrganizationService service, ITracingService tracingService, IWorkflowContext context)
        {
            var placeholders = new Dictionary<string, string>
    {
        {"pxquotenumber", ""},
        {"pxCustomerNo", "N/A"},
        {"pxEffectiveTo", ""},
        {"pxRentalOut", ""},
        {"pxRentalIn", ""},
        {"pxPONumber", ""},
        {"pxOrderedby", ""},
        {"pxsalesperson", "N/A"},
        {"pxQuoteAmount", ""},
        {"pxrentalSubtotal", ""},
        {"pxSMSubtotal", ""},
        {"pxDAmount", ""},
        {"pxDWAmount", ""},
        {"pxASubtotal", ""},
        {"pxTAmount", ""},
        {"pxTamt", ""},
        {"pxRentalSubtotal", ""},
        {"pxStreetAddress", ""},
        {"pxCity", ""},
        {"pxState", ""},
        {"pxPostalCode", ""},
        {"pxCountry", ""},
        {"pxOfficeNumber", "N/A"},
        {"pxprintedOn", ""},
        {"pxCustomerName", ""},
        {"pxCellNo", "N/A"},
        {"pxcusAdd", ""},
        {"pxcustomerCit", ""},
        {"pxcusS", ""},
        {"pxcusP", ""},
        {"pxcusCo", ""}
    };

            try
            {
                string quoteFetch = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='avpx_quote'>
            <attribute name='avpx_name' />
            <attribute name='avpx_customerscontact' />
            <attribute name='avpx_account' />
            <attribute name='avpx_quoteid' />
            <attribute name='avpx_ponumber' />
            <attribute name='avpx_detailedamount' />
            <attribute name='avpx_damagewaiveramount' />
            <attribute name='avpx_othercharges' />
            <attribute name='avpx_totalmanualdiscount' />
            <attribute name='avpx_totaltax' />
            <attribute name='avpx_totalamount' />
            <attribute name='avpx_subtotalamount' />
            <attribute name='avpx_totalrentalamountnew' />
            <attribute name='avpx_effectivefrom' />
            <attribute name='avpx_effectiveto' />
            <attribute name='avpx_streetaddress' />
            <attribute name='avpx_city' />
            <attribute name='avpx_stateorprovince' />
            <attribute name='avpx_postalcode' />
            <attribute name='avpx_country' />
            <attribute name='avpx_pricelist' />
            <attribute name='avpx_jobsiteaddress' />
            <attribute name='avpx_tentativerentalstartdate' />
            <attribute name='avpx_tentativerentalenddate' />
            <attribute name='avpx_descriptions' />
            <attribute name='ownerid' />
            <order attribute='createdon' descending='true' />
                <filter type='and'>
                <condition attribute='avpx_quoteid' operator='eq' uitype='avpx_quote' value='{0}' />
                </filter>
            <link-entity name='account' from='accountid' to='avpx_account' link-type='inner' alias='ac'>
                <attribute name='accountnumber' />
                <attribute name='name'  />
                <attribute name='avpx_searchaddress1' />
                <attribute name='address1_line2' />
                <attribute name='address1_city' />
                <attribute name='address1_stateorprovince' />
                <attribute name='address1_postalcode' />
                <attribute name='accountnumber' />
                <attribute name='telephone1' />
                <attribute name='telephone3' />
            </link-entity>
            </entity>
        </fetch>", quoteId);

                tracingService.Trace("Fetch Expression: " + quoteFetch);
                EntityCollection quoteEntityCollection;
                try
                {
                    quoteEntityCollection = service.RetrieveMultiple(new FetchExpression(quoteFetch));
                }
                catch (Exception ex)
                {
                    tracingService.Trace("Error retrieving quote data: " + ex.ToString());
                    throw new InvalidPluginExecutionException("Failed to retrieve quote data. " + ex.Message);
                }

                if (quoteEntityCollection == null || quoteEntityCollection.Entities == null || quoteEntityCollection.Entities.Count == 0)
                {
                    tracingService.Trace("No quotes found with the given ID.");
                    return placeholders;
                }

                Entity quote = quoteEntityCollection.Entities[0];
                tracingService.Trace("Quote Retrieved Successfully: " + quote);

                string GetValueOrDefault(string attribute, string defaultValue = "")
                {
                    try
                    {
                        if (quote.Contains(attribute))
                        {
                            var value = quote[attribute];
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
                                return dateValue.ToString("o"); // Format dates as round-trip date/time pattern
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
                        if (quote.FormattedValues.Contains(attribute))
                        {
                            var value = quote.FormattedValues[attribute];
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
                        return quote.Contains(attribute) ? (quote[attribute] as AliasedValue)?.Value.ToString() ?? defaultValue : defaultValue;
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace("Error retrieving aliased value for attribute " + attribute + ": " + ex.ToString());
                        return defaultValue;
                    }
                }

                // Get the user's timezone
                int userTimeZone = GetUserTimezone(context.UserId, service, tracingService);
                tracingService.Trace($"User TimeZone: {userTimeZone}");

                // Convert relevant dates to local time
                DateTime rentalOutDateUtc = DateTime.Parse(GetValueOrDefault("avpx_tentativerentalstartdate", DateTime.UtcNow.ToString("o")));
                DateTime rentalInDateUtc = DateTime.Parse(GetValueOrDefault("avpx_tentativerentalenddate", DateTime.UtcNow.ToString("o")));
                DateTime effectiveToDateUtc = DateTime.Parse(GetValueOrDefault("avpx_effectiveto", DateTime.UtcNow.ToString("o")));
                DateTime printedOnUtc = DateTime.Parse(DateTime.Now.ToString("o"));

                tracingService.Trace($"Rental Out Date (UTC): {rentalOutDateUtc}");
                tracingService.Trace($"Rental In Date (UTC): {rentalInDateUtc}");
                tracingService.Trace($"Effective To Date (UTC): {effectiveToDateUtc}");
                tracingService.Trace($"Printed On Date (UTC): {printedOnUtc}");

                DateTime rentalOutLocal = GetLocalTime(rentalOutDateUtc, userTimeZone, service, tracingService);
                DateTime rentalInLocal = GetLocalTime(rentalInDateUtc, userTimeZone, service, tracingService);
                DateTime effectiveToLocal = GetLocalTime(effectiveToDateUtc, userTimeZone, service, tracingService);
                DateTime printedOnToLocal = GetLocalTime(printedOnUtc, userTimeZone, service, tracingService);

                tracingService.Trace($"Rental Out Date (Local): {rentalOutLocal}");
                tracingService.Trace($"Rental In Date (Local): {rentalInLocal}");
                tracingService.Trace($"Effective To Date (Local): {effectiveToLocal}");
                tracingService.Trace($"Printed On Date (Local): {printedOnToLocal}");
                // Populate placeholders with data
                placeholders["pxquotenumber"] = GetValueOrDefault("avpx_name");
                placeholders["pxCustomerNo"] = GetAliasedValueOrDefault("ac.accountnumber", "N/A");
                placeholders["pxEffectiveTo"] = effectiveToLocal.ToString("dd/MM/yyyy");
                placeholders["pxRentalOut"] = rentalOutLocal.ToString("dd/MM/yyyy");
                placeholders["pxRentalIn"] = rentalInLocal.ToString("dd/MM/yyyy");
                placeholders["pxPONumber"] = GetValueOrDefault("avpx_ponumber");
                placeholders["pxOrderedby"] = GetAliasedValueOrDefault("ac.name");
                placeholders["pxsalesperson"] = GetFormattedValueOrDefault("ownerid", "N/A");
                placeholders["pxQuoteAmount"] = GetFormattedValueOrDefault("avpx_totalamount");
                placeholders["pxrentalSubtotal"] = GetFormattedValueOrDefault("avpx_subtotalamount");
                placeholders["pxSMSubtotal"] = GetFormattedValueOrDefault("avpx_othercharges");
                placeholders["pxDAmount"] = GetFormattedValueOrDefault("avpx_totalmanualdiscount");
                placeholders["pxDWAmount"] = GetFormattedValueOrDefault("avpx_damagewaiveramount");
                placeholders["pxASubtotal"] = GetFormattedValueOrDefault("avpx_subtotalamount");
                placeholders["pxTAmount"] = GetFormattedValueOrDefault("avpx_totaltax");
                placeholders["pxTamt"] = GetFormattedValueOrDefault("avpx_totalamount");
                placeholders["pxRentalSubtotal"] = GetFormattedValueOrDefault("avpx_detailedamount");
                placeholders["pxStreetAddress"] = GetValueOrDefault("avpx_streetaddress");
                placeholders["pxCity"] = GetValueOrDefault("avpx_city");
                //placeholders["State"] = GetValueOrDefault("avpx_stateorprovince");
                placeholders["pxState"] = string.IsNullOrEmpty(GetValueOrDefault("avpx_stateorprovince")) ? "" : GetValueOrDefault("avpx_stateorprovince") + ",";

                placeholders["pxPostalCode"] = GetValueOrDefault("avpx_postalcode");
                placeholders["pxCountry"] = GetValueOrDefault("avpx_country");
                placeholders["pxOfficeNumber"] = GetAliasedValueOrDefault("ac.telephone1", "N/A");
                placeholders["pxCustomerName"] = GetAliasedValueOrDefault("ac.name");
                placeholders["pxCellNo"] = GetAliasedValueOrDefault("ac.telephone3", "N/A");
                placeholders["pxcusAdd"] = GetAliasedValueOrDefault("ac.address1_line2");
                placeholders["pxcustomerCit"] = GetAliasedValueOrDefault("ac.address1_city");
                //placeholders["cusS"] = GetAliasedValueOrDefault("ac.address1_stateorprovince");
                placeholders["pxcusS"] = String.IsNullOrEmpty(GetAliasedValueOrDefault("ac.address1_stateorprovince")) ? "" : GetAliasedValueOrDefault("ac.address1_stateorprovince") + ",";
                placeholders["pxcusP"] = GetAliasedValueOrDefault("ac.address1_postalcode");
                placeholders["pxcusCo"] = GetAliasedValueOrDefault("ac.address1_country");
                placeholders["pxprintedOn"] = printedOnToLocal.ToString("dd/MM/yyyy");
                return placeholders;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error in extractQuoteData: " + ex.ToString());
                throw new InvalidPluginExecutionException("Failed to extract quote data. " + ex.Message);
            }
        }




        /* private Dictionary<string, string> ExtractPlaceholders(byte[] fileData, ITracingService tracingService)
         {
             var placeholders = new Dictionary<string, string>();
             try
             {
                 using (MemoryStream memoryStream = new MemoryStream(fileData))
                 {
                     using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(memoryStream, false))
                     {
                         var docText = new StringBuilder();
                         foreach (var text in wordDoc.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                         {
                             docText.Append(text.Text);
                         }

                         var matches = Regex.Matches(docText.ToString(), @"{{(.*?)}}");
                         foreach (Match match in matches)
                         {
                             placeholders[match.Value] = match.Groups[1].Value;
                         }
                     }
                 }
                 tracingService.Trace($"Extracted placeholders: {string.Join(", ", placeholders.Keys)}");
             }
             catch (Exception ex)
             {
                 tracingService.Trace($"Error extracting placeholders: {ex.Message}");
                 throw;
             }
             return placeholders;
         }*/



        /* private Dictionary<string, string> PopulatePlaceholderData(IOrganizationService service, Dictionary<string, string> placeholders, String recordID, ITracingService tracingService)
         {
             var placeholderData = new Dictionary<string, string>();
             try
             {
                 string fetchXml = GenerateFetchXml("avpx_quote", placeholders, recordID, tracingService);
                 EntityCollection results = service.RetrieveMultiple(new FetchExpression(fetchXml));
                 tracingService.Trace($"Number of entities retrieved: {results.Entities.Count}");

                 if (results.Entities.Count > 0)
                 {
                     Entity resultEntity = results.Entities[0];

                     foreach (var placeholder in placeholders)
                     {
                         if (resultEntity.Contains(placeholder.Value))
                         {
                             var value = resultEntity[placeholder.Value];
                             string formattedValue = string.Empty;

                             switch (value)
                             {
                                 case Money money:
                                     formattedValue = money.Value.ToString("N2");
                                     break;
                                 case DateTime dateTime:
                                     formattedValue = dateTime.ToString("MM-dd-yyyy"); // Format as needed
                                     break;
                                 case OptionSetValue optionSet:
                                     formattedValue = optionSet.Value.ToString(); // You may want to get the label instead
                                     break;
                                 default:
                                     formattedValue = value.ToString();
                                     break;
                             }

                             string cleanKey = placeholder.Key.TrimStart('{').TrimEnd('}');
                             placeholderData[cleanKey] = formattedValue;
                             tracingService.Trace($"Populated placeholder: {cleanKey} with value: {formattedValue}");
                         }
                         else
                         {
                             tracingService.Trace($"Placeholder {placeholder.Key} not found in the entity.");
                         }
                     }
                 }
                 else
                 {
                     tracingService.Trace("No entities retrieved.");
                 }
             }
             catch (Exception ex)
             {
                 tracingService.Trace($"Error populating placeholder data: {ex.Message}");
                 throw;
             }
             return placeholderData;
         }*/



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


        /*private async Task<byte[]> DownloadFileData(string fileUrl, IOrganizationService service)
        {
            using (HttpClient client = new HttpClient())
            {
                string accessToken = GetAccessToken(service);

                if (string.IsNullOrEmpty(accessToken))
                {
                    throw new InvalidOperationException("Access token is null or empty.");
                }

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                HttpResponseMessage response = await client.GetAsync(fileUrl);
                response.EnsureSuccessStatusCode();

                return await response.Content.ReadAsByteArrayAsync();
            }
        }*/

        public async Task<string> PostBase64ToPowerAutomate(string base64String, string powerAutomateUrl, ITracingService tracing)
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
                 .Replace(",", "")

                 ;


                        
                        foreach (var placeholder in placeholders)
                        {
                            
                            //string escapedValue = EscapeXmlSpecialCharacters(placeholder.Value);
                            tracingService.Trace($"Replacing placeholder: {placeholder.Key} with {placeholder.Value}");
                            docText = docText.Replace(placeholder.Key, System.Net.WebUtility.HtmlEncode(placeholder.Value));


                        }

                        using (StreamWriter writer = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                        {
                            writer.Write(docText);
                            tracingService.Trace("Modified document text written.");
                        }

                        // Split combined data into assetGroup and charges
                        var assetGroup = new List<string[]>();
                        var charges = new List<string[]>();
                        bool isChargesSection = false;

                        tracingService.Trace("Splitting combined data into assetGroup and charges.");

                        foreach (var record in combinedData)
                        {
                            if (record.Length == 1 && record[0] == "---- End of AssetGroup ----")
                            {
                                isChargesSection = true;
                                continue;
                            }

                            if (isChargesSection)
                            {
                                charges.Add(record);
                            }
                            else
                            {
                                assetGroup.Add(record);
                            }
                        }

                        tracingService.Trace($"Data split completed. AssetGroup count: {assetGroup.Count}, Charges count: {charges.Count}");

                        // Find tables by caption
                        var tables = wordDoc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList();
                        bool assetGroupTableFound = false;
                        bool chargesTableFound = false;
                        /*
                        foreach (var table in tables)
                        {
                            // Check for a paragraph immediately preceding the table that contains the caption
                            var previousParagraph = table.PreviousSibling<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
                            if (previousParagraph != null)
                            {
                                var captionText = previousParagraph.InnerText;
                                tracingService.Trace($"Found table caption: {captionText}");

                                if (captionText.Contains("AssetGroup"))
                                {
                                    tracingService.Trace("AssetGroup table identified by caption.");
                                    PopulateTableWithData(table, assetGroup, imageData, wordDoc, tracingService);
                                    assetGroupTableFound = true;
                                }
                                else if (captionText.Contains("Charges"))
                                {
                                    tracingService.Trace("Charges table identified by caption.");
                                    PopulateTableWithData(table, charges, imageData, wordDoc, tracingService);
                                    chargesTableFound = true;
                                }
                            }
                        }*/

                        foreach (var table in tables)
                        {
                            var tableProperties = table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableProperties>().FirstOrDefault();
                            var tableCaption = tableProperties?.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCaption>().FirstOrDefault();

                            if (tableCaption != null)
                            {
                                var altText = tableCaption.Val.Value;
                                tracingService.Trace($"Found table with alt text: {altText}");

                                if (altText.Contains("AssetGroup"))
                                {
                                    tracingService.Trace("AssetGroup table identified by alt text.");
                                    PopulateTableWithData(table, assetGroup, imageData, wordDoc, tracingService);
                                    assetGroupTableFound = true;
                                }
                                else if (altText.Contains("Charges"))
                                {
                                    tracingService.Trace("Charges table identified by alt text.");
                                    PopulateTableWithData(table, charges, imageData, wordDoc, tracingService);
                                    chargesTableFound = true;
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

        //Genrate Fetch XML Dynamically 
        /*public static string GenerateFetchXml(string entityName, Dictionary<string, string> placeholders, string id, ITracingService tracingService)
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
        }*/

        // Other methods like DownloadFileData, PostBase64ToPowerAutomate, etc., remain unchanged

        public List<string[]> RetrieveAndReturnRecords(IOrganizationService service, ITracingService tracingService, string recordId)
        {
            var data = new List<string[]>();
            var assetGroup = new List<string[]>();
            var charges = new List<string[]>();

            try
            {
                // Define the Fetch XML
                string fetchXml = String.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='avpx_quotelineitem'>
                                <attribute name='avpx_quotelineitemid' />
                                <attribute name='avpx_name' />
                                <attribute name='avpx_assetgroup' />
                                <attribute name='avpx_quote' />
                                <attribute name='avpx_quantity' />
                                <attribute name='avpx_priceperunit' />
                                <attribute name='avpx_amount' />
                                <attribute name='avpx_type' />
                                <attribute name='avpx_extendedamount' />
                                <attribute name='avpx_description' />
                                <order attribute='avpx_name' descending='false' />
                                <filter type='and'>
                                  <condition attribute='statuscode' operator='eq' value='1' />
                                  <condition attribute='avpx_quote' operator='eq' uiname='Quo-01814' uitype='avpx_quote' value='{0}' />
                                </filter>
                              </entity>
                            </fetch>", recordId);

                // Retrieve records using the Fetch XML
                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXml));

                // Process the results
                if (entityCollection.Entities.Count > 0)
                {
                    tracingService.Trace("Records retrieved successfully. Count: {0}", entityCollection.Entities.Count);
                    foreach (Entity entity in entityCollection.Entities)
                    {
                        string avpxName = GetValueAsString(entity, "avpx_name", service, tracingService);
                        string avpxQuantity = RoundOffValue(entity, "avpx_quantity");
                        string avpxAmount = GetFormattedValueOrDefault(entity, "avpx_priceperunit", tracingService);
                        string avpxExtendedAmount = GetFormattedValueOrDefault(entity, "avpx_amount", tracingService);
                        string avpxType = GetValueAsString(entity, "avpx_type", service, tracingService);
                        string description = GetValueAsString(entity, "avpx_description", service, tracingService);

                        string displayName = (!string.IsNullOrEmpty(description) && description != "N/A") ? description : avpxName;
                        string[] recordData = new string[] { displayName, avpxQuantity, avpxAmount, avpxExtendedAmount };

                        if (avpxType == "Asset group") // Asset group
                        {
                            assetGroup.Add(recordData);
                        }
                        else if (avpxType == "Charge") // Charge
                        {
                            charges.Add(recordData);
                        }
                        data.Add(recordData);
                    }
                }
                else
                {
                    tracingService.Trace("No records retrieved.");
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("An error occurred: {0}", ex.Message);
            }

            tracingService.Trace("Combining List");
            // Combine lists
            var combinedResult = new List<string[]>();
            combinedResult.AddRange(assetGroup);
            combinedResult.Add(new string[] { "---- End of AssetGroup ----" });
            combinedResult.AddRange(charges);
            tracingService.Trace("Combine Completed");
            return combinedResult;
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
                            return dateTime.ToString("dd/MM/yyyy"); // Format as needed
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


