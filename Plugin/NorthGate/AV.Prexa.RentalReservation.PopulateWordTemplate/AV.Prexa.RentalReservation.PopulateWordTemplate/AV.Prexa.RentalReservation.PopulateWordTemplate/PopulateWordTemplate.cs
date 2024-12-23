﻿using DocumentFormat.OpenXml.Packaging;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Newtonsoft.Json;
using System;
using System.Activities;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Office.Drawing;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Globalization;

namespace AV.Prexa.RentalReservation.PopulateWordTemplate
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

        //public decimal damageWaiverPercentage = 0;
        //public decimal SubtotalRentals = 0;
        //public decimal SubtotalCharges = 0;
        //public decimal TotalTaxes = 0;
        //public decimal TotalAmount = 0;
        //public decimal environmentFees = 0;
        //public decimal damageWaiverAmount = 0;
        //public decimal amount = 0;

        public string SubtotalRentals = "0.00";
        public string SubtotalCharges = "0.00";
        public string TotalTaxes = "0.00";
        public string TotalAmount = "0.00";
        public decimal environmentFees = 0.00m;
        public string environmentFeeStr = "0.00";
        public string damageWaiverAmount = "0.00";
        public string amount = "0.00";



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

                    var placeholder = extractrentalReservationData(recordID, service, tracingService, context);

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

       

        private Dictionary<string, string> extractrentalReservationData(string rentalReservationId, IOrganizationService service, ITracingService tracingService, IWorkflowContext context)
        {
            var placeholders = new Dictionary<string, string>
{
    {"pxRentalId", ""},
    {"pxCustomerNo", "N/A"},
    {"pxEffectiveTo", ""},
    {"pxRentalOut", ""},
    {"pxRentalIn", ""},
    {"pxPONumber", ""},
    {"pxOrderedby", ""},
    {"pxsalesperson", "N/A"},
    {"pxrentalReservation Amount", ""},
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
    {"pxOffNumber", "N/A"},
    {"pxprintedOn", ""},
    {"pxCustoNae", ""},
    {"pxCellNo", "N/A"},
    {"pxcusAdd", ""},
    {"pxcustomerCit", ""},
    {"pxcusS", ""},
    {"pxcusP", ""},
    {"pxcusCo", ""},
    {"pxPricing",""},
    {"pxPaymentTerms",""},
    {"pxdateDue",""},
    {"pxPartsSubtotal","" },
    {"pxEnvironmentfee","" },
    {"pxSDate","" },
    {"pxEDate","" },
    {"pxRenPer",""},
    {"pxQName",""},
    {"pxQNo","" },
    {"pxOpp","" },
    {"pxSPho","" },
    {"pxSEmail","" },
    {"pxEmail","" }
};

            try
            {
                string rentalReservationFetch = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' >
    <entity name='avpx_rentalreservation' >
        <attribute name='avpx_rentalreservationid' />
        <attribute name='avpx_rentalorderid' />
        <attribute name='avpx_serviceaccount' />
        <attribute name='avpx_recordstatus' />
        <attribute name='avpx_rentalcontractstartdate' />
        <attribute name='avpx_nextinvoicedate' />
        <attribute name='avpx_jobsitestreetaddress' />
        <attribute name='avpx_jobsitecity' />
        <attribute name='avpx_jobsitestateorprovince' />
        <attribute name='avpx_jobsitepostalcode' />
        <attribute name='avpx_jobsitecountry' />
        <attribute name='avpx_estimatestartdate' />
        <attribute name='avpx_estimateenddate' />
        <attribute name='avpx_rentalperiod' />
        <attribute name='avpx_pricing' />
        <attribute name='avpx_quote' />
        <attribute name='avpx_opportunity' />
        <attribute name='avpx_damagewaiverpercentage' />
        <attribute name='ownerid' />
        <attribute name='createdon' />
        <order attribute='avpx_rentalorderid' descending='true' />
        <filter type='and' >
            <condition attribute='statuscode' operator='eq' value='1' />
            <condition attribute='avpx_rentalreservationid' operator='eq' uiname='1805' uitype='avpx_rentalreservation' value='{0}' />
        </filter>
        <link-entity name='account' from='accountid' to='avpx_serviceaccount' link-type='inner' alias='ac' >
            <attribute name='accountnumber' />
            <attribute name='name' />
            <attribute name='avpx_searchaddress1' />
            <attribute name='address1_line2' />
            <attribute name='address1_city' />
            <attribute name='address1_stateorprovince' />
            <attribute name='address1_postalcode' />
            <attribute name='address1_country' />
            <attribute name='accountnumber' />
            <attribute name='telephone1' />
            <attribute name='telephone3' />
            <attribute name='paymenttermscode' />
             <attribute name='emailaddress1' />
        </link-entity>
    </entity>
</fetch>", rentalReservationId);

                tracingService.Trace("Fetch Expression: " + rentalReservationFetch);
                EntityCollection rentalReservationEntityCollection;
                try
                {
                    rentalReservationEntityCollection = service.RetrieveMultiple(new FetchExpression(rentalReservationFetch));
                }
                catch (Exception ex)
                {
                    tracingService.Trace("Error retrieving rentalReservation data: " + ex.ToString());
                    throw new InvalidPluginExecutionException("Failed to retrieve rentalReservation data. " + ex.Message);
                }

                if (rentalReservationEntityCollection == null || rentalReservationEntityCollection.Entities == null || rentalReservationEntityCollection.Entities.Count == 0)
                {
                    tracingService.Trace("No rentalReservations found with the given ID.");
                    return placeholders;
                }

                Entity rentalReservation = rentalReservationEntityCollection.Entities[0];
                tracingService.Trace("rentalReservation Retrieved Successfully: " + rentalReservation);

                string GetValueOrDefault(string attribute, string defaultValue = "")
                {
                    try
                    {
                        if (rentalReservation.Contains(attribute))
                        {
                            var value = rentalReservation[attribute];
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
                        if (rentalReservation.FormattedValues.Contains(attribute))
                        {
                            var value = rentalReservation.FormattedValues[attribute];
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
                        return rentalReservation.Contains(attribute) ? (rentalReservation[attribute] as AliasedValue)?.Value.ToString() ?? defaultValue : defaultValue;
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace("Error retrieving aliased value for attribute " + attribute + ": " + ex.ToString());
                        return defaultValue;
                    }
                }

                int userTimeZone = GetUserTimezone(context.UserId, service, tracingService);
                tracingService.Trace($"User TimeZone: {userTimeZone}");

                string quoteName = "";
                string effectiveTo = "";
                string ownerTel = "";
                string ownerEmail = "";
                string jobTitle = "";
                DateTime effectiveLocal = DateTime.MinValue;
                if (rentalReservation.Contains("avpx_quote"))
                {
                    var quoteReference = (EntityReference)rentalReservation["avpx_quote"];
                    Guid quoteId = quoteReference.Id;

                    // Step 3: Retrieve the Contact Entity Using the GUID
                    var quote = service.Retrieve("avpx_quote", quoteId,
                        new ColumnSet("avpx_effectiveto", "avpx_name"));

                    // Step 4: Combine the Name and Display the Details
                    if (quote != null)
                    {
                        DateTime effectiveDateUtc = DateTime.Parse(GetValueOrDefault("avpx_effectiveto", DateTime.UtcNow.ToString("o")));
                         effectiveLocal = GetLocalTime(effectiveDateUtc, userTimeZone, service, tracingService);
                        quoteName = quote.GetAttributeValue<string>("avpx_name");
                        /* fullName = $"{contact.GetAttributeValue<string>("firstname")} " +
                                          $"{contact.GetAttributeValue<string>("lastname")}";
                         emailContact = contact.GetAttributeValue<string>("emailaddress1");
                         contactTel = contact.GetAttributeValue<string>("telephone1");
                         tracingService.Trace($"Full Name: {fullName}");
                         tracingService.Trace($"Email: {emailContact}");*/

                    }
                }
                else
                {
                    tracingService.Trace("No quote associated with this reservation.");
                }

                if (rentalReservation.Contains("avpx_opportunity"))
                {
                    var oppReference = (EntityReference)rentalReservation["avpx_opportunity"];
                    Guid opportunityId = oppReference.Id;

                    // Step 3: Retrieve the opportunitu Entity Using the GUID
                    var opportunity = service.Retrieve("avpx_potentialopportunity", opportunityId,
                        new ColumnSet("avpx_name"));

                    // Step 4: Combine the Name and Display the Details
                    if (opportunity != null)
                    {
                        jobTitle = opportunity.GetAttributeValue<string>("avpx_name");
                        tracingService.Trace($"jobTitle: {jobTitle}");

                    }
                }
                else
                {
                    tracingService.Trace("No opportunity associated with this reservation.");
                }

                tracingService.Trace("Quote Name"+ quoteName);

                if (rentalReservation.Contains("ownerid"))
                {
                    var ownerReference = (EntityReference)rentalReservation["ownerid"];
                    Guid ownerId = ownerReference.Id;

                    // Step 3: Retrieve the opportunitu Entity Using the GUID
                    var owner = service.Retrieve("systemuser", ownerId,
                        new ColumnSet("mobilephone", "internalemailaddress"));

                    // Step 4: Combine the Name and Display the Details
                    if (owner != null)
                    {
                        ownerEmail = owner.GetAttributeValue<string>("internalemailaddress");
                        ownerTel = owner.GetAttributeValue<string>("mobilephone");
                        tracingService.Trace($"jobTitle: {jobTitle}");

                    }
                }
                else
                {
                    tracingService.Trace("No opportunity associated with this reservation.");
                }

                // Get the user's timezone


                // Convert relevant dates to local time
                DateTime rentalOutDateUtc = DateTime.Parse(GetValueOrDefault("avpx_tentativerentalstartdate", DateTime.UtcNow.ToString("o")));
                DateTime rentalInDateUtc = DateTime.Parse(GetValueOrDefault("avpx_tentativerentalenddate", DateTime.UtcNow.ToString("o")));
                DateTime effectiveToDateUtc = DateTime.Parse(GetValueOrDefault("avpx_effectiveto", DateTime.UtcNow.ToString("o")));
                DateTime estimatedStartDate = DateTime.Parse(GetValueOrDefault("avpx_estimatestartdate", DateTime.UtcNow.ToString("o")));
                DateTime estimatedEndtDate = DateTime.Parse(GetValueOrDefault("avpx_estimateenddate", DateTime.UtcNow.ToString("o")));
                DateTime printedOnUtc = DateTime.Parse(DateTime.Now.ToString("o"));

                tracingService.Trace($"Rental Out Date (UTC): {rentalOutDateUtc}");
                tracingService.Trace($"Rental In Date (UTC): {rentalInDateUtc}");
                tracingService.Trace($"Effective To Date (UTC): {effectiveToDateUtc}");
                tracingService.Trace($"Printed On Date (UTC): {printedOnUtc}");

                DateTime rentalOutLocal = GetLocalTime(rentalOutDateUtc, userTimeZone, service, tracingService);
                DateTime rentalInLocal = GetLocalTime(rentalInDateUtc, userTimeZone, service, tracingService);
                DateTime effectiveToLocal = GetLocalTime(effectiveToDateUtc, userTimeZone, service, tracingService);
                DateTime printedOnToLocal = GetLocalTime(printedOnUtc, userTimeZone, service, tracingService);
                DateTime estimatedStartDateToLocal = GetLocalTime(estimatedStartDate, userTimeZone, service, tracingService);
                tracingService.Trace($"Rental Out Date (Local): {rentalOutLocal}");
                tracingService.Trace($"Rental In Date (Local): {rentalInLocal}");
                tracingService.Trace($"Effective To Date (Local): {effectiveToLocal}");
                tracingService.Trace($"Printed On Date (Local): {printedOnToLocal}");

                // Populate placeholders with data
                if (rentalReservation.Contains("avpx_pricing"))
                {
                    int pricingValue = ((OptionSetValue)rentalReservation["avpx_pricing"]).Value;
                    switch (pricingValue)
                    {
                        case 783090000:
                            placeholders["pxPricing"] = "Minimum Price";
                            break;
                        case 783090001:
                            placeholders["pxPricing"] = "Daily Price";
                            break;
                        case 783090002:
                            placeholders["pxPricing"] = "Weekly Price";
                            break;
                        case 783090003:
                            placeholders["pxPricing"] = "Monthly Price";
                            break;
                        default:
                            placeholders["pxPricing"] = "Unknown";
                            break;
                    }
                }

                if (rentalReservation.Contains("ac.paymenttermscode"))
                {
                    int paymentTermsValue = ((OptionSetValue)((AliasedValue)rentalReservation["ac.paymenttermscode"]).Value).Value;
                    switch (paymentTermsValue)
                    {
                        case 1:
                            placeholders["pxPaymentTerms"] = "Net 30";
                            break;
                        case 2:
                            placeholders["pxPaymentTerms"] = "2% 10, Net 30";
                            break;
                        case 3:
                            placeholders["pxPaymentTerms"] = "Net 45";
                            break;
                        case 4:
                            placeholders["pxPaymentTerms"] = "Net 60";
                            break;
                        case 783090000:
                            placeholders["pxPaymentTerms"] = "COD";
                            break;
                        default:
                            placeholders["pxPaymentTerms"] = "Unknown";
                            break;
                    }
                }

                placeholders["pxSDate"] = estimatedStartDateToLocal.ToString("dd/MM/yyyy");
                placeholders["pxEDate"] = estimatedEndtDate.ToString("dd/MM/yyyy");
                placeholders["pxRentalId"] = GetValueOrDefault("avpx_rentalorderid");
                placeholders["pxCustomerNo"] = GetAliasedValueOrDefault("ac.accountnumber", "N/A");
                placeholders["pxRenPer"] = GetFormattedValueOrDefault("avpx_rentalperiod");
                //placeholders["EffectiveTo"] = effectiveToLocal.ToString("MM/dd/yyyy");
                placeholders["pxRentalOut"] = rentalOutLocal.ToString("dd/MM/yyyy");
                placeholders["pxRentalIn"] = rentalInLocal.ToString("dd/MM/yyyy");
                placeholders["pxPONumber"] = GetValueOrDefault("avpx_ponumber");
                //placeholders["pxsalesperson"] = GetAliasedValueOrDefault("ownerid", "");
                placeholders["pxsalesperson"] = GetFormattedValueOrDefault("ownerid", "N/A");
                placeholders["pxrentalReservation Amount"] = GetFormattedValueOrDefault("avpx_totalamount");
                placeholders["pxrentalSubtotal"] = GetFormattedValueOrDefault("avpx_subtotalamount");
                placeholders["pxSMSubtotal"] = GetFormattedValueOrDefault("avpx_subtotalamount");
                //placeholders["DAmount"] = GetFormattedValueOrDefault("avpx_damagewaiveramount");
                placeholders["pxDWAmount"] = GetFormattedValueOrDefault("avpx_damagewaiveramount");
                placeholders["pxASubtotal"] = GetFormattedValueOrDefault("avpx_othercharges");
                placeholders["pxTAmount"] = GetFormattedValueOrDefault("avpx_totaltax");
                placeholders["pxTamt"] = GetFormattedValueOrDefault("avpx_totalamount");
                placeholders["pxRentalSubtotal"] = GetFormattedValueOrDefault("avpx_subtotalamount");
                placeholders["pxStreetAddress"] = GetValueOrDefault("avpx_jobsitestreetaddress");
                placeholders["pxCity"] = GetValueOrDefault("avpx_jobsitecity");
                //placeholders["State"] = GetValueOrDefault("avpx_stateorprovince");
                placeholders["pxState"] = string.IsNullOrEmpty(GetValueOrDefault("avpx_jobsitestateorprovince")) ? "" : GetValueOrDefault("avpx_jobsitestateorprovince") + ",";
               
                placeholders["pxPostalCode"] = GetValueOrDefault("avpx_jobsitepostalcode");
                placeholders["pxCountry"] = GetValueOrDefault("avpx_jobsitecountry");
                placeholders["pxOffNumber"] = GetAliasedValueOrDefault("ac.telephone1", "N/A");
                placeholders["pxCustoNae"] = GetAliasedValueOrDefault("ac.name");
                placeholders["pxCellNo"] = GetAliasedValueOrDefault("ac.telephone1", "N/A");
                placeholders["pxcusAdd"] = GetAliasedValueOrDefault("ac.address1_line2");
                placeholders["pxcustomerCit"] = GetAliasedValueOrDefault("ac.address1_city");
                //placeholders["cusS"] = GetAliasedValueOrDefault("ac.address1_stateorprovince");
                placeholders["pxcusS"] = String.IsNullOrEmpty(GetAliasedValueOrDefault("ac.address1_stateorprovince")) ? "" : GetAliasedValueOrDefault("ac.address1_stateorprovince") + ",";
                placeholders["pxcusP"] = GetAliasedValueOrDefault("ac.address1_postalcode");
                placeholders["pxcusCo"] = GetAliasedValueOrDefault("ac.address1_country");
                placeholders["pxprintedOn"] = printedOnToLocal.ToString("dd/MM/yyyy");
                placeholders["pxQName"] = effectiveLocal.ToString("dd/MM/yyyy");
                placeholders["pxQNo"] = quoteName;
                placeholders["pxOpp"] = jobTitle;
                placeholders["pxSEmail"] = ownerEmail;
                placeholders["pxSPho"] = ownerTel;
                placeholders["pxEmail"] = GetAliasedValueOrDefault("ac.emailaddress1");
                
                return placeholders;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error in extractrentalReservationData: " + ex.ToString());
                throw new InvalidPluginExecutionException("Failed to extract rentalReservation data. " + ex.Message);
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

        public async Task<string> PostBase64ToPowerAutomate(string base64String, string powerAutomateUrl, ITracingService tracingService)
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

        public  byte[] WordTemplate(byte[] fileData, byte[] imageData, Dictionary<string, string> placeholders, List<string[]> combinedData, ITracingService tracingService)
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
                       // string damageWaiverPercentageStr = damageWaiverPercentage;
                        string SubtotalRentalsStr = SubtotalRentals;
                        string SubtotalChargesStr = SubtotalCharges;
                        string TotalTaxesStr = TotalTaxes;
                        string TotalAmountStr = TotalAmount;
                        string environmentFeesStr = environmentFeeStr;
                        string damageWaiverAmountStr = damageWaiverAmount;
                        string amountStr = amount;

                        if (docText.Contains("pxDWAmount"))
                        {
                            tracingService.Trace("Found DWAmount" + damageWaiverAmountStr);
                            placeholders["pxDWAmount"] = damageWaiverAmountStr;
                           // docText.Replace("DWAmount", damageWaiverAmountStr);
                            
                        }
                        if (docText.Contains("pxSMSubtotal"))
                        {
                            
                            tracingService.Trace("Found SMSubtotal" + SubtotalRentalsStr);
                            placeholders["pxSMSubtotal"] = SubtotalRentalsStr;
                            //docText.Replace("SMSubtotal", SubtotalRentalsStr);
                        }
                        if (docText.Contains("pxPartsSubtotal"))
                        {
                            tracingService.Trace("Found PartsSubtotal" + SubtotalChargesStr);
                            placeholders["pxPartsSubtotal"] = SubtotalChargesStr;
                           // docText.Replace("PartsSubtotal", SubtotalChargesStr);
                            
                        }
                        if (docText.Contains("pxTAmount"))
                        {
                            tracingService.Trace("Found TAmount" + TotalTaxesStr);
                            placeholders["pxTAmount"] = TotalTaxesStr;
                            //docText.Replace("TAmount", TotalTaxesStr);
                            
                        }
                        if (docText.Contains("pxTamt"))
                        {
                            
                            tracingService.Trace("Found Tamt" + TotalAmountStr);
                            placeholders["pxTamt"] = TotalAmountStr;
                            //docText.Replace("Tamt", TotalAmountStr);
                        }
                        if (docText.Contains("pxEnvironmentfee"))
                        {
                            tracingService.Trace("Found Environmentfee" + environmentFeesStr);
                            placeholders["pxEnvironmentfee"] = environmentFeesStr;
                        }
                       


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

                // Use an index to get properties from existing cells if needed
                int cellIndex = 0;

                foreach (var cellData in rowData)
                {
                    tracingService.Trace($"Adding new cell with data: {cellData}");

                    // Create a new cell and paragraph
                    DocumentFormat.OpenXml.Wordprocessing.Paragraph newParagraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text(cellData)
                        )
                    );

                    // Copy alignment from an existing cell if available
                    if (table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().FirstOrDefault()?
                        .Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ElementAtOrDefault(cellIndex)?
                        .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is DocumentFormat.OpenXml.Wordprocessing.Paragraph existingParagraph)
                    {
                        var existingParagraphProperties = existingParagraph.ParagraphProperties;
                        if (existingParagraphProperties != null)
                        {
                            // Clone and add existing paragraph properties to the new paragraph
                            newParagraph.ParagraphProperties = (DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties)existingParagraphProperties.CloneNode(true);
                        }
                    }

                    DocumentFormat.OpenXml.Wordprocessing.TableCell newCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell(newParagraph);
                    newRow.Append(newCell);

                    cellIndex++;
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

        public List<string[]> RetrieveAndReturnRecords(IOrganizationService service, ITracingService tracingService, string rentalReservationId)
        {
            var data = new List<string[]>();
            var assetGroup = new List<string[]>();
            var charges = new List<string[]>();

            string currencySymbol = "";
            decimal localSubtotalRentals = 0;
            decimal localSubtotalCharges = 0;
            decimal localTotalTaxes = 0;
            decimal localTotalAmount = 0;
            decimal localEnvironmentFees = 0;
            decimal localDamageWaiverAmount = 0;
            decimal localEnvironmentFeeTotal = 0;
            decimal localAmount = 0;
            decimal damageWaiverPercentage = 0;
            bool isDamageWaiverApplicable = false;

            try
            {
                // Define the Fetch XML
                string fetchXml = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' >
    <entity name='avpx_rentalorderlineitem' >
        <attribute name='avpx_rentalorderlineitemid' />
        <attribute name='avpx_type' />
        <attribute name='avpx_assetgroup' />
        <attribute name='avpx_quantity' />
        <attribute name='avpx_priceperunit' />
        <attribute name='avpx_amount' />
        <attribute name='avpx_subtotalamount' />
        <attribute name='avpx_taxamount' />
        <attribute name='avpx_environmentfees' />
        <attribute name='avpx_extendedamount' />
        <attribute name='avpx_damagewaiverapplicable'/> 
         <attribute name='av_description'/>
        <attribute name='avpx_asset'/>
        <order attribute='createdon' descending='true' />
        <filter type='and' >
            <condition attribute='avpx_rentalreservation' operator='eq' uiname='1803' uitype='avpx_rentalreservation' value='{0}' />
            <condition attribute='statuscode' operator='eq' value='1' />
            <condition attribute='avpx_returnitem' operator='null' />
        </filter>
        <link-entity name='avpx_rentalreservation' from='avpx_rentalreservationid' to='avpx_rentalreservation' link-type='inner' alias='al' >
            <attribute name='avpx_damagewaiverpercentage' />
        </link-entity>
    </entity>
</fetch>", rentalReservationId);

                // Retrieve records using the Fetch XML
                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (entityCollection != null && entityCollection.Entities.Count > 0)
                {
                    // Retrieve the first entity
                    Entity entity = entityCollection.Entities[0];

                    // Check if the aliased attribute exists and is not null
                    if (entity.Contains("al.avpx_damagewaiverpercentage") && entity["al.avpx_damagewaiverpercentage"] is AliasedValue aliasedValue && aliasedValue.Value != null)
                    {
                        damageWaiverPercentage = (decimal)aliasedValue.Value;
                        tracingService.Trace($"Damage Waiver Percentage: {damageWaiverPercentage}");
                    }
                    else
                    {
                        tracingService.Trace("Damage Waiver Percentage is not found or is null.");
                    }
                }
                tracingService.Trace("Line Item Fetch" + fetchXml);
                tracingService.Trace("Fetch XML executed. Number of records retrieved: {0}", entityCollection.Entities.Count);

                // Process the results
                if (entityCollection.Entities.Count > 0)
                {
                    foreach (Entity entity in entityCollection.Entities)
                    {
                        tracingService.Trace("Processing Entity ID: {0}", entity.Id);

                        // Retrieve values
                        string avpxType = GetValueAsString(entity, "avpx_type", service, tracingService);
                        string avpxAssetGroupName = "";
                        string assetName = "N/A";
                        string make = "N/A";
                        string serialNo = "N/A";
                        string unitNo = "N/A";
                        string chargename = "N/A";
                        //string assetID = "N/A";
                        Guid assetGroupId = entity.GetAttributeValue<EntityReference>("avpx_assetgroup")?.Id ?? Guid.Empty;
                        Guid assetId = entity.GetAttributeValue<EntityReference>("avpx_asset")?.Id ?? Guid.Empty; // Retrieve the asset ID
                        if (assetGroupId != Guid.Empty)
                        {
                            Entity partEntity = service.Retrieve("avpx_device", assetGroupId, new ColumnSet("avpx_name"));
                            avpxAssetGroupName = partEntity.Contains("avpx_name") ? partEntity["avpx_name"].ToString() : string.Empty;
                            tracingService.Trace("Part Name: {0}", avpxAssetGroupName);
                        }
                        else
                        {
                            tracingService.Trace("No asset group information available.");
                        }
                        if (assetGroupId != Guid.Empty)
                        {
                            if (avpxType == "Charge")
                            {
                                Entity chargeEntity = service.Retrieve("avpx_device", assetGroupId, new ColumnSet("avpx_chargename"));
                                chargename = chargeEntity.Contains("avpx_chargename") ? chargeEntity["avpx_chargename"].ToString() : string.Empty;
                                tracingService.Trace("cHARGE NAME" + chargename);
                            }
                        }

                            if (assetId != Guid.Empty) 
                        {
                            Entity assetEntity = service.Retrieve("avpx_asset", assetId, new ColumnSet("avpx_name", "avpx_makeyear", "av_unit", "avpx_serialnumbervin", "avpx_recordid"));
                             assetName = assetEntity.Contains("avpx_name") ? assetEntity["avpx_name"].ToString() : "N/A";
                            tracingService.Trace("Asset Name: {0}", assetName);
                             make = assetEntity.Contains("avpx_makeyear") ? assetEntity["avpx_makeyear"].ToString() : "N/A";
                             unitNo = assetEntity.Contains("avpx_recordid") ? assetEntity["avpx_recordid"].ToString() : "N/A";
                             serialNo = assetEntity.Contains("avpx_serialnumbervin") ? assetEntity["avpx_serialnumbervin"].ToString() : "N/A";
                            //assetID = assetEntity.Contains("avpx_recordid") ? assetEntity["avpx_recordid"].ToString() : "N/A";
                            // Concatenate asset group name with asset name if asset exists
                            if (!string.IsNullOrEmpty(assetName))
                            {
                                avpxAssetGroupName = $"{avpxAssetGroupName} - {assetName}";
                            }
                        }

                        tracingService.Trace("Final Asset Group Name with Asset: {0}", avpxAssetGroupName);

                        string description = "";
                        description = GetValueAsString(entity, "av_description", service, tracingService);

                        // Retrieve and trace attributes
                        string avpxQuantity = RoundOffValue(entity, "avpx_quantity");
                        tracingService.Trace("Quantity (formatted): {0}", avpxQuantity);

                        string avpxPricePerUnit = GetFormattedValueOrDefault(entity, "avpx_priceperunit", tracingService, ref currencySymbol);
                        tracingService.Trace("Price Per Unit (formatted): {0}", avpxPricePerUnit);

                        string avpxTaxAmount = GetFormattedValueOrDefault(entity, "avpx_taxamount", tracingService, ref currencySymbol);
                        tracingService.Trace("Tax Amount (formatted): {0}", avpxTaxAmount);

                        string avpxAmount = GetFormattedValueOrDefault(entity, "avpx_amount", tracingService, ref currencySymbol);
                        tracingService.Trace("Amount (formatted): {0}", avpxAmount);

                        string avpxExtendedAmount = GetFormattedValueOrDefault(entity, "avpx_extendedamount", tracingService, ref currencySymbol);
                        tracingService.Trace("Extended Amount (formatted): {0}", avpxExtendedAmount);

                        string avpxEnvironmentFees = GetFormattedValueOrDefault(entity, "avpx_environmentfees", tracingService, ref currencySymbol);
                        tracingService.Trace("Environment Fees (formatted): {0}", avpxEnvironmentFees);

                        // Retrieve and store the value of avpx_damagewaiverapplicable
                        isDamageWaiverApplicable = entity.GetAttributeValue<bool?>("avpx_damagewaiverapplicable") ?? false;

                        // Update local variables
                        localAmount = GetRawValue(entity, "avpx_amount");
                        localEnvironmentFees = GetRawValue(entity, "avpx_environmentfees");
                        tracingService.Trace("Raw Amount: {0}, Raw Environment Fees: {1}", localAmount, localEnvironmentFees);
                        tracingService.Trace("Damage Waiver Applicable: " + (isDamageWaiverApplicable ? "Yes" : "No"));

                        if (isDamageWaiverApplicable)
                        {
                            localDamageWaiverAmount = (localAmount * damageWaiverPercentage) / 100;
                            tracingService.Trace("Calculated Damage Waiver Amount: {0}", localDamageWaiverAmount);
                        }
                        else
                        {
                            localDamageWaiverAmount = 0;
                            tracingService.Trace("Damage Waiver is not applicable, skipping calculation.");
                        }

                        localEnvironmentFeeTotal += localEnvironmentFees;
                        tracingService.Trace("Updated Environment Fee Total: {0}", localEnvironmentFeeTotal);

                        if (avpxType == "Asset group") // Asset group
                        {
                            string[] recordData = new string[] { make, description, unitNo, serialNo };
                            assetGroup.Add(recordData);

                            // Calculate Subtotal Rentals
                            localSubtotalRentals += localAmount + localEnvironmentFees + localDamageWaiverAmount;
                            tracingService.Trace("Updated Subtotal Rentals: {0}", localSubtotalRentals);
                        }
                        else if (avpxType == "Charge") // Charge
                        {
                            string[] recordData = new string[] { description, avpxQuantity, avpxPricePerUnit, avpxAmount };
                            charges.Add(recordData);

                            // Calculate Subtotal Charges
                            localSubtotalCharges += GetRawValue(entity, "avpx_amount");
                            tracingService.Trace("Updated Subtotal Charges: {0}", localSubtotalCharges);
                        }

                        // Calculate Total Taxes
                        localTotalTaxes += GetRawValue(entity, "avpx_taxamount");
                        tracingService.Trace("Updated Total Taxes: {0}", localTotalTaxes);

                        data.Add(new string[] { avpxAssetGroupName, avpxQuantity, avpxTaxAmount, avpxPricePerUnit, avpxExtendedAmount });
                    }

                    // Calculate Total Amount
                    localTotalAmount = localSubtotalRentals + localSubtotalCharges + localTotalTaxes;
                    tracingService.Trace("Subtotal Rentals: {0}, Subtotal Charges: {1}, Total Taxes: {2}, Calculated Total Amount: {3}",
                    localSubtotalRentals, localSubtotalCharges, localTotalTaxes, localTotalAmount);
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

            // Format the totals with the currency symbol
            SubtotalRentals = FormatWithCurrency(localSubtotalRentals, currencySymbol);
            SubtotalCharges = FormatWithCurrency(localSubtotalCharges, currencySymbol);
            TotalTaxes = FormatWithCurrency(localTotalTaxes, currencySymbol);
            TotalAmount = FormatWithCurrency(localTotalAmount, currencySymbol);
            environmentFeeStr = FormatWithCurrency(localEnvironmentFeeTotal, currencySymbol); // Assign the local environment fees to the global variable
            damageWaiverAmount = FormatWithCurrency(localDamageWaiverAmount, currencySymbol);
            tracingService.Trace("Final Results - Subtotal Rentals: {0}, Subtotal Charges: {1}, Total Taxes: {2}, Total Amount: {3}, Environment Fees: {4}",
                SubtotalRentals, SubtotalCharges, TotalTaxes, TotalAmount, environmentFeeStr);

            return combinedResult;
        }




        private string GetFormattedValueOrDefault(Entity entity, string attributeName, ITracingService tracingService, ref string currencySymbol, string defaultValue = "N/A")
        {
            try
            {
                if (entity.FormattedValues.Contains(attributeName))
                {
                    var value = entity.FormattedValues[attributeName];
                    if (currencySymbol == "")
                    {
                        // Extract the currency symbol from the formatted value
                        currencySymbol = new string(value.Where(c => !char.IsDigit(c) && c != '.' && c != ',').ToArray());
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

        private decimal GetRawValue(Entity entity, string attributeName)
        {
            return entity.Contains(attributeName) ? entity.GetAttributeValue<Money>(attributeName).Value : 0;
        }

        private string FormatWithCurrency(decimal value, string currencySymbol)
        {
            return $"{currencySymbol}{value:N2}";
        }


        private string RoundOffValue(Entity entity, string attributeName)
        {
            if (entity.Contains(attributeName))
            {
                var value = entity[attributeName];
                if (value != null)
                {
                    if (decimal.TryParse(value.ToString(), out decimal decimalValue))
                    {
                        return decimalValue.ToString("F2");
                    }
                }
            }
            return "0.00";
        }

        private decimal ParseDecimal(string value)
        {
            return decimal.TryParse(value, NumberStyles.Currency, CultureInfo.InvariantCulture, out decimal result) ? result : 0m;
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
