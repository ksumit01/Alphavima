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
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Office.Drawing;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Text.RegularExpressions;



namespace AV.Prexa.Invoice.PopulateWordTemplate
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
        public string totalQtyShipped = "";
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
                string base64Pdf = RetrieveEntityColumnData(service, new Guid(docTemplateID), recordId, tracingService,context).Result;

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

                    var placeholder = extractinvoiceData(recordID, service, tracingService, context);

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

        /*private Dictionary<string, string> extractinvoiceData(string invoiceId, IOrganizationService service, ITracingService tracingService)
        {
            var placeholders = new Dictionary<string, string>
    {
        {"InvoiceNumber", ""},
        {"CustomerNo", "N/A"},
        {"EffectiveTo", ""},
        {"RentalOut", ""},
        {"RentalIn", ""},
        {"PONumber", ""},
        {"Orderedby", ""},
        {"salesperson", "N/A"},
        {"ReservedBy","N/A" },
        {"InvoiceAmount", ""},
        {"SMSubtotal", ""},
        {"DAmount", ""},
        {"DWAmount", ""},
        {"ASubtotal", ""},
        {"TAmount", ""},
        {"Tamt", ""},
        {"PartsSubtotal","" },
        {"RentalSubtotal", ""},
        {"EnvironmentFee","" },
        {"StreetAddress", ""},
        {"City", ""},
        {"State", ""},
        {"PostalCode", ""},
        {"Country", ""},
        {"OfficeNumber", "N/A"},
        {"printedOn", DateTime.Now.ToString("MM/dd/yyyy")},
        {"CustomerName", ""},
        {"CellNo", "N/A"},
        {"cusAdd", ""},
        {"customerCit", ""},
        {"cusS", ""},
        {"cusP", ""},
        {"cusCo", ""},
        {"BillFrom","" },
        {"BillTo","" }
    };

            try
            {
                string invoiceFetch = string.Format(@"
                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='avpx_invoice'>
                        <attribute name='avpx_name' />
                        <attribute name='createdon' />
                        <attribute name='avpx_invoicedate' />
                        <attribute name='avpx_billfromdate' />
                        <attribute name='avpx_billto' />
                        <attribute name='avpx_unitsprice' />
                        <attribute name='avpx_discountamount' />
                        <attribute name='avpx_totalrentalamountnew' />
                        <attribute name='avpx_damagewaiverpercentage' />
                        <attribute name='avpx_damagewaiveramountnew' />
                        <attribute name='avpx_additionalcharges' />
                        <attribute name='avpx_environmentfee' />
                        <attribute name='avpx_subtotalamount' />
                        <attribute name='avpx_totaltax' />
                        <attribute name='avpx_totalamount' />
                        <attribute name='avpx_totalamountreceived' />
                        <attribute name='avpx_invoicetype' />
                        <attribute name='avpx_returnid' />
                        <attribute name='avpx_customerscontact' />
                        <attribute name='avpx_customer' />
                        <attribute name='avpx_rentalcontract' />
                        <attribute name='avpx_invoiceid' />
                        <order attribute='createdon' descending='true' />
                        <filter type='and'>
                            <condition attribute='statuscode' operator='eq' value='1' />
                            <condition attribute='avpx_invoiceid' operator='eq' uitype='avpx_invoice' value='{0}' />
                        </filter>
                        <link-entity name='account' from='accountid' to='avpx_customer' link-type='inner' alias='ac'>
                            <attribute name='accountnumber' />
                            <attribute name='avpx_searchaddress1' />
                            <attribute name='address1_line2' />
                            <attribute name='address1_city' />
                            <attribute name='address1_stateorprovince' />
                            <attribute name='address1_postalcode' />
                            <attribute name='address1_country' />
                            <attribute name='accountnumber' />
                            <attribute name='telephone1' />
                            <attribute name='telephone3' />
                        </link-entity>
                        <link-entity name='avpx_rentalreservation' from='avpx_rentalreservationid' to='avpx_rentalcontract' link-type='inner' alias='ar'>
                            <attribute name='avpx_estimatestartdate' />
                            <attribute name='avpx_jobsitestreetaddress' />
                            <attribute name='avpx_jobsitecity' />
                            <attribute name='avpx_jobsitestateorprovince' />
                            <attribute name='avpx_jobsitecountry' />
                            <attribute name='avpx_estimateenddate' />
                            <Attribute name='avpx_rentalorderid' />
                            <attribute name='avpx_billtodate' />
                            <attribute name='avpx_ponumber' />
                        </link-entity>
                    </entity>
                </fetch>", invoiceId);



                tracingService.Trace("Fetch Expression" + invoiceFetch);
                EntityCollection invoiceEntityCollection;
                try
                {
                    invoiceEntityCollection = service.RetrieveMultiple(new FetchExpression(invoiceFetch));
                }
                catch (Exception ex)
                {
                    tracingService.Trace("Error retrieving invoice data: " + ex.ToString());
                    throw new InvalidPluginExecutionException("Failed to retrieve invoice data. " + ex.Message);
                }

                if (invoiceEntityCollection == null || invoiceEntityCollection.Entities == null || invoiceEntityCollection.Entities.Count == 0)
                {
                    tracingService.Trace("No invoices found with the given ID.");
                    return placeholders;
                }

                Entity invoice = invoiceEntityCollection.Entities[0];
                tracingService.Trace("invoice Retrieved Successfully: " + invoice);

                string GetValueOrDefault(string attribute, string defaultValue = "")
                {
                    try
                    {
                        if (invoice.Contains(attribute))
                        {
                            var value = invoice[attribute];
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
                        if (invoice.FormattedValues.Contains(attribute))
                        {
                            var value = invoice.FormattedValues[attribute];
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
                        return invoice.Contains(attribute) ? (invoice[attribute] as AliasedValue)?.Value.ToString() ?? defaultValue : defaultValue;
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace("Error retrieving aliased value for attribute " + attribute + ": " + ex.ToString());
                        return defaultValue;
                    }
                }

                placeholders["invoicenumber"] = GetValueOrDefault("avpx_name");
                placeholders["CustomerNo"] = GetAliasedValueOrDefault("ac.accountnumber", "N/A");
                placeholders["EffectiveTo"] = GetValueOrDefault("avpx_effectiveto");
                placeholders["RentalOut"] = GetValueOrDefault("avpx_tentativerentalstartdate");
                placeholders["RentalIn"] = GetValueOrDefault("avpx_tentativerentalenddate");
                placeholders["PONumber"] = GetValueOrDefault("avpx_ponumber");
                placeholders["Orderedby"] = GetValueOrDefault("avpx_customerscontact");
                placeholders["salesperson"] = GetValueOrDefault("ownerid", "N/A");
                placeholders["invoice Amount"] = GetFormattedValueOrDefault("avpx_totalamount");
                placeholders["rentalSubtotal"] = GetFormattedValueOrDefault("avpx_subtotalamount");
                placeholders["SMSubtotal"] = GetFormattedValueOrDefault("avpx_othercharges");
                placeholders["DAmount"] = GetFormattedValueOrDefault("avpx_totalmanualdiscount");
                placeholders["DWAmount"] = GetFormattedValueOrDefault("avpx_damagewaiveramount");
                placeholders["ASubtotal"] = GetFormattedValueOrDefault("avpx_subtotalamount");
                placeholders["TAmount"] = GetFormattedValueOrDefault("avpx_totaltax");
                placeholders["Tamt"] = GetFormattedValueOrDefault("avpx_totalamount");
                placeholders["RentalSubtotal"] = GetFormattedValueOrDefault("avpx_detailedamount");
                placeholders["StreetAddress"] = GetValueOrDefault("avpx_streetaddress");
                placeholders["City"] = GetValueOrDefault("avpx_city");
                placeholders["State"] = string.IsNullOrEmpty(GetValueOrDefault("avpx_stateorprovince")) ? "" : GetValueOrDefault("avpx_stateorprovince") + ",";
                placeholders["PostalCode"] = GetValueOrDefault("avpx_postalcode");
                placeholders["Country"] = GetValueOrDefault("avpx_country");
                placeholders["OfficeNumber"] = GetAliasedValueOrDefault("ac.telephone1", "N/A");
                placeholders["CustomerName"] = GetValueOrDefault("avpx_account");
                placeholders["CellNo"] = GetAliasedValueOrDefault("ac.telephone3", "N/A");
                placeholders["cusAdd"] = GetAliasedValueOrDefault("ac.address1_line2");
                placeholders["customerCit"] = GetAliasedValueOrDefault("ac.address1_city");
                placeholders["cusS"] = string.IsNullOrEmpty(GetAliasedValueOrDefault("ac.address1_stateorprovince")) ? "" : GetAliasedValueOrDefault("ac.address1_stateorprovince") + ",";
                placeholders["cusP"] = GetAliasedValueOrDefault("ac.address1_postalcode");
                placeholders["cusCo"] = GetAliasedValueOrDefault("ac.address1_country");

                tracingService.Trace("Placeholders populated successfully.");
            }
            catch (Exception ex)
            {
                tracingService.Trace("An error occurred: " + ex.ToString());
            }

            return placeholders;
        }*/

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

        public string GetinvoiceFetchXml(Guid invoiceId, int orderType)
        {
            string dynamicLinkEntity = "";
            string extraAttributes = ""; // For additional fields specific to Part Sale

            switch (orderType)
            {
                case 783090000: // Rental
                    dynamicLinkEntity = @"
    <link-entity name='avpx_rentalreservation' from='avpx_rentalreservationid' to='avpx_rentalcontract' link-type='inner' alias='ar'>
        <attribute name='avpx_estimatestartdate' />
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
        <attribute name='avpx_estimateenddate' />
        <attribute name='avpx_rentalorderid' />
        <attribute name='avpx_billtodate' />
        <attribute name='avpx_ponumber' />
        <attribute name='avpx_rentalorderid' />
    </link-entity>";
                    break;
                case 783090001: // Sales
                    dynamicLinkEntity = @"
    <link-entity name='avpx_salesorders' from='avpx_salesordersid' to='avpx_salesorders' link-type='inner' alias='ab'>
        <attribute name='avpx_orderplacementdate' />
        <attribute name='avpx_streetaddressda' />
        <attribute name='avpx_cityda' />
        <attribute name='avpx_stateorprovinceda' />
        <attribute name='avpx_postalcodeda' />
        <attribute name='avpx_countryda' />
    </link-entity>";
                    break;
                case 783090003: // Part Sale
                    dynamicLinkEntity = "";

                    // Adding extra attributes specific to Part Sale
                    extraAttributes = @"
        <attribute name='av_streetaddress' />
        <attribute name='av_city' />
        <attribute name='av_stateprovince' />
         <attribute name='av_postalcodezip' />
        <attribute name='av_country' />
        <attribute name='av_shippedvia' />
        <attribute name='av_invoicedate' />
        <attribute name='av_shippingdate' />
        <attribute name='av_partsamount' />
        <attribute name='av_discount' />
        <attribute name='av_hst13' />
        <attribute name='av_totalamount' />
        <attribute name='av_ponumber' />
        <attribute name='av_id' />
        <attribute name='av_ponumber' />
        <attribute name='ownerid' />
        <attribute name='av_terms' />";

                    break;

                case 783090004: //Float
                    dynamicLinkEntity = "";

                    extraAttributes = @"
        <attribute name='av_streetaddressst' />
        <attribute name='av_cityst' />
        <attribute name='av_stateprovincest' />
         <attribute name='av_postalcodezipst' />
        <attribute name='av_countryst' />
        <attribute name='av_streetaddress' />
        <attribute name='av_city' />
        <attribute name='av_stateprovince' />
        <attribute name='av_postalcodezip' />
        <attribute name='av_country' />
        <attribute name='av_ponumber' />
        <attribute name='av_unitdescription' />
        <attribute name='av_transportdate' />
        <attribute name='av_floatchargeforfloattype' />
        <attribute name='av_adminfeeforfloattype' />
        <attribute name='av_subtotal' />
        <attribute name='av_hst' />
        <attribute name='av_total' />";
                    break;
                default:
                    throw new ArgumentException("Invalid order type");
            }

            string invoiceFetch = string.Format(@"
<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
    <entity name='avpx_invoice'>
        <attribute name='avpx_name' />
        <attribute name='createdon' />
        <attribute name='avpx_invoicedate' />
        <attribute name='avpx_billfromdate' />
        <attribute name='avpx_billto' />
        <attribute name='avpx_unitsprice' />
        <attribute name='avpx_discountamount' />
        <attribute name='avpx_totalrentalamountnew' />
        <attribute name='avpx_salesamount' />
        <attribute name='avpx_damagewaiveramountnew' />
        <attribute name='avpx_additionalcharges' />
        <attribute name='avpx_environmentfee' />
        <attribute name='avpx_subtotalamount' />
        <attribute name='avpx_totaltax' />
        <attribute name='avpx_totalamount' />
        <attribute name='avpx_totalamountreceived' />
        <attribute name='avpx_invoicetype' />
        <attribute name='avpx_returnid' />
        <attribute name='avpx_customerscontact' />
        <attribute name='avpx_customer' />
        <attribute name='avpx_rentalcontract' />
        <attribute name='avpx_invoiceid' />
        <attribute name='avpx_invoiceno' />
        <attribute name='av_fuelsurcharge' />
        <attribute name='av_floatcharge' />
        <attribute name='av_adminfee' />
        <attribute name='av_missingkeycharge' />
        <attribute name='av_missingmanualdocumentscharge' />
        <attribute name='av_deposittaken' />
        <attribute name='av_delivery' />
        <attribute name='av_pickup' />
       <attribute name='avpx_invoicedays' />
        <attribute name='av_transportcharge' />
        <attribute name='av_idnew' />
        <attribute name='avpx_notes' />
        <attribute name='avpx_jsl_streetaddress' />
        <attribute name='avpx_jsl_city' />
        <attribute name='avpx_jsl_stateorprovince' />
        <attribute name='avpx_jsl_postalcode' />
        <attribute name='avpx_jsl_country' />
        
        {0} <!-- Insert extra attributes here -->
        <order attribute='createdon' descending='true' />
        <filter type='and'>
            <condition attribute='statuscode' operator='eq' value='1' />
            <condition attribute='avpx_invoiceid' operator='eq' uitype='avpx_invoice' value='{1}' />
        </filter>
        <link-entity name='account' from='accountid' to='avpx_customer' link-type='outer' alias='ac'>
            <attribute name='name' />
            <attribute name='accountnumber' />
            <attribute name='avpx_searchaddress1' />
            <attribute name='address1_line2' />
            <attribute name='address1_city' />
            <attribute name='address1_stateorprovince' />
            <attribute name='address1_postalcode' />
            <attribute name='address1_country' />
            <attribute name='telephone1' />
            <attribute name='telephone3' />
            <attribute name='emailaddress1' />
            <attribute name='paymenttermscode' />
        </link-entity>
        {2} <!-- Insert dynamic link entity here -->
    </entity>
</fetch>", extraAttributes, invoiceId, dynamicLinkEntity);

            return invoiceFetch;
        }




        private Dictionary<string, string> extractinvoiceData(string invoiceId, IOrganizationService service, ITracingService tracingService, IWorkflowContext context)
        {
            var placeholders = new Dictionary<string, string>
    {
        {"pxInvoiceNumber", ""},
        {"pxCustomerNo", "N/A"},
        {"pxEffectiveTo", ""},
        {"pxshippingDate","" },
        {"pxTerms","" },
        {"pxRentalOut", ""},
        {"pxRentalIn", ""},
        {"pxPONumber", ""},
        {"pxOrderedby", ""},
        {"pxsalesperson", "N/A"},
        {"pxReservedBy","N/A" },
        {"pxInvoiceAmount", ""},
        {"pxSMSubtotal", ""},
        {"pxDAmount", ""},
        {"pxDWAmount", ""},
        {"pxASubtotal", ""},
        {"pxTAmount", ""},
        {"pxTamt", ""},
        {"pxPartsSubtotal","" },
        {"pxRentalSubtotal", ""},
        {"pxEnvironmentFee","" },
        {"pxStreetAddress", ""},
        {"pxCity", ""},
        {"pxState", ""},
        {"pxPostalCode", ""},
        {"pxCountry", ""},
        {"pxOffNumber", "N/A"},
        {"pxprintedOn", "" },
        {"pxCustomerName", ""},
        {"pxCellNo", "N/A"},
        {"pxcusAdd", ""},
        {"pxcustomerCit", ""},
        {"pxcusS", ""},
        {"pxcusP", ""},
        {"pxcusCo", ""},
        {"pxBillFrom","" },
        {"pxFuelAmt","" },
        {"pxFloatAmt","" },
        {"pxAdminAmt","" },
        {"pxMisAmt","" },
        {"pxKeyAmt","" },
        {"pxDepAmt","" },
        {"pxVia","" },
        {"pxAcntTerms","" },
        {"pxDueDate","" },
        {"pxTQS","" },
        {"pxAccName","" },
        {"pxDescription","" },
        {"pxPickUp","" },
        {"pxDelivery","" },
        {"pxRental","" },
        {"pxDays","" },
        {"pxFrStAdd",""},
        {"pxFrCity","" },
        {"pxFrState","" },
        {"pxFrPC","" },
        {"pxFrCon","" },
        {"pxEmail","" },
        {"pxNotes","" },
        {"pxTransCh","" }
       
    };

            try
            {
                // Initial fetch to get the order type
                string initialFetch = string.Format(@"
        <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='avpx_invoice'>
                <attribute name='avpx_ordertype' />
                <filter type='and'>
                    <condition attribute='avpx_invoiceid' operator='eq' uitype='avpx_invoice' value='{0}' />
                </filter>
            </entity>
        </fetch>", invoiceId);

                tracingService.Trace("Initial Fetch Expression: " + initialFetch);
                EntityCollection initialEntityCollection;
                try
                {
                    initialEntityCollection = service.RetrieveMultiple(new FetchExpression(initialFetch));
                }
                catch (Exception ex)
                {
                    tracingService.Trace("Error retrieving initial invoice data: " + ex.ToString());
                    throw new InvalidPluginExecutionException("Failed to retrieve initial invoice data. " + ex.Message);
                }

                if (initialEntityCollection == null || initialEntityCollection.Entities == null || initialEntityCollection.Entities.Count == 0)
                {
                    tracingService.Trace("No invoices found with the given ID.");
                    return placeholders;
                }

                Entity initialinvoice = initialEntityCollection.Entities[0];
                int orderType = initialinvoice.Contains("avpx_ordertype") ? ((OptionSetValue)initialinvoice["avpx_ordertype"]).Value : 0;
                tracingService.Trace("OrderType:" +  orderType);
                // Generate the dynamic FetchXML based on the order type
                string dynamicFetchXml = GetinvoiceFetchXml(new Guid(invoiceId), orderType);

                tracingService.Trace("Dynamic Fetch Expression: " + dynamicFetchXml);
                EntityCollection invoiceEntityCollection;
                try
                {
                    invoiceEntityCollection = service.RetrieveMultiple(new FetchExpression(dynamicFetchXml));
                }
                catch (Exception ex)
                {
                    tracingService.Trace("Error retrieving invoice data: " + ex.ToString());
                    throw new InvalidPluginExecutionException("Failed to retrieve invoice data. " + ex.Message);
                }

                if (invoiceEntityCollection == null || invoiceEntityCollection.Entities == null || invoiceEntityCollection.Entities.Count == 0)
                {
                    tracingService.Trace("No invoices found with the given ID.");
                    return placeholders;
                }

                Entity invoice = invoiceEntityCollection.Entities[0];
                tracingService.Trace("invoice Retrieved Successfully: " + invoice);


                string fullName = "";
                string emailContact = "";
                string contactTel = "";
                if (invoice.Contains("avpx_customerscontact"))
                {
                    var contactReference = (EntityReference)invoice["avpx_customerscontact"];
                    Guid contactId = contactReference.Id;

                    // Step 3: Retrieve the Contact Entity Using the GUID
                    var contact = service.Retrieve("contact", contactId,
                        new ColumnSet("firstname", "lastname", "emailaddress1", "telephone1"));

                    // Step 4: Combine the Name and Display the Details
                    if (contact != null)
                    {
                        fullName = $"{contact.GetAttributeValue<string>("firstname")} " +
                                         $"{contact.GetAttributeValue<string>("lastname")}";
                        emailContact = contact.GetAttributeValue<string>("emailaddress1");
                        contactTel = contact.GetAttributeValue<string>("telephone1");
                        tracingService.Trace($"Full Name: {fullName}");
                        tracingService.Trace($"Email: {emailContact}");

                    }
                }
                else
                {
                    tracingService.Trace("No contact associated with this invoice.");
                }
                // Helper methods to get values from the entity
                /*string GetValueOrDefault(string attribute, string defaultValue = "")
                {
                    try
                    {
                        if (invoice.Contains(attribute))
                        {
                            var value = invoice[attribute];
                            if (value is EntityReference entityRef)
                            {
                                return entityRef.Name ?? defaultValue;
                            }
                            if (value is Money moneyValue)
                            {
                                return moneyValue.Value.ToString("N2"); // Format money values with two decimal places
                            }
                            *//*if (value is DateTime dateValue)
                            {
                                return dateValue.ToString("dd/MM/yyyy"); // Format dates as dd/MM/yyyy
                            }*//*
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


*/                /*string GetFormattedValueOrDefault(string attribute, string defaultValue = "")
                {
                    try
                    {
                        if (invoice.FormattedValues.Contains(attribute))
                        {
                            var value = invoice.FormattedValues[attribute];
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
                }*/
                

                string GetValueOrDefault(string attribute, string defaultValue = "")
                {
                    try
                    {
                        if (invoice.Contains(attribute))
                        {
                            var value = invoice[attribute];

                            if (value is EntityReference entityRef)
                            {
                                return entityRef.Name ?? defaultValue;
                            }

                            if (value is Money moneyValue)
                            {
                                // Format the money value as a string and apply regex to remove any country prefix before the dollar sign
                                var moneyString = moneyValue.Value.ToString("C2"); // Convert to currency format
                                moneyString = Regex.Replace(moneyString, @"^[A-Z]{2}\$", "$"); // Replace prefix like "CA$" with "$"

                                return moneyString;
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
                        if (invoice.FormattedValues.Contains(attribute))
                        {
                            var value = invoice.FormattedValues[attribute];

                            // Use regular expression to remove any country prefix before the dollar sign
                            value = Regex.Replace(value, @"^[A-Z]{2}\$", "$"); // Replace prefix like "CA$" with "$"

                            if (decimal.TryParse(value.Replace("$", "").Trim(), out decimal decimalValue))
                            {
                                return "$" + decimalValue.ToString("N2"); // Format with dollar sign and two decimal places
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

                /*string GetAliasedValueOrDefault(string attribute, string defaultValue = "")
                {
                    try
                    {
                        if (invoice.Contains(attribute) && invoice[attribute] is AliasedValue aliasedValue)
                        {
                            if (aliasedValue.Value is Money moneyValue)
                            {
                                return moneyValue.Value.ToString("N2"); // Format money values with two decimal places
                            }
                            if (aliasedValue.Value is DateTime dateValue)
                            {
                                tracingService.Trace(attribute + ": Raw DateTime: " + dateValue.ToString());
                                tracingService.Trace(attribute + ": Formatted DateTime (ISO 8601): " + dateValue.ToString("o"));
                                // Optionally use specific format or custom parsing if needed
                                return dateValue.ToString("o"); // You can adjust this if needed
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
                }*/

                string GetAliasedValueOrDefault(string attribute, string defaultValue = "")
                {
                    try
                    {
                        if (invoice.Contains(attribute) && invoice[attribute] is AliasedValue aliasedValue)
                        {
                            if (aliasedValue.Value is Money moneyValue)
                            {
                                // Format the money value as a string and apply regex to remove any country prefix before the dollar sign
                                var moneyString = moneyValue.Value.ToString("C2"); // Convert to currency format
                                moneyString = Regex.Replace(moneyString, @"^[A-Z]{2}\$", "$"); // Replace prefix like "CA$" with "$"

                                return moneyString;
                            }

                            if (aliasedValue.Value is DateTime dateValue)
                            {
                                tracingService.Trace(attribute + ": Raw DateTime: " + dateValue.ToString());
                                tracingService.Trace(attribute + ": Formatted DateTime (ISO 8601): " + dateValue.ToString("o"));
                                return dateValue.ToString("o"); // Format dates as round-trip date/time pattern
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


                string GetAliasedValueOrDefaultName(string attribute, string defaultValue = "")
                {
                    try
                    {
                        return invoice.Contains(attribute) ? (invoice[attribute] as AliasedValue)?.Value.ToString() ?? defaultValue : defaultValue;
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace("Error retrieving aliased value for attribute " + attribute + ": " + ex.ToString());
                        return defaultValue;
                    }
                }
                // Populate placeholders
                string rentalOrderId = GetAliasedValueOrDefault("ar.avpx_rentalorderid");
                //string invoiceNo = GetValueOrDefault("avpx_invoiceno");
                string invoiceNo = GetValueOrDefault("av_idnew");


                int userTimeZone = GetUserTimezone(context.UserId, service, tracingService);
                tracingService.Trace($"User TimeZone: {userTimeZone}");

                DateTime invoiceDateUtc = DateTime.Parse(GetValueOrDefault("avpx_invoicedate", DateTime.UtcNow.ToString("o")));
                // DateTime rentalOutDateUtc = DateTime.Parse(GetValueOrDefault("avpx_tentativerentalstartdate", DateTime.UtcNow.ToString("o")));
                DateTime rentalInDateUtc = DateTime.Parse(GetAliasedValueOrDefault("ar.avpx_estimatestartdate", DateTime.UtcNow.ToString("o")));
                DateTime printedOnUtc = DateTime.Parse(DateTime.Now.ToString("o"));
                DateTime billFromUtc = DateTime.Parse(GetValueOrDefault("avpx_billfromdate", DateTime.UtcNow.ToString("o")));
                DateTime billToUtc = DateTime.Parse(GetValueOrDefault("avpx_billto", DateTime.UtcNow.ToString("o")));
                DateTime dueDateUtc = DateTime.Parse(GetAliasedValueOrDefault("ar.avpx_estimateenddate", DateTime.UtcNow.ToString("o")));
                
                DateTime invoiceDatePartSaleUtc = DateTime.Parse(GetValueOrDefault("av_invoicedate", DateTime.UtcNow.ToString("o")));
                DateTime invoiceShippingDatePartSaleUtc = DateTime.Parse(GetValueOrDefault("av_shippingdate", DateTime.UtcNow.ToString("o")));
                tracingService.Trace($"Rental Out Date (UTC): {invoiceDateUtc}");
                //tracingService.Trace($"Rental In Date (UTC): {rentalOutDateUtc}");
                tracingService.Trace($"Effective To Date (UTC): {rentalInDateUtc}");
                tracingService.Trace($"Printed On Date (UTC): {printedOnUtc}");

                DateTime invoiceDateUtcLocal = GetLocalTime(invoiceDateUtc, userTimeZone, service, tracingService);
               // DateTime rentalOutDateLocal = GetLocalTime(rentalOutDateUtc, userTimeZone, service, tracingService);
                DateTime rentalInDateUtcLocal = GetLocalTime(rentalInDateUtc, userTimeZone, service, tracingService);
                DateTime printedOnUtcLocal = GetLocalTime(printedOnUtc, userTimeZone, service, tracingService);
                DateTime billFromUtcLocal = GetLocalTime(billFromUtc, userTimeZone, service, tracingService);
                DateTime billToUtcLocal = GetLocalTime(billToUtc, userTimeZone, service, tracingService);


                DateTime invoiceDatePartSaleUtcLocal = GetLocalTime(invoiceDatePartSaleUtc, userTimeZone, service, tracingService);
                DateTime invoiceShippingDatePartSaleUtcLocal = GetLocalTime(invoiceShippingDatePartSaleUtc, userTimeZone, service, tracingService);
                DateTime dueDateUtcLocal = GetLocalTime(dueDateUtc, userTimeZone, service, tracingService);


                // Check if 'paymenttermscode' exists in the linked 'account' entity
                if (invoice.Contains("ac.paymenttermscode"))
                {
                    // Extract the AliasedValue and cast the Value to OptionSetValue
                    AliasedValue aliasedPaymentTerms = (AliasedValue)invoice["ac.paymenttermscode"];
                    OptionSetValue paymentTermsValue = (OptionSetValue)aliasedPaymentTerms.Value;

                    // Use the corrected GetOptionSetValueLabel function
                    string paymentTerms = GetOptionSetValueLabel(service, "account", "paymenttermscode", paymentTermsValue.Value);

                    placeholders["pxAcntTerms"] = paymentTerms;
                    // Output or use the payment terms label
                    tracingService.Trace("Payment Terms Label: " + paymentTerms);
                }
                else
                {
                    tracingService.Trace("Payment terms code not found.");
                }


                tracingService.Trace($"Rental Out Date (Local): {invoiceDateUtcLocal}");
               // tracingService.Trace($"Rental In Date (Local): {rentalOutDateLocal}");
                tracingService.Trace($"Effective To Date (Local): {rentalInDateUtcLocal}");
                tracingService.Trace($"Printed On Date (Local): {printedOnUtcLocal}");

                placeholders["pxInvoiceNumber"] = string.IsNullOrEmpty(rentalOrderId) || string.IsNullOrEmpty(invoiceNo) ? "" : rentalOrderId + " - " + invoiceNo;

                placeholders["pxCustomerNo"] = GetAliasedValueOrDefault("ac.accountnumber", "N/A");
                placeholders["pxEffectiveTo"] = invoiceDateUtcLocal.ToString("MM/dd/yyyy");
               // placeholders["pxRentalOut"] = rentalOutDateLocal.ToString("MM/dd/yyyy");
                placeholders["pxRentalIn"] = rentalInDateUtcLocal.ToString("MM/dd/yyyy");
                placeholders["pxPONumber"] = GetAliasedValueOrDefault("ar.avpx_ponumber");
                placeholders["pxOrderedby"] = GetValueOrDefault("avpx_customerscontact");
                placeholders["pxsalesperson"] = GetValueOrDefault("ownerid", "N/A");
                placeholders["pxReservedBy"] = GetValueOrDefault("reservedby");
                placeholders["pxInvoiceAmount"] = GetFormattedValueOrDefault("avpx_totalamount");
               // placeholders["pxSMSubtotal"] = GetFormattedValueOrDefault("avpx_additionalcharges");
                placeholders["pxDAmount"] = GetFormattedValueOrDefault("avpx_discountamount");
                placeholders["pxDWAmount"] = GetFormattedValueOrDefault("avpx_damagewaiveramountnew");
                placeholders["pxASubtotal"] = GetFormattedValueOrDefault("avpx_subtotalamount");
                placeholders["pxTAmount"] = GetFormattedValueOrDefault("avpx_totaltax");
                placeholders["pxTamt"] = GetFormattedValueOrDefault("avpx_totalamount");
                placeholders["pxPartsSubtotal"] = GetFormattedValueOrDefault("avpx_salesamount");
                placeholders["pxRentalSubtotal"] = GetFormattedValueOrDefault("avpx_unitsprice");
                placeholders["pxEnvironmentFee"] = GetFormattedValueOrDefault("avpx_environmentfee");
                placeholders["pxprintedOn"] = printedOnUtcLocal.ToString("MM/dd/yyyy");
                placeholders["pxDueDate"] = dueDateUtcLocal.ToString("MM/dd/yyyy");
                //placeholders["pxOffNumber"] = GetAliasedValueOrDefault("ac.telephone1", "N/A");
                placeholders["pxprintedOn"] = DateTime.Now.ToString("MM/dd/yyyy");
                placeholders["pxCustomerName"] = GetAliasedValueOrDefaultName("ac.name");
                //placeholders["pxCustomerName"] = fullName;
                placeholders["pxCellNo"] = GetAliasedValueOrDefault("ac.telephone1", "");
                //placeholders["pxCellNo"] = contactTel;
                placeholders["pxBillFrom"] = billFromUtcLocal.ToString("MM/dd/yyyy");
                placeholders["pxbillto"] = billToUtcLocal.ToString("MM/dd/yyyy");
                placeholders["pxVia"] = GetValueOrDefault("av_shippedvia");
                if (orderType == 783090000) //rental
                {
                    //Jobsite Address
                    placeholders["pxStreetAddress"] = GetAliasedValueOrDefault("ar.avpx_jobsitestreetaddress");
                    placeholders["pxCity"] = string.IsNullOrEmpty(GetAliasedValueOrDefault("ar.avpx_jobsitecity")) ? "" : GetAliasedValueOrDefault("ar.avpx_jobsitecity") + ",";
                    placeholders["pxState"] = string.IsNullOrEmpty(GetAliasedValueOrDefault("ar.avpx_jobsitestateorprovince")) ? "" : GetAliasedValueOrDefault("ar.avpx_jobsitestateorprovince") + ",";
                    placeholders["pxPostalCode"] = string.IsNullOrEmpty(GetAliasedValueOrDefault("ar.avpx_jobsitepostalcode")) ? "" : GetAliasedValueOrDefault("ar.avpx_jobsitepostalcode") + ",";
                    placeholders["pxCountry"] = GetAliasedValueOrDefault("ar.avpx_jobsitecountry");

                    //Branch Address
                    placeholders["pxcusAdd"] = GetAliasedValueOrDefault("ac.address1_line2");
                    placeholders["pxcustomerCit"] = string.IsNullOrEmpty(GetAliasedValueOrDefault("ac.address1_city")) ? "" : GetAliasedValueOrDefault("ac.address1_city") + ",";
                    placeholders["pxcusS"] = string.IsNullOrEmpty(GetAliasedValueOrDefault("ac.address1_stateorprovince")) ? "" : GetAliasedValueOrDefault("ac.address1_stateorprovince") + ",";
                    placeholders["pxcusP"] = string.IsNullOrEmpty(GetAliasedValueOrDefault("ac.address1_postalcode")) ? "" : GetAliasedValueOrDefault("ac.address1_postalcode") + ",";
                    placeholders["pxcusCo"] = GetAliasedValueOrDefault("ac.address1_country");
                    placeholders["pxInvoiceNumber"] = string.IsNullOrEmpty(rentalOrderId) || string.IsNullOrEmpty(invoiceNo) ? "" : rentalOrderId + " - " + invoiceNo;

                    placeholders["pxFuelAmt"] = GetFormattedValueOrDefault("av_fuelsurcharge");
                    placeholders["pxFloatAmt"] = GetFormattedValueOrDefault("av_floatcharge");
                    placeholders["pxAdminAmt"] = GetFormattedValueOrDefault("av_adminfee");
                    placeholders["pxMisAmt"] = GetFormattedValueOrDefault("av_missingmanualdocumentscharge");
                    placeholders["pxKeyAmt"] = GetFormattedValueOrDefault("av_missingkeycharge");
                    placeholders["pxDepAmt"] = GetFormattedValueOrDefault("av_deposittaken");
                    placeholders["pxPONumber"] = GetAliasedValueOrDefault("ar.avpx_ponumber");
                    placeholders["pxDelivery"] = GetFormattedValueOrDefault("av_delivery");
                    placeholders["pxPickUp"] = GetFormattedValueOrDefault("av_pickup");
                    placeholders["pxSMSubtotal"] = GetFormattedValueOrDefault("avpx_additionalcharges");
                    //placeholders["pxAcntTerms"] = GetValueAsString(invoice, "ac.paymenttermscode",service,tracingService);
                    placeholders["pxRental"] = GetFormattedValueOrDefault("avpx_totalrentalamountnew");
                    placeholders["pxDays"] = GetValueOrDefault("avpx_invoicedays");
                    placeholders["pxRcId"] = GetAliasedValueOrDefault("ar.avpx_rentalorderid");
                    placeholders["pxEmail"] = GetAliasedValueOrDefault("ac.emailaddress1");
                    placeholders["pxNotes"] = GetValueOrDefault("avpx_notes");
                }
                else if(orderType == 783090001) //sale
                {
                    placeholders["pxInvoiceNumber"] = GetValueOrDefault("av_idnew");
             /*       placeholders["pxStreetAddress"] = GetAliasedValueOrDefault("ab.avpx_streetaddressda");
                    placeholders["pxCity"] = GetAliasedValueOrDefault("ab.avpx_cityda");
                    placeholders["pxState"] = string.IsNullOrEmpty(GetAliasedValueOrDefault("ab.avpx_stateorprovinceda")) ? "" : GetAliasedValueOrDefault("ab.avpx_stateorprovinceda") + ",";
                    placeholders["pxPostalCode"] = GetAliasedValueOrDefault("ab.avpx_postalcodeda");
                    placeholders["pxCountry"] = GetAliasedValueOrDefault("ab.avpx_countryda");*/


                    placeholders["pxStreetAddress"] = GetValueOrDefault("avpx_jsl_streetaddress");
                    placeholders["pxCity"] = GetValueOrDefault("avpx_jsl_city");
                    placeholders["pxState"] = string.IsNullOrEmpty(GetValueOrDefault("avpx_jsl_stateorprovince")) ? "" : GetValueOrDefault("avpx_jsl_stateorprovince") + ",";
                    placeholders["pxPostalCode"] = GetValueOrDefault("avpx_jsl_postalcode");
                    placeholders["pxCountry"] = GetValueOrDefault("avpx_jsl_country");

                    placeholders["pxcusAdd"] = GetAliasedValueOrDefault("ac.address1_line2");
                    placeholders["pxcustomerCit"] = GetAliasedValueOrDefault("ac.address1_city");
                    placeholders["pxcusS"] = String.IsNullOrEmpty( GetAliasedValueOrDefault("ac.address1_stateorprovince")) ? "" : GetAliasedValueOrDefault("ac.address1_stateorprovince") + ",";
                    placeholders["pxcusP"] = GetAliasedValueOrDefault("ac.address1_postalcode");
                    placeholders["pxcusCo"] = GetAliasedValueOrDefault("ac.address1_country");
                    placeholders["pxAdminAmt"] = GetFormattedValueOrDefault("av_adminfee");
                    placeholders["pxInvoiceNumber"] = invoiceNo;
                    placeholders["pxNotes"] = GetValueOrDefault("avpx_notes");
                    placeholders["pxTransCh"] = GetFormattedValueOrDefault("av_transportcharge");
                    
                }
                else if(orderType == 783090003) //Part Sale
                {
                    placeholders["pxStreetAddress"] = GetAliasedValueOrDefault("ac.address1_line2");
                    placeholders["pxCity"] = GetAliasedValueOrDefault("ac.address1_city");
                    placeholders["pxState"]  = String.IsNullOrEmpty(GetAliasedValueOrDefault("ac.address1_stateorprovince")) ? "" : GetAliasedValueOrDefault("ac.address1_stateorprovince") + ",";
                    placeholders["pxPostalCode"] = GetAliasedValueOrDefault("ac.address1_postalcode");
                    placeholders["pxCountry"] = GetAliasedValueOrDefault("ac.address1_country");


                    placeholders["pxcusAdd"] = GetValueOrDefault("av_streetaddress");
                    placeholders["pxcustomerCit"] = GetValueOrDefault("av_city");
                    placeholders["pxcusS"] = String.IsNullOrEmpty(GetValueOrDefault("av_stateprovince")) ? "" : GetValueOrDefault("av_stateprovince") + ",";
                    placeholders["pxcusP"] = GetValueOrDefault("av_postalcodezip");
                    placeholders["pxcusCo"] = GetValueOrDefault("av_country");


                    placeholders["pxInvoiceNumber"] = GetValueOrDefault("av_idnew");
                    placeholders["pxEffectiveTo"] = invoiceDatePartSaleUtcLocal.ToString("MM/dd/yyyy");
                    placeholders["pxshippingDate"] = invoiceShippingDatePartSaleUtcLocal.ToString("MM/dd/yyyy");
                    placeholders["pxTerms"] = GetValueAsString(invoice, "av_terms", service,tracingService);
                    placeholders["pxASubtotal"] = GetFormattedValueOrDefault("av_partsamount");
                    placeholders["pxTAmount"] = GetFormattedValueOrDefault("av_hst13");
                    placeholders["pxTamt"] = GetFormattedValueOrDefault("av_totalamount");
                    placeholders["pxPONumber"] = GetValueOrDefault("av_ponumber");
                    placeholders["pxsalesperson"] = GetValueOrDefault("ownerid", "N/A");
                    placeholders["pxAccName"] = GetAliasedValueOrDefault("ac.name");
                }
                else if (orderType == 783090004) //Float Invoice
                {
                    placeholders["pxStreetAddress"] = GetValueOrDefault("av_streetaddressst");
                    placeholders["pxCity"] = GetValueOrDefault("av_cityst");
                    placeholders["pxState"] = String.IsNullOrEmpty(GetValueOrDefault("av_stateprovincest")) ? "" : GetValueOrDefault("av_stateprovincest") + ",";
                    placeholders["pxPostalCode"] = GetValueOrDefault("av_postalcodezipst");
                    placeholders["pxCountry"] = GetValueOrDefault("av_countryst");


                    placeholders["pxFrStAdd"] = GetValueOrDefault("av_streetaddress");
                    placeholders["pxFrCity"] = GetValueOrDefault("av_city");
                    placeholders["pxFrState"] = String.IsNullOrEmpty(GetValueOrDefault("av_stateprovince")) ? "" : GetValueOrDefault("av_stateprovince") + ",";
                    placeholders["pxFrPC"] = GetValueOrDefault("av_postalcodezip");
                    placeholders["pxFrCon"] = GetValueOrDefault("av_country");


                    placeholders["pxcusAdd"] = GetAliasedValueOrDefault("ac.address1_line2");
                    placeholders["pxcustomerCit"] = GetAliasedValueOrDefault("ac.address1_city");
                    placeholders["pxcusS"] = String.IsNullOrEmpty(GetAliasedValueOrDefault("ac.address1_stateorprovince")) ? "" : GetAliasedValueOrDefault("ac.address1_stateorprovince") + ",";
                    placeholders["pxcusP"] = GetAliasedValueOrDefault("ac.address1_postalcode");
                    placeholders["pxcusCo"] = GetAliasedValueOrDefault("ac.address1_country");


                    DateTime transportDateUtc = DateTime.Parse(GetValueOrDefault("av_transportdate", DateTime.UtcNow.ToString("o")));
                    DateTime transportDateUtcLocal = GetLocalTime(transportDateUtc, userTimeZone, service, tracingService);
                    placeholders["pxInvoiceNumber"] = GetValueOrDefault("av_idnew");
                    placeholders["pxEffectiveTo"] = transportDateUtcLocal.ToString("MM/dd/yyyy");
                    placeholders["pxbillto"] = invoiceDateUtcLocal.ToString("MM/dd/yyyy");
                    //placeholders["pxbillto"] = 
                    placeholders["pxAdminAmt"] = GetFormattedValueOrDefault("av_adminfeeforfloattype");
                    placeholders["pxPONumber"] = GetValueOrDefault("av_ponumber");
                    placeholders["pxFloatAmt"] = GetFormattedValueOrDefault("av_floatchargeforfloattype");
                    placeholders["pxASubtotal"] = GetFormattedValueOrDefault("av_subtotal");
                    placeholders["pxTAmount"] = GetFormattedValueOrDefault("av_hst");
                    placeholders["pxTamt"] = GetFormattedValueOrDefault("av_total");
                    placeholders["pxDescription"] = GetValueOrDefault("av_unitdescription");
                    placeholders["pxNotes"] = GetValueOrDefault("avpx_notes");
                    placeholders["pxEmail"] = GetAliasedValueOrDefault("ac.emailaddress1");
                    
                    
                }
                

                return placeholders;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error in extracting invoice data: " + ex.ToString());
                throw new InvalidPluginExecutionException("Failed to extract invoice data. " + ex.Message);
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
                        docText = docText.Replace("{{", "")
                 .Replace("}}", "")
                 .Replace("}},", ",")
                 .Replace("},{{", ",")
                 .Replace("}},", "")
                 .Replace(",{{", "")
                 .Replace("}},", "")
                 .Replace(",", "")

                 ;

                        tracingService.Trace("Main document part read successfully.");

                        string EscapeXmlSpecialCharacters(string text)
                        {
                            if (string.IsNullOrEmpty(text))
                                return text;

                            var replacements = new Dictionary<string, string>
                    {
                        { "&", "&amp;" },
                        { "<", "&lt;" },
                        { ">", "&gt;" },
                        { "\"", "&quot;" },
                        { "'", "&apos;" },
                        { "`", "&#x60;" },
                        { "=", "&#x3D;" },
                        { "@", "&#x40;" },
                        { "{", "&#x7B;" },
                        { "}", "&#x7D;" },
                        { "[", "&#x5B;" },
                        { "]", "&#x5D;" },
                        { "(", "&#x28;" },
                        { ")", "&#x29;" },
                        { ";", "&#x3B;" },
                        { ":", "&#x3A;" },
                        { "/", "&#x2F;" },
                        { "\\", "&#x5C;" },
                        { "|", "&#x7C;" },
                        { "^", "&#x5E;" },
                        { "~", "&#x7E;" }
                    };

                            foreach (var kvp in replacements)
                            {
                                text = text.Replace(kvp.Key, kvp.Value);
                            }

                            return text;
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

                        // Split combined data into assetGroup, charges, and parts
                        var assetGroup = new List<string[]>();
                        var charges = new List<string[]>();
                        var parts = new List<string[]>();
                        bool isChargesSection = false;
                        bool isPartsSection = false;

                        tracingService.Trace("Splitting combined data into assetGroup, charges, and parts.");

                        foreach (var record in combinedData)
                        {
                            if (record.Length == 1 && record[0] == "---- End of AssetGroup ----")
                            {
                                isChargesSection = true;
                                continue;
                            }

                            if (record.Length == 1 && record[0] == "---- End of Charges ----")
                            {
                                isPartsSection = true;
                                continue;
                            }

                            if (isPartsSection)
                            {
                                parts.Add(record);
                            }
                            else if (isChargesSection)
                            {
                                charges.Add(record);
                            }
                            else
                            {
                                assetGroup.Add(record);
                            }
                        }

                        tracingService.Trace($"Data split completed. AssetGroup count: {assetGroup.Count}, Charges count: {charges.Count}, Parts count: {parts.Count}");

                        // Find tables by caption
                        var tables = wordDoc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList();
                        bool assetGroupTableFound = false;
                        bool chargesTableFound = false;
                        bool partsTableFound = false;

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
                                else if (altText.Contains("Parts"))
                                {
                                    tracingService.Trace("Parts table identified by alt text.");
                                    PopulateTableWithData(table, parts, imageData, wordDoc, tracingService);
                                    partsTableFound = true;
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



        /* static void PopulateTableWithData(DocumentFormat.OpenXml.Wordprocessing.Table table, List<string[]> tableData, byte[] imageData, WordprocessingDocument wordDoc, ITracingService tracingService)
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
 */

        static void PopulateTableWithData(DocumentFormat.OpenXml.Wordprocessing.Table table, List<string[]> tableData, byte[] imageData, WordprocessingDocument wordDoc, ITracingService tracingService)
        {
            tracingService.Trace("Starting PopulateTableWithData.");

            // Initialize a counter for the first column
            int rowIndex = 1;

            foreach (var rowData in tableData)
            {
                tracingService.Trace("Adding new row.");

                DocumentFormat.OpenXml.Wordprocessing.TableRow newRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();

                bool isFirstCell = true; // Track if it's the first cell of the row

                foreach (var cellData in rowData)
                {
                    string cellContent;

                    // For the first cell, use the rowIndex
                    if (isFirstCell)
                    {
                        cellContent = rowIndex.ToString();
                        isFirstCell = false;
                    }
                    else
                    {
                        cellContent = cellData;
                    }

                    tracingService.Trace($"Adding new cell with data: {cellContent}");

                    DocumentFormat.OpenXml.Wordprocessing.TableCell newCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                        new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                            new DocumentFormat.OpenXml.Wordprocessing.Run(
                                new DocumentFormat.OpenXml.Wordprocessing.Text(cellContent)
                            )
                        )
                    );
                    newRow.Append(newCell);
                }

                // Add image if provided
                if (imageData != null)
                {
                    tracingService.Trace("Adding image to row.");

                    DocumentFormat.OpenXml.Wordprocessing.TableCell imageCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                    AddImageToCell(wordDoc, imageCell, imageData, tracingService);
                    newRow.Append(imageCell);
                }

                // Append the new row to the table
                table.Append(newRow);

                // Increment the row index for the next row
                rowIndex++;
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
            var data = new List<string[]>();
            var assetGroup = new List<string[]>();
            var charges = new List<string[]>();
            var parts = new List<string[]>();

            decimal localTQS = 0;
            try
            {
                // Define the base Fetch XML for AssetGroup and Charge
                string fetchXmlBase = @"
                    <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='avpx_invoicelineitems'>
                        <attribute name='avpx_name' />
                        <attribute name='avpx_assetname' />
                        <attribute name='avpx_quantity' />
                        <attribute name='avpx_amount' />
                        <attribute name='avpx_extendedamount' />
                        <attribute name='avpx_priceperunit' />
                        <attribute name='av_description' />
                        <attribute name='avpx_invoicelineitemsid' />
                        <attribute name='avpx_type' />
                        <attribute name='avpx_type' />
                        <attribute name='avpx_invoiceid' />
                        <order attribute='avpx_type' descending='false' />
                        <filter type='and'>
                          <condition attribute='statuscode' operator='eq' value='1' />
                          <condition attribute='avpx_invoiceid' operator='eq' uitype='avpx_invoice' value='{0}' />
                          <condition attribute='avpx_type' operator='neq' value='783090002' /> <!-- Exclude Part/Kit -->
                        </filter>
                      </entity>
                    </fetch>";

                // Retrieve AssetGroup and Charge records
                string fetchXml = string.Format(fetchXmlBase, recordId);
                tracingService.Trace("Fetch XML for AssetGroup and Charge: " + fetchXml);
                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXml));

                // Process AssetGroup and Charge records
                if (entityCollection.Entities.Count > 0)
                {
                    tracingService.Trace("Records retrieved successfully. Count: {0}", entityCollection.Entities.Count);

                    foreach (Entity entity in entityCollection.Entities)
                    {


                        string avpxQuantity = RoundOffValue(entity, "avpx_quantity");
                        string avpxName = GetValueAsString(entity, "avpx_name", service, tracingService);
                        string avpxAmount = GetFormattedValueOrDefault(entity, "avpx_priceperunit", tracingService);
                        string avpxExtendedAmount = GetFormattedValueOrDefault(entity, "avpx_amount", tracingService);
                        string avpxType = GetValueAsString(entity, "avpx_type", service, tracingService);
                        string avDescription = GetValueAsString(entity, "av_description", service, tracingService);
                        tracingService.Trace("Type: " + avpxType);
                        // Use avDescription if it is not empty and not "N/A"; otherwise, fall back to avpxName
                        string displayName = (!string.IsNullOrEmpty(avDescription) && avDescription != "N/A") ? avDescription : avpxName;
                        int avpxOrderType = 0;

                        if (entity.Contains("avpx_invoiceid") && entity["avpx_invoiceid"] is EntityReference invoiceRef)
                        {
                            Guid invoiceId = invoiceRef.Id;

                            // Retrieve avpx_ordertype from avpx_invoice
                            ColumnSet columns = new ColumnSet("avpx_ordertype");
                            Entity invoiceEntity = service.Retrieve("avpx_invoice", invoiceId, columns);

                            // Default value
                            if (invoiceEntity != null && invoiceEntity.Contains("avpx_ordertype"))
                            {
                                avpxOrderType = invoiceEntity.GetAttributeValue<OptionSetValue>("avpx_ordertype")?.Value ?? 0;
                            }

                            tracingService.Trace("Retrieved Order Type: " + avpxOrderType);

                            // Use avpxOrderType in conditional logic
                            
                        }



                        string[] recordData = null;

                        if (avpxOrderType == 783090001 && avpxType == "Asset group") // SALES
                        {
                            recordData = new string[] { "Unit Information", avpxQuantity,displayName, avpxExtendedAmount };
                            assetGroup.Add(recordData);
                        }
                        else if (avpxType == "Asset group") //RENTAL
                        {
                            recordData = new string[] { "Unit Information", displayName, avpxAmount, avpxExtendedAmount };
                            assetGroup.Add(recordData);
                        }
                        else // Asset group and other types
                        {
                            recordData = new string[] { "Unit Information", displayName, avpxAmount, avpxExtendedAmount };
                            assetGroup.Add(recordData);
                        }

                        if (recordData != null)
                        {
                            data.Add(recordData);
                        }
                    }
                }
                else
                {
                    tracingService.Trace("No records retrieved for AssetGroup and Charge.");
                }

                // Define the Fetch XML for Part/Kit
                string fetchXmlPartKit = @"
                        <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='avpx_invoicelineitems'>
                            <attribute name='avpx_name' />
                            <attribute name='avpx_assetname' />
                            <attribute name='avpx_quantity' />
                            <attribute name='avpx_amount' />
                            <attribute name='avpx_extendedamount' />
                            <attribute name='avpx_priceperunit' />
                            <attribute name='avpx_invoicelineitemsid' />
                            <attribute name='av_quantitybackordered' />
                            <attribute name='av_description' />
                            <attribute name='av_quantityordered' />
                            <attribute name='avpx_type' />
                            <attribute name='avpx_part' />
                            <order attribute='avpx_type' descending='false' />
                            <filter type='and'>
                              <condition attribute='statuscode' operator='eq' value='1' />
                              <condition attribute='avpx_invoiceid' operator='eq' uitype='avpx_invoice' value='{0}' />
                              <condition attribute='avpx_type' operator='eq' value='783090002' /> <!-- Include only Part/Kit -->
                            </filter>
                          </entity>
                        </fetch>";
                    
                // Retrieve Part/Kit records
                fetchXml = string.Format(fetchXmlPartKit, recordId);
                tracingService.Trace("Fetch XML for Part/Kit: " + fetchXml);
                entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXml));

                // Process Part/Kit records
                if (entityCollection.Entities.Count > 0)  
                {
                    tracingService.Trace("Part/Kit records retrieved successfully. Count: {0}", entityCollection.Entities.Count);

                    foreach (Entity entity in entityCollection.Entities)
                    {
                        string avpxQuantity = RoundOffValue(entity, "av_quantityordered");
                        tracingService.Trace("Retrived avpxQuantity");
                        string avpxName = GetValueAsString(entity, "avpx_name", service, tracingService);
                        tracingService.Trace("Retrived avpxName");
                        string avpxAmount = GetFormattedValueOrDefault(entity, "avpx_priceperunit", tracingService);
                        tracingService.Trace("Retrived avpxAmount");
                        string avpxExtendedAmount = GetFormattedValueOrDefault(entity, "avpx_amount", tracingService);
                        tracingService.Trace("Retrived avpxExtendedAmount");
                        string avpxQuantityBackOrder = GetFormattedValueOrDefault(entity, "av_quantitybackordered", tracingService);
                        tracingService.Trace("Retrived avpxQuantityBackOrder");
                        string avpxPartName = "";

                        string qtyShipped = GetNumericValueOrDefault(entity, "avpx_quantity", tracingService);
                        tracingService.Trace("Retrieved qtyShipped: " + qtyShipped);

                        if (decimal.TryParse(qtyShipped, out decimal qtyShippedDecimal))
                        {
                            tracingService.Trace("Retrieved qtyShippedDecimal: " + qtyShippedDecimal);
                            localTQS += qtyShippedDecimal;
                            tracingService.Trace("Updated localTQS: " + localTQS);
                        }
                        else
                        {
                            tracingService.Trace("qtyShipped could not be parsed as a decimal.");
                        }
                        /*tracingService.Trace("Retrived qtyShippedDecimal");
                        localTQS += qtyShippedDecimal;*/
                        
                        Guid partId = entity.GetAttributeValue<EntityReference>("avpx_part")?.Id ?? Guid.Empty;
                        if (partId != Guid.Empty)
                        {
                            Entity partEntity = service.Retrieve("avpx_parts", partId, new ColumnSet("avpx_name"));
                            avpxPartName = partEntity.Contains("avpx_name") ? partEntity["avpx_name"].ToString() : string.Empty;
                            tracingService.Trace("Part Name: " + avpxPartName);
                        }
                        tracingService.Trace("Retrived avpxPartName");
                        string avDescription = GetValueAsString(entity, "av_description", service, tracingService);
                        // Use avDescription if it is not empty and not "N/A"; otherwise, fall back to avpxName
                         string displayName = (!string.IsNullOrEmpty(avDescription) && avDescription != "N/A") ? avDescription : avpxPartName;

                        string[] recordData = new string[] { "", displayName, avpxQuantity, avpxQuantityBackOrder, qtyShipped, avpxAmount,"EA", avpxExtendedAmount };
                        parts.Add(recordData);
                        data.Add(recordData);
                    }
                    totalQtyShipped = localTQS.ToString("F2");
                    tracingService.Trace("End");
                }
                else
                {
                    tracingService.Trace("No Part/Kit records retrieved.");
                }


                //Float Line Items


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
            combinedResult.Add(new string[] { "---- End of Charges ----" });
            combinedResult.AddRange(parts);
            tracingService.Trace("Combine Completed");
            return combinedResult;
        }


        // Helper methods (GetValueAsString, GetFormattedValueOrDefault, RoundOffValue) remain unchanged


        // Helper methods (GetValueAsString, GetFormattedValueOrDefault, RoundOffValue) remain unchanged


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
                            return dateTime.ToString("MM/dd/yyyy"); // Format as needed
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
            return "0";
        }

        /*private string GetFormattedValueOrDefault(Entity entity, string attributeName, ITracingService tracingService, string defaultValue = "N/A")
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
        }*/

   

    private string GetFormattedValueOrDefault(Entity entity, string attributeName, ITracingService tracingService, string defaultValue = "N/A")
    {
        try
        {
            if (entity.FormattedValues.Contains(attributeName))
            {
                var value = entity.FormattedValues[attributeName];

                // Use regex to remove any country prefix before the dollar sign
                value = Regex.Replace(value, @"^[A-Z]{2}\$", "$"); // Replace prefix like "CA$" with "$"

                if (decimal.TryParse(value.Replace("$", "").Trim(), out decimal decimalValue))
                {
                    return "$" + decimalValue.ToString("N2"); // Format with dollar sign and two decimal places
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

        private string GetNumericValueOrDefault(Entity entity, string attributeName, ITracingService tracingService, string defaultValue = "N/A")
        {
            try
            {
                if (entity.FormattedValues.Contains(attributeName))
                {
                    var value = entity.FormattedValues[attributeName];

                    // Clean out non-numeric characters, except decimal points, commas, and minus signs
                    var cleanedValue = Regex.Replace(value, @"[^0-9.,-]", "").Trim();

                    if (decimal.TryParse(cleanedValue, out decimal decimalValue))
                    {
                        return decimalValue.ToString("N2"); // Format with two decimal places
                    }

                    tracingService.Trace("Numeric value could not be parsed as a decimal: " + value);
                    return defaultValue; // Return default if value is not numeric
                }
                return defaultValue;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error retrieving numeric value for attribute " + attributeName + ": " + ex.ToString());
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
