using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
/// <summary>
/// This plugin gets triggered on create and update of amount, Additional Charges, Effective Discount Amount, Subtotal Amount, Tax Amount and Extended Amount and on delete of ILI, Based on this triggers this plugin will update the Rollup Inovice Line Items to Invocie.
/// <summary>

namespace AV.Prexa.Plugins.InvoiceLineItem
{
    public class RollupInvoiceLineItemsToInvoice : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            try
            {
                if (context.PrimaryEntityName.ToLower() != "avpx_invoicelineitems") return;
                if (context.MessageName.ToUpper() == "CREATE" || context.MessageName.ToUpper() == "UPDATE" || context.MessageName.ToUpper() == "DELETE")
                {
                    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                    {
                        Entity entity = (Entity)context.InputParameters["Target"];
                        ProcessLogic(entity, context, service, tracingService);
                    }
                    else if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                    {
                        ProcessLogic(null, context, service, tracingService);
                    }
                    else
                    {
                        tracingService.Trace("Tracing: Input Parameter is not an entity.");
                    }
                }
            }
            catch (FaultException ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: Execute >> " + ex.ToString());
                throw ex;
            }
        }

        private void ProcessLogic(Entity targetEntity, IPluginExecutionContext context, IOrganizationService service, ITracingService tracingService)
        {
            try
            {
                EntityReference invoice = null;
                Entity invoiceEntity = null;

                if (context.MessageName.ToUpper() == "CREATE")
                {
                    invoice = targetEntity.GetAttributeValue<EntityReference>("avpx_invoiceid");
                }
                else if (context.MessageName.ToUpper() == "UPDATE")
                {
                    targetEntity = service.Retrieve("avpx_invoicelineitems", context.PrimaryEntityId, new ColumnSet("avpx_invoiceid"));
                    invoice = targetEntity.GetAttributeValue<EntityReference>("avpx_invoiceid");
                }
                else if (context.MessageName.ToUpper() == "DELETE")
                {
                    Entity DeleteQLI = (Entity)context.PreEntityImages["Pre-Image"];
                    invoice = (EntityReference)DeleteQLI.Attributes["avpx_invoiceid"];
                }

                if (invoice != null)
                {
                    decimal adminFee = 0;
                    decimal totalTaxAmount = 0;
                    invoiceEntity = service.Retrieve("avpx_invoice", invoice.Id, new ColumnSet("avpx_damagewaiverpercentage", "avpx_taxcode", "avpx_ordertype"));

                    decimal damageWaiverPercentage = invoiceEntity.Contains("avpx_damagewaiverpercentage") ? Convert.ToDecimal(invoiceEntity["avpx_damagewaiverpercentage"]) : 0;

                    int orderTypeValue = invoiceEntity.Contains("avpx_ordertype") && invoiceEntity["avpx_ordertype"] != null ? ((OptionSetValue)invoiceEntity["avpx_ordertype"]).Value : -1;

                    if(orderTypeValue == 783090000) //Rental
                    {
                        adminFee = 40;
                    }else if(orderTypeValue == 783090001) //Sale
                    {
                        adminFee = 40;
                    }
                    
                    Guid taxCodeId = invoiceEntity.Contains("avpx_taxcode") ? invoiceEntity.GetAttributeValue<EntityReference>("avpx_taxcode").Id : Guid.Empty;
                    
                    decimal unitsPrice = GetAmountsOfInvoiceLineItemsBasedOnType(invoice.Id, "avpx_amount", "avpx_unitsprice", "783090000", service, tracingService);
                    decimal additionalCharges = GetAmountsOfInvoiceLineItemsBasedOnType(invoice.Id, "avpx_amount", "avpx_additionalcharges", "783090001", service, tracingService);
                    decimal effectiveDiscountAmount = GetAmountsOfInvoiceLineItems(invoice.Id, "avpx_effectivediscountamount", service, tracingService);
                    //decimal subTotalAmount = GetAmountsOfInvoiceLineItems(invoice.Id, "avpx_subtotalamount", service, tracingService);
                    decimal totalTax = GetAmountsOfInvoiceLineItems(invoice.Id, "avpx_taxamount", service, tracingService);
                    if(totalTax == 0)
                    {
                        totalTax = adminFee * 13 / 100;

                    }
                    else
                    {
                        totalTax = totalTax + 5.2m;
                    }
                    tracingService.Trace("Admin Fee" + adminFee);
                    decimal environmentFee = GetAmountsOfInvoiceLineItems(invoice.Id, "avpx_environmentfeeamount", service, tracingService);
                    // decimal totalAmount = GetAmountsOfInvoiceLineItems(invoice.Id, "avpx_extendedamount", service, tracingService);
                    decimal salesAmount = GetAmountsOfInvoiceLineItemsBasedOnType(invoice.Id, "avpx_amount", "avpx_unitsprice", "783090002", service, tracingService);
                    decimal totalDamageWaiverAmount = GetSumOfDamageWaiverAmountFromInvoiceLineItem(invoice.Id, "avpx_discountamount", service, tracingService);
                    decimal damageWaiverAmount = GetDamageWaiverAmount(invoice.Id, totalDamageWaiverAmount, damageWaiverPercentage, service, tracingService);
                    decimal damageWaiverTaxAmount = GetDamageWaiverTaxAmount(invoice.Id, damageWaiverAmount, taxCodeId, service, tracingService);
                    decimal subTotalAmount = unitsPrice - effectiveDiscountAmount + additionalCharges + damageWaiverAmount + environmentFee + salesAmount + adminFee;
                    tracingService.Trace($"SubTotal Amount {subTotalAmount}");
                    if (taxCodeId != Guid.Empty)
                    {
                         totalTaxAmount = subTotalAmount * 13 * 0.01m;
                    }
                    
                    //throw new Exception(String.Format("damageWaiverAmount:{4},subtotalAmount:{0},damageWaiverTaxAmount:{1},avpx_totaltax:{2},totalAmount:{3}", subTotalAmount, damageWaiverTaxAmount, totalTax, (subTotalAmount + totalTax + damageWaiverTaxAmount), damageWaiverAmount));


                    UpdateInvoiceColumns(invoice.Id, unitsPrice, additionalCharges, effectiveDiscountAmount, subTotalAmount, damageWaiverTaxAmount, totalTaxAmount,
                       (Math.Round(subTotalAmount, 2) + totalTaxAmount), environmentFee, salesAmount, totalDamageWaiverAmount, service, tracingService);
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: ProcessLogic >> " + ex.ToString());
                throw ex;
            }
        }

        private decimal GetDamageWaiverAmount(Guid id, decimal damageWaiverEligibleAmount, decimal damageWaiverPercentage, IOrganizationService service, ITracingService tracingService)
        {
            try
            {
                decimal damageWaiverAmount = 0;
                damageWaiverAmount = Convert.ToDecimal(damageWaiverEligibleAmount * damageWaiverPercentage * 0.01M);
                return damageWaiverAmount;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private decimal GetDamageWaiverTaxAmount(Guid id, decimal damageWaiverAmount, Guid taxCodeId, IOrganizationService service, ITracingService tracingService)
        {
            try
            {
                if (taxCodeId == Guid.Empty)
                {
                    return 0;
                }

                decimal taxRate = 0;
                Entity taxCode = service.Retrieve("avpx_taxcode", taxCodeId, new ColumnSet("avpx_actastaxgroup", "avpx_taxrate"));
                if (taxCode.GetAttributeValue<bool>("avpx_actastaxgroup"))
                {
                    string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='avpx_taxcode'>
                                            <attribute name='avpx_taxrate' />
                                            <filter type='and'>
                                              <condition attribute='avpx_parenttaxcode' operator='eq' value='{0}' />
                                              <condition attribute='avpx_actastaxgroup' operator='eq' value='0' />
                                              <condition attribute='statecode' operator='eq' value='0' />
                                            </filter>
                                          </entity>
                                        </fetch>", taxCodeId);

                    EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

                    if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                    {
                        foreach (Entity entity in entityCollection.Entities)
                        {
                            taxRate += entity.GetAttributeValue<decimal>("avpx_taxrate");
                        }
                    }
                }
                else
                {
                    taxRate = taxCode.GetAttributeValue<decimal>("avpx_taxrate");
                }

                decimal damageWaiverTaxAmount = 0;
                damageWaiverTaxAmount = Convert.ToDecimal(damageWaiverAmount * taxRate * 0.01M);
                return damageWaiverTaxAmount;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //This function will get the amount of all invoice line items.
        private decimal GetAmountsOfInvoiceLineItems(Guid invoiceId, string column, IOrganizationService service, ITracingService tracingService)
        {
            try
            {
                decimal sum = 0;
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
									  <entity name='avpx_invoicelineitems'>
									    <attribute name='{1}'/>
									    <filter type='and'>
									      <condition attribute='avpx_invoiceid' operator='eq' value='{0}' />
									      <condition attribute='statecode' operator='eq' value='0' />
									    </filter>
									  </entity>
									</fetch>", invoiceId, column);

                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    foreach (Entity entity in entityCollection.Entities)
                    {
                        if (entity.GetAttributeValue<Money>(column) != null)
                        {
                            sum += entity.GetAttributeValue<Money>(column).Value;
                        }
                    }
                }

                return sum;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetAmountOfInvoiceLineItems >> " + ex.ToString());
                throw ex;
            }
        }

        //This function will get amount based on Type
        private decimal GetAmountsOfInvoiceLineItemsBasedOnType(Guid invoiceId, string column, string updateColumn, string type, IOrganizationService service, ITracingService tracingService)
        {
            try
            {
                //Money totalColumnAmount;
                decimal sum = 0;
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
									  <entity name='avpx_invoicelineitems'>
									    <attribute name='{1}' />
									    <filter type='and'>
									      <condition attribute='avpx_invoiceid' operator='eq' value='{0}' />
									      <condition attribute='statecode' operator='eq' value='0' />
                                          <condition attribute='avpx_type' operator='eq' value='{2}' />
                                        </filter>
									  </entity>
									</fetch>", invoiceId, column, type);

                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    //Entity entity = entityCollection.Entities[0];
                    //totalColumnAmount = (Money)((AliasedValue)entity["columnSum"]).Value;
                    foreach (Entity entity in entityCollection.Entities)
                    {
                        if (entity.GetAttributeValue<Money>(column) != null)
                        {
                            sum += entity.GetAttributeValue<Money>(column).Value;
                        }
                    }
                    //UpdateInvoiceColumns(invoiceId, sum, updateColumn, service, tracingService);
                }

                return sum;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetAmountOfInvoiceLineItems >> " + ex.ToString());
                throw ex;
            }
        }

        //This function will update the invoice fields.
        private void UpdateInvoiceColumns(Guid invoiceId, decimal unitsPrice, decimal additionalCharges, decimal discountAmount, decimal subtotalAmount, decimal damageWaiverTaxAmount, decimal taxAmount,
             decimal totalAmount, decimal environmentFee, decimal salesAmount, decimal totalDamageWaiverAmount, IOrganizationService service, ITracingService tracingService)
        {
            try
            {
                Entity invoiceUpdate = new Entity("avpx_invoice");
                invoiceUpdate["avpx_invoiceid"] = invoiceId;
                invoiceUpdate["avpx_unitsprice"] = new Money(unitsPrice);
                invoiceUpdate["avpx_additionalcharges"] = new Money(additionalCharges);
                invoiceUpdate["avpx_discountamount"] = new Money(discountAmount);
                invoiceUpdate["avpx_subtotalamount"] = new Money(subtotalAmount);
                invoiceUpdate["avpx_damagewaivertaxamount"] = new Money(damageWaiverTaxAmount);
                invoiceUpdate["avpx_totaltax"] = new Money(taxAmount);
                invoiceUpdate["avpx_totalamount"] = new Money(totalAmount);
                invoiceUpdate["avpx_environmentfee"] = new Money(environmentFee);
                invoiceUpdate["avpx_salesamount"] = new Money(salesAmount);
                invoiceUpdate["avpx_damagewaivereligibleamount"] = new Money(totalDamageWaiverAmount);
                //invoiceUpdate["avpx_notes"] = string.Format("avpx_unitsprice:{0}, avpx_additionalcharges:{1},avpx_discountamount:{2}, avpx_subtotalamount:{3}" +
                //    "avpx_damagewaivertaxamount:{4}, avpx_totaltax:{5},avpx_totalamount:{6}", unitsPrice, additionalCharges, discountAmount, subtotalAmount, damageWaiverTaxAmount, 
                //    taxAmount, totalAmount);

                service.Update(invoiceUpdate);
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: UpdateInvoice >> " + ex.ToString());
                throw ex;
            }
        }

        private decimal GetSumOfDamageWaiverAmountFromInvoiceLineItem(Guid invoiceId, string atributeName, IOrganizationService service, ITracingService tracingService)
        {
            //Sum quantity of all Quote line items for same Quote
            try
            {
                decimal sum = 0;
                /*(Start) Modified By Pratik Telaviya on 28-April-23 to fix the issue of */
                /*Earlier we were using fetch xml and aggregate function to sum the taxamount but we found that it returns the sum in base currency only*/
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                              <entity name='avpx_invoicelineitems'>
                                                <attribute name='{0}' />
                                                <filter type='and'>
                                                  <condition attribute='avpx_invoiceid' operator='eq' value='{1}' />
                                                  <condition attribute='statecode' operator='eq' value='0' />
                                                  <condition attribute='avpx_damagewaiverapplicable' operator='ne' value='0' />
                                                  <condition attribute='avpx_type' operator='eq' value='783090000' />
                                                </filter>
                                              </entity>
                                            </fetch>", atributeName, invoiceId);


                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    foreach (Entity invoiceLineItem in entityCollection.Entities)
                    {
                        sum += (invoiceLineItem.Contains(atributeName) && ((Money)(invoiceLineItem[atributeName])) != null) ? ((Money)(invoiceLineItem[atributeName])).Value : 0;
                    }
                    //Entity entity = entityCollection.Entities[0];
                    //taxamount = entity.Contains("totaltaxamountSum") ? ((Money)((AliasedValue)entity["totaltaxamountSum"]).Value != null ? ((Money)((AliasedValue)entity["totaltaxamountSum"]).Value).Value : 0) : 0;
                }

                return sum;
                /*(End) Modified By Pratik Telaviya on 28-April-23 to fix the issue of */
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetSumOfDamageWaiverAmountFromInvoiceLineItem >> " + ex.ToString());
                throw ex;
            }
        }
    }
}
