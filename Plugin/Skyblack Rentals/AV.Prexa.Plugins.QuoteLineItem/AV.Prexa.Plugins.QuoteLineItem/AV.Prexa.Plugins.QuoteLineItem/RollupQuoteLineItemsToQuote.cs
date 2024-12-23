using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
/// <summary>
/// This plugin gets triggered on create and change of amount,EffectiveDiscountAmount, Extended Amount, Tax Amount and on delete of QLI and it rolls up the other amount related fields into quote.
/// <summary>
namespace Av.Prexa.Plugins.QuoteLineItem
{
    public class RollupQuoteLineItemsToQuote : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            try
            {
                if (context.PrimaryEntityName.ToLower() != "avpx_quotelineitem") return;

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
                EntityReference quote = null;
                Entity quoteEntity = null;
                string message = context.MessageName.ToUpper();

                if (message == "CREATE" || message == "UPDATE")
                {
                    targetEntity = service.Retrieve("avpx_quotelineitem", context.PrimaryEntityId, new ColumnSet("avpx_quote"));
                    quote = targetEntity.GetAttributeValue<EntityReference>("avpx_quote");
                }
                else if (message == "DELETE")
                {
                    Entity DeleteQLI = (Entity)context.PreEntityImages["Pre-Image"];
                    quote = (EntityReference)DeleteQLI.Attributes["avpx_quote"];
                }


                if (quote != null)
                {
                    quoteEntity = service.Retrieve("avpx_quote", quote.Id, new ColumnSet("avpx_damagewaiverpercentage", "avpx_taxcode", "av_environmentfee", "av_cardprocessingfee"));

                    decimal damageWaiverPercentage = quoteEntity.Contains("avpx_damagewaiverpercentage") ? Convert.ToDecimal(quoteEntity["avpx_damagewaiverpercentage"]) : 0;
                    Guid taxCodeId = quoteEntity.Contains("avpx_taxcode") ? quoteEntity.GetAttributeValue<EntityReference>("avpx_taxcode").Id : Guid.Empty;
                    decimal extendedAmount = GetSumOfAttributeFromQuoteLineItem(quote.Id, "avpx_discountedamount", 0, service, tracingService);//GetSumOfExtendedAmount(quote.Id, service, tracingService);
                    decimal taxAmount = GetSumOfAttributeFromQuoteLineItem(quote.Id, "avpx_taxamount", 0, service, tracingService); //GetSumOfTaxAmount(quote.Id, service, tracingService);
                    decimal effectiveDiscountAmount = GetSumOfAttributeFromQuoteLineItem(quote.Id, "avpx_effectivediscountamount", 0, service, tracingService); //GetSumOfEffectiveDiscountAmount(quote.Id, service, tracingService);
                    decimal detailedAmount = GetSumOfAttributeFromQuoteLineItem(quote.Id, "avpx_amount", 783090000, service, tracingService); ; // GetSumOfAmount(quote.Id, service, tracingService);
                    decimal chargeAmount = GetSumOfAttributeFromQuoteLineItem(quote.Id, "avpx_amount", 783090001, service, tracingService); //GetSumOfCharges(quote.Id, service, tracingService);
                    decimal totalDamageWaiverAmount = GetSumOfDamageWaiverAmountFromQuoteLineItem(quote.Id, "avpx_discountedamount", service, tracingService);
                    decimal damageWaiverAmount = GetDamageWaiverAmount(quote.Id, totalDamageWaiverAmount, damageWaiverPercentage, service, tracingService);
                    decimal damageWaiverTaxAmount = GetDamageWaiverTaxAmount(quote.Id, damageWaiverAmount, taxCodeId, service, tracingService);
                    decimal rentalTotalTaxAmount = GetSumOfTaxAmount(quote.Id, service, tracingService);
                    decimal environmentFee = quoteEntity.Contains("av_environmentfee") && quoteEntity["av_environmentfee"] != null ? Convert.ToDecimal(quoteEntity["av_environmentfee"]) : 0;
                    decimal cardProcessingFee = quoteEntity.Contains("av_cardprocessingfee") && quoteEntity["av_cardprocessingfee"] != null ? Convert.ToDecimal(quoteEntity["av_cardprocessingfee"]) : 0;
                    tracingService.Trace("cardProcessingFee" + cardProcessingFee);
                    tracingService.Trace("rentalTotalTaxAmount" + rentalTotalTaxAmount);
                    //UpdateQuote(quote.Id, extendedAmount, (taxAmount + damageWaiverTaxAmount), effectiveDiscountAmount, detailedAmount, chargeAmount, damageWaiverAmount,environmentFee, cardProcessingFee, damageWaiverTaxAmount, totalDamageWaiverAmount, service, tracingService);
                    UpdateQuote(quote.Id, extendedAmount,rentalTotalTaxAmount, effectiveDiscountAmount, detailedAmount, chargeAmount, damageWaiverAmount, environmentFee, cardProcessingFee, damageWaiverTaxAmount, totalDamageWaiverAmount, service, tracingService);
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: ProcessLogic >> " + ex.ToString());
                throw ex;
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

        private decimal GetSumOfTaxAmount(Guid quoteId, IOrganizationService service, ITracingService tracingService)
        {
            // Sum tax amount of all Quote line items for the same Quote
            try
            {
                decimal totalTaxAmount = 0;
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' aggregate='true'>
                                  <entity name='avpx_quotelineitem'>
                                    <attribute name='avpx_taxamount' aggregate='sum' alias='totalTaxAmountSum'/>
                                    <filter type='and'>
                                      <condition attribute='avpx_quote' operator='eq' value='{0}' />
                                      <condition attribute='statecode' operator='eq' value='0' />
                                      <condition attribute='avpx_type' operator='eq' value='783090000' />
                                    </filter>
                                  </entity>
                                </fetch>", quoteId);

                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    Entity entity = entityCollection.Entities[0];

                    // Extract the value from the AliasedValue and handle the Money type
                    if (entity.Contains("totalTaxAmountSum") && entity["totalTaxAmountSum"] is AliasedValue aliasedValue)
                    {
                        if (aliasedValue.Value is Money moneyValue)
                        {
                            totalTaxAmount = moneyValue.Value;
                        }
                    }
                }

                tracingService.Trace("totalTaxAmount: " + totalTaxAmount);
                return totalTaxAmount;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetSumOfTaxAmount >> " + ex.ToString());
                throw;
            }
        }


        private decimal GetSumOfExtendedAmount(Guid quoteId, IOrganizationService service, ITracingService tracingService)
        {
            //Sum quantity of all Quote line items for same Quote
            try
            {
                decimal extendedamount = 0;
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' aggregate='true'>
									  <entity name='avpx_quotelineitem'>
									    <attribute name='avpx_discountedamount' aggregate='sum' alias='totalextendedamountSum'/>
									    <filter type='and'>
									      <condition attribute='avpx_quote' operator='eq' value='{0}' />
									      <condition attribute='statecode' operator='eq' value='0' />
									    </filter>
									  </entity>
									</fetch>", quoteId);

                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    Entity entity = entityCollection.Entities[0];
                    extendedamount = entity.Contains("totalextendedamountSum") ? ((Money)((AliasedValue)entity["totalextendedamountSum"]).Value != null ? ((Money)((AliasedValue)entity["totalextendedamountSum"]).Value).Value : 0) : 0;
                }

                return extendedamount;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetSumOfExtendedAmount >> " + ex.ToString());
                throw ex;
            }
        }
        private decimal GetSumOfAttributeFromQuoteLineItem(Guid quoteId, string atributeName, int lineItemType, IOrganizationService service, ITracingService tracingService)
        {
            //Sum quantity of all Quote line items for same Quote
            try
            {
                decimal sum = 0;
                string filterCondition = lineItemType > 0 ? string.Format(@"<condition attribute='avpx_type' operator='eq' value='{0}' />", lineItemType) : string.Empty;
                /*(Start) Modified By Pratik Telaviya on 28-April-23 to fix the issue of */
                /*Earlier we were using fetch xml and aggregate function to sum the taxamount but we found that it returns the sum in base currency only*/
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
									  <entity name='avpx_quotelineitem'>
									    <attribute name='{0}'/>
									    <filter type='and'>
									      <condition attribute='avpx_quote' operator='eq' value='{1}' />
									      <condition attribute='statecode' operator='eq' value='0' />
                                          {2}
									    </filter>
									  </entity>
									</fetch>", atributeName, quoteId, filterCondition);


                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    foreach (Entity quoteLineItem in entityCollection.Entities)
                    {
                        sum += (quoteLineItem.Contains(atributeName) && ((Money)(quoteLineItem[atributeName])) != null) ? ((Money)(quoteLineItem[atributeName])).Value : 0;
                    }
                    //Entity entity = entityCollection.Entities[0];
                    //taxamount = entity.Contains("totaltaxamountSum") ? ((Money)((AliasedValue)entity["totaltaxamountSum"]).Value != null ? ((Money)((AliasedValue)entity["totaltaxamountSum"]).Value).Value : 0) : 0;
                }

                return sum;
                /*(End) Modified By Pratik Telaviya on 28-April-23 to fix the issue of */
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetSumOfAttributeFromQuoteLineItem >> " + ex.ToString());
                throw ex;
            }
        }
        private decimal GetSumOfEffectiveDiscountAmount(Guid quoteId, IOrganizationService service, ITracingService tracingService)
        {
            //Sum effectivediscountamount of all Quote line items for same Quote
            try
            {
                decimal effectivediscountamount = 0;
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' aggregate='true'>
									  <entity name='avpx_quotelineitem'>
									    <attribute name='avpx_effectivediscountamount' aggregate='sum' alias='totaleffectivediscountamountSum'/>
									    <filter type='and'>
									      <condition attribute='avpx_quote' operator='eq' value='{0}' />
									      <condition attribute='statecode' operator='eq' value='0' />
									    </filter>
									  </entity>
									</fetch>", quoteId);

                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    Entity entity = entityCollection.Entities[0];
                    effectivediscountamount = entity.Contains("totaleffectivediscountamountSum") ? ((Money)((AliasedValue)entity["totaleffectivediscountamountSum"]).Value != null ? ((Money)((AliasedValue)entity["totaleffectivediscountamountSum"]).Value).Value : 0) : 0;
                }

                return effectivediscountamount;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetSumOfEffectiveDiscountAmount >> " + ex.ToString());
                throw ex;
            }
        }
        private decimal GetSumOfAmount(Guid quoteId, IOrganizationService service, ITracingService tracingService)
        {
            //Sumamount of all Quote line items for same Quote
            try
            {
                decimal amount = 0;
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' aggregate='true'>
									  <entity name='avpx_quotelineitem'>
									    <attribute name='avpx_amount' aggregate='sum' alias='totalamountSum'/>
									    <filter type='and'>
									      <condition attribute='avpx_quote' operator='eq' value='{0}' />
									      <condition attribute='statecode' operator='eq' value='0' />
									      <condition attribute='avpx_type' operator='eq' value='783090000' />
									    </filter>
									  </entity>
									</fetch>", quoteId);

                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    Entity entity = entityCollection.Entities[0];
                    amount = entity.Contains("totalamountSum") ? ((Money)((AliasedValue)entity["totalamountSum"]).Value != null ? ((Money)((AliasedValue)entity["totalamountSum"]).Value).Value : 0) : 0;
                }

                return amount;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetSumOfAmount >> " + ex.ToString());
                throw ex;
            }
        }
        private decimal GetSumOfCharges(Guid quoteId, IOrganizationService service, ITracingService tracingService)
        {
            //Sum charges of all Quote line items for same Quote
            try
            {
                decimal charges = 0;
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' aggregate='true'>
									  <entity name='avpx_quotelineitem'>
									    <attribute name='avpx_amount' aggregate='sum' alias='totalCharges'/>
									    <filter type='and'>
									      <condition attribute='avpx_quote' operator='eq' value='{0}' />
									      <condition attribute='statecode' operator='eq' value='0' />
									      <condition attribute='avpx_type' operator='eq' value='783090001' />
									    </filter>
									  </entity>
									</fetch>", quoteId);

                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    Entity entity = entityCollection.Entities[0];
                    charges = entity.Contains("totalCharges") ? ((Money)((AliasedValue)entity["totalCharges"]).Value != null ? ((Money)((AliasedValue)entity["totalCharges"]).Value).Value : 0) : 0;
                }

                return charges;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetSumOfCharges >> " + ex.ToString());
                throw ex;
            }
        }
        private void UpdateQuote(Guid quoteId, decimal extendedAmount, decimal taxAmount, decimal effectiveDiscountAmount, decimal detailedAmount, decimal chargeAmount, decimal damageWaiverAmount,decimal environmentFee,decimal cardProcessingFee, decimal damageWaiverTaxAmount,decimal totalDamageWaiverAmount, IOrganizationService service, ITracingService tracingService)
        {
            try
            {
                Entity quoteupdate = new Entity("avpx_quote", quoteId);
                quoteupdate["avpx_detailedamount"] = new Money(detailedAmount);
                quoteupdate["avpx_totalmanualdiscount"] = new Money(effectiveDiscountAmount);
                decimal environmentFeeAmount = (detailedAmount - effectiveDiscountAmount) * environmentFee / 100;
                quoteupdate["avpx_subtotalamount"] = new Money(extendedAmount + damageWaiverAmount + environmentFeeAmount);
                quoteupdate["avpx_totaltax"] = new Money(taxAmount);
                decimal cardProcessingFeeAmount = (extendedAmount + damageWaiverAmount + environmentFeeAmount) * cardProcessingFee / 100;
                tracingService.Trace($"{environmentFeeAmount},{cardProcessingFeeAmount} ");
                tracingService.Trace("cardProcessingFeeAmount = "+ cardProcessingFeeAmount);
               
                quoteupdate["avpx_othercharges"] = new Money(chargeAmount);
                quoteupdate["avpx_damagewaiveramount"] = new Money(damageWaiverAmount);
                quoteupdate["avpx_damagewaivertaxamount"] = new Money(damageWaiverTaxAmount);
                quoteupdate["avpx_damagewaivereligibleamount"] = new Money(totalDamageWaiverAmount);
                quoteupdate["avpx_totalamount"] = new Money(extendedAmount + damageWaiverAmount + environmentFeeAmount + taxAmount + cardProcessingFeeAmount);
                service.Update(quoteupdate);
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: UpdateQuote >> " + ex.ToString());
                throw ex;
            }
        }

        private decimal GetSumOfDamageWaiverAmountFromQuoteLineItem(Guid quoteId, string atributeName, IOrganizationService service, ITracingService tracingService)
        {
            //Sum quantity of all Quote line items for same Quote
            try
            {
                decimal sum = 0;
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
									  <entity name='avpx_quotelineitem'>
									    <attribute name='{0}'/>
									    <filter type='and'>
									      <condition attribute='avpx_quote' operator='eq' value='{1}' />
									      <condition attribute='statecode' operator='eq' value='0' />
                                          <condition attribute='avpx_damagewaiverapplicable' operator='eq' value='1' />
                                           <condition attribute='avpx_type' operator='eq' value='783090000' />
                                        </filter >
                                      </entity>
									</fetch>", atributeName, quoteId);


                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    foreach (Entity quoteLineItem in entityCollection.Entities)
                    {
                        sum += (quoteLineItem.Contains(atributeName) && ((Money)(quoteLineItem[atributeName])) != null) ? ((Money)(quoteLineItem[atributeName])).Value : 0;
                    }
                }

                return sum;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetSumOfDamageWaiverAmountFromQuoteLineItem >> " + ex.ToString());
                throw ex;
            }
        }
    }
}
