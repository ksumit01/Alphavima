using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
/// <summary>
/// This plugin gets triggered on create and change of amount,Discounted Amount, Tax Amount, Tax Amount and on delete of SOLI and it rolls up the amount related fields into Sales Order.
/// <summary>
namespace AV.Prexa.Plugins.SalesOrderLineItems
{
    public class RollupSalesOrderLineItemsToSalesOrder : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            try
            {
                if (context.PrimaryEntityName.ToLower() != "avpx_salesorderlineitem") return;

                if (context.MessageName.ToUpper() == "CREATE" || context.MessageName.ToUpper() == "UPDATE" || context.MessageName.ToUpper() == "DELETE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if ((context.InputParameters.Contains("Target")) && ((context.InputParameters["Target"] is Entity) || context.InputParameters["Target"] is EntityReference))
                    {
                        ProcessLogic(entity, context, service, tracingService);
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
                EntityReference salesOrder = null;
                string message = context.MessageName.ToUpper();

                if (message == "CREATE" || message == "UPDATE")
                {
                    targetEntity = service.Retrieve("avpx_salesorderlineitem", context.PrimaryEntityId, new ColumnSet("avpx_salesorder"));
                    salesOrder = targetEntity.GetAttributeValue<EntityReference>("avpx_salesorder");
                }
                else if (message == "DELETE")
                {
                    Entity DeleteSLI = (Entity)context.PreEntityImages["Pre-Image"];
                    salesOrder = (EntityReference)DeleteSLI.Attributes["avpx_salesorder"];
                }

                if (salesOrder != null)
                {
                    decimal detailedAmount = GetSumOfAmount(salesOrder.Id, service, tracingService);
                    decimal chargeAmount = GetSumOfCharges(salesOrder.Id, service, tracingService);
                    decimal effectiveDiscountAmount = GetSumOfEffectiveDiscountAmounts(salesOrder.Id, service, tracingService);
                    decimal subTotalAmount = detailedAmount + chargeAmount - effectiveDiscountAmount;
                    decimal taxAmount = GetSumOfTaxAmount(salesOrder.Id, service, tracingService);

                    UpdateSalesOrder(salesOrder.Id, detailedAmount, chargeAmount, effectiveDiscountAmount, subTotalAmount, taxAmount, service, tracingService);
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: ProcessLogic >> " + ex.ToString());
                throw ex;
            }
        }
        private decimal GetSumOfAmount(Guid salesOrderId, IOrganizationService service, ITracingService tracingService)
        {
            //Sum amount of all order line items for same Sales Order
            try
            {
                decimal amount = 0;
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
									  <entity name='avpx_salesorderlineitem'>
									    <attribute name='avpx_amount'/>
									    <filter type='and'>
									      <condition attribute='avpx_salesorder' operator='eq' value='{0}' />
									      <condition attribute='statecode' operator='eq' value='0' />
                                          <condition attribute='avpx_type' operator='eq' value='783090000' />
									    </filter>
									  </entity>
									</fetch>", salesOrderId);
                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    //Entity entity = entityCollection.Entities[0];
                    //amount = entity.Contains("totalamountSum") ? ((Money)((AliasedValue)entity["totalamountSum"]).Value != null ? ((Money)((AliasedValue)entity["totalamountSum"]).Value).Value : 0) : 0;

                    ////Update subtotal field in sales order entity.
                    //Entity salesorderupdate = new Entity("avpx_salesorders", salesOrderId);
                    //salesorderupdate["avpx_subtotalamount"] = amount;
                    //service.Update(salesorderupdate);
                    foreach (Entity entity in entityCollection.Entities)
                    {
                        if (entity.GetAttributeValue<Money>("avpx_amount") != null)
                        {
                            amount += entity.GetAttributeValue<Money>("avpx_amount").Value;
                        }
                    }
                }

                return amount;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetSumOfAmount >> " + ex.ToString());
                throw ex;
            }
        }
        private decimal GetSumOfCharges(Guid salesOrderId, IOrganizationService service, ITracingService tracingService)
        {
            //Sum amount of all order line items for same Sales Order
            try
            {
                decimal charges = 0;
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
									  <entity name='avpx_salesorderlineitem'>
									    <attribute name='avpx_amount'/>
									    <filter type='and'>
									      <condition attribute='avpx_salesorder' operator='eq' value='{0}' />
									      <condition attribute='statecode' operator='eq' value='0' />
                                          <condition attribute='avpx_type' operator='eq' value='783090001' />
									    </filter>
									  </entity>
									</fetch>", salesOrderId);

                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    //Entity entity = entityCollection.Entities[0];
                    //charges = entity.Contains("totalCharges") ? ((Money)((AliasedValue)entity["totalCharges"]).Value != null ? ((Money)((AliasedValue)entity["totalCharges"]).Value).Value : 0) : 0;
                    foreach (Entity entity in entityCollection.Entities)
                    {
                        if (entity.GetAttributeValue<Money>("avpx_amount") != null)
                        {
                            charges += entity.GetAttributeValue<Money>("avpx_amount").Value;
                        }
                    }
                }

                return charges;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetSumOfCharges >> " + ex.ToString());
                throw ex;
            }
        }
        private decimal GetSumOfEffectiveDiscountAmounts(Guid salesOrderId, IOrganizationService service, ITracingService tracingService)
        {
            //Sum discounted amount of all sales order line items for same Sales Order
            try
            {
                decimal effectivediscountamount = 0;
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
									  <entity name='avpx_salesorderlineitem'>
									    <attribute name='avpx_effectivediscountamount'/>
									    <filter type='and'>
									      <condition attribute='avpx_salesorder' operator='eq' value='{0}' />
									      <condition attribute='statecode' operator='eq' value='0' />
									    </filter>
									  </entity>
									</fetch>", salesOrderId);

                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    foreach (Entity entity in entityCollection.Entities)
                    {
                        if (entity.GetAttributeValue<Money>("avpx_effectivediscountamount") != null)
                        {
                            effectivediscountamount += entity.GetAttributeValue<Money>("avpx_effectivediscountamount").Value;
                        }
                    }
                }

                return effectivediscountamount;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetSumOfEffectiveDiscountAmounts >> " + ex.ToString());
                throw ex;
            }
        }
        private decimal GetSumOfTaxAmount(Guid salesOrderId, IOrganizationService service, ITracingService tracingService)
        {
            //Sum tax amount of all sales order line items for same sales order
            try
            {
                decimal taxAmount = 0;
                string fetchXML = string.Format(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
									  <entity name='avpx_salesorderlineitem'>
									    <attribute name='avpx_taxamount'/>
									    <filter type='and'>
									      <condition attribute='avpx_salesorder' operator='eq' value='{0}' />
									      <condition attribute='statecode' operator='eq' value='0' />
									    </filter>
									  </entity>
									</fetch>", salesOrderId);

                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

                if (entityCollection != null && entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                {
                    foreach (Entity entity in entityCollection.Entities)
                    {
                        if (entity.GetAttributeValue<Money>("avpx_taxamount") != null)
                        {
                            taxAmount += entity.GetAttributeValue<Money>("avpx_taxamount").Value;
                        }
                    }
                }

                return taxAmount;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: GetSumOfTaxAmount >> " + ex.ToString());
                throw ex;
            }
        }
        private void UpdateSalesOrder(Guid salesOrderId, decimal detailedAmount, decimal chargeAmount, decimal discountAmount, decimal subtotalAmount, decimal taxAmount, IOrganizationService service, ITracingService tracingService)
        {
            try
            {
                Entity salesOrderToUpdate = new Entity("avpx_salesorders", salesOrderId);
                salesOrderToUpdate["avpx_productsamount"] = new Money(detailedAmount);
                salesOrderToUpdate["avpx_additionalcharges"] = new Money(chargeAmount);
                salesOrderToUpdate["avpx_discountamount"] = new Money(discountAmount);
                salesOrderToUpdate["avpx_subtotalamount"] = new Money(subtotalAmount);
                salesOrderToUpdate["avpx_totaltax"] = new Money(taxAmount);
                service.Update(salesOrderToUpdate);
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: UpdateSalesOrder >> " + ex.ToString());
                throw ex;
            }
        }
    }
}
