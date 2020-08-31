using System;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace MoE_ISFS.ISFSPlugins
{

    /// <summary>
    /// PostGrantPaymentCreate Plugin.
    /// </summary>    
    public class PostGrantPaymentCreate: PluginBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PostGrantPaymentCreate"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public PostGrantPaymentCreate(string unsecure, string secure)
            : base(typeof(PostGrantPaymentCreate))
        {
            
           // TODO: Implement your custom configuration handling.
        }


        /// <summary>
        /// Main entry point for he business logic that the plug-in is to execute.
        /// </summary>
        /// <param name="localContext">The <see cref="LocalPluginContext"/> which contains the
        /// <see cref="IPluginExecutionContext"/>,
        /// <see cref="IOrganizationService"/>
        /// and <see cref="ITracingService"/>
        /// </param>
        /// <remarks>
        /// For improved performance, Microsoft Dynamics 365 caches plug-in instances.
        /// The plug-in's Execute method should be written to be stateless as the constructor
        /// is not called for every invocation of the plug-in. Also, multiple system threads
        /// could execute the plug-in at the same time. All per invocation state information
        /// is stored in the context. This means that you should not use global variables in plug-ins.
        /// </remarks>
        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new InvalidPluginExecutionException("localContext");
            }

            IPluginExecutionContext context = localContext.PluginExecutionContext;

            Entity postImageEntity = (context.PostEntityImages != null && context.PostEntityImages.Contains(this.postImageAlias)) ? context.PostEntityImages[this.postImageAlias] : null;
            if (postImageEntity == null)
            {
                throw new InvalidPluginExecutionException("Missing PostImage.\r\n");
            }

            AddFundingSchedule(localContext, postImageEntity);
        }

        private void AddFundingSchedule(LocalPluginContext localContext, Entity grantPayment)
        {
            var grantProgram = grantPayment.GetAttributeValue<EntityReference>("isfs_grantprogram");
            var disbursementDate = grantPayment.GetAttributeValue<DateTime>("isfs_disbursementdate");

            var queryFundingSchedule = new QueryExpression("isfs_fundingschedule");
            queryFundingSchedule.ColumnSet.AllColumns = false;
            queryFundingSchedule.Criteria.AddCondition("isfs_grantprogram", ConditionOperator.Equal, grantProgram.Id);
            queryFundingSchedule.Criteria.AddCondition("isfs_disbursementdate", ConditionOperator.Equal, disbursementDate);
            queryFundingSchedule.Criteria.AddCondition("isfs_grantpayment", ConditionOperator.Null);
            queryFundingSchedule.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

            var fundingSchedules = localContext.OrganizationService.RetrieveMultiple(queryFundingSchedule);
            foreach (var fundingSchedule in fundingSchedules.Entities)
            {
                fundingSchedule["isfs_grantpayment"] = grantPayment.ToEntityReference();
                localContext.OrganizationService.Update(fundingSchedule);
            }
        }
    }
}
