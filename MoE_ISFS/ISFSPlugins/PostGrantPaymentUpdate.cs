using System;
using System.Collections.Generic;
using System.ServiceModel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace MoE_ISFS.ISFSPlugins
{

    /// <summary>
    /// PostGrantPaymentUpdate Plugin.
    /// Fires when the following attributes are updated:
    /// isfs_paymentdetailcreatedon
    /// </summary>    
    public class PostGrantPaymentUpdate: PluginBase
    {

        private class PaymentDetail
        {
            public decimal PaymentAmount { get; set; }
            public List<EntityReference> Schools { get; set; }
            public List<Entity> SchoolDisbursements { get; set; }

            public PaymentDetail(decimal paymentAmount)
            {
                PaymentAmount = paymentAmount;
                Schools = new List<EntityReference>();
                SchoolDisbursements = new List<Entity>();
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PostGrantPaymentUpdate"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public PostGrantPaymentUpdate(string unsecure, string secure)
            : base(typeof(PostGrantPaymentUpdate))
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

            CreateAuthorityDisbursement(localContext, postImageEntity);
        }

        private void CreateAuthorityDisbursement(LocalPluginContext localContext, Entity postImageEntity)
        {
            if (postImageEntity != null && postImageEntity.LogicalName == "isfs_grantpayment")
            {
                var grantProgram = postImageEntity.GetAttributeValue<EntityReference>("isfs_grantprogram");
                if (grantProgram == null) return;

                var schoolYear = postImageEntity.GetAttributeValue<EntityReference>("isfs_schoolyear");

                // Get Funding Schedules under this Grant Payment
                var queryFundingSchedules = new QueryExpression("isfs_fundingschedule");
                queryFundingSchedules.ColumnSet.AllColumns = false;
                queryFundingSchedules.Criteria.AddCondition("isfs_grantpayment", ConditionOperator.Equal, postImageEntity.Id);
                var fundingSchedules = localContext.OrganizationService.RetrieveMultiple(queryFundingSchedules);
                if (fundingSchedules.Entities.Count == 0) return;

                // Delete all existing Authority Disbursement records
                var queryAuthorityPaymentDetail = new QueryExpression("isfs_authoritydisbursement");
                queryAuthorityPaymentDetail.ColumnSet.AllColumns = false;
                queryAuthorityPaymentDetail.Criteria.AddCondition("isfs_grantpayment", ConditionOperator.Equal, postImageEntity.Id);
                var paymentDetails = localContext.OrganizationService.RetrieveMultiple(queryAuthorityPaymentDetail);
                foreach (var paymentDetail in paymentDetails.Entities)
                {
                    localContext.OrganizationService.Delete(paymentDetail.LogicalName, paymentDetail.Id);
                }

                Dictionary<EntityReference, PaymentDetail> authorityPaymentDetails = new Dictionary<EntityReference, PaymentDetail>();

                foreach (var fundingSchedule in fundingSchedules.Entities)
                {
                    // Aggregate School Disbursements
                    var querySchoolDisbursement = new QueryExpression("isfs_schooldisbursement");
                    querySchoolDisbursement.ColumnSet.AllColumns = true;
                    querySchoolDisbursement.Criteria.AddCondition("isfs_grantprogram", ConditionOperator.Equal, grantProgram.Id);
                    querySchoolDisbursement.Criteria.AddCondition("isfs_fundingschedule", ConditionOperator.Equal, fundingSchedule.Id);

                    var schoolDisbursements = localContext.OrganizationService.RetrieveMultiple(querySchoolDisbursement);

                    foreach (var schoolDisbursement in schoolDisbursements.Entities)
                    {
                        var authority = schoolDisbursement.GetAttributeValue<EntityReference>("isfs_authority");
                        if (authority == null) continue;

                        if (!authorityPaymentDetails.ContainsKey(authority))
                        {
                            authorityPaymentDetails.Add(authority, new PaymentDetail(0));
                        }

                        var disbursementAmount = schoolDisbursement.GetAttributeValue<Money>("isfs_disbursementamount");
                        if (disbursementAmount == null) continue;

                        var school = schoolDisbursement.GetAttributeValue<EntityReference>("isfs_school");
                        if (school == null) continue;

                        var paymentDetail = authorityPaymentDetails[authority];
                        paymentDetail.PaymentAmount += disbursementAmount.Value;
                        paymentDetail.SchoolDisbursements.Add(schoolDisbursement);
                        if (!authorityPaymentDetails[authority].Schools.Contains(school))
                        {
                            authorityPaymentDetails[authority].Schools.Add(school);
                        }
                    }
                }

                foreach (KeyValuePair<EntityReference, PaymentDetail> paymentDetail in authorityPaymentDetails)
                {
                    // Get Authority Number
                    var authorityEntity = localContext.OrganizationService.Retrieve(paymentDetail.Key.LogicalName, paymentDetail.Key.Id, new ColumnSet("isfs_authorityno"));

                    // Create Authority Disbursement record
                    Entity authorityDisbursement = new Entity("isfs_authoritydisbursement");

                    authorityDisbursement["isfs_name"] = string.Format("{0} {1}", authorityEntity.GetAttributeValue<string>("isfs_authorityno"), postImageEntity.GetAttributeValue<string>("isfs_name"));
                    authorityDisbursement["isfs_authority"] = paymentDetail.Key;
                    authorityDisbursement["isfs_grantpayment"] = postImageEntity.ToEntityReference();
                    authorityDisbursement["isfs_paymentamount"] = new Money(paymentDetail.Value.PaymentAmount);
                    authorityDisbursement["isfs_schoolcount"] = paymentDetail.Value.Schools.Count;
                    authorityDisbursement["isfs_schoolyear"] = schoolYear;

                    authorityDisbursement.Id = localContext.OrganizationService.Create(authorityDisbursement);

                    // Update School Disbursement record
                    foreach (var schoolDisbursement in paymentDetail.Value.SchoolDisbursements)
                    {
                        schoolDisbursement["isfs_authoritydisbursement"] = authorityDisbursement.ToEntityReference();
                        localContext.OrganizationService.Update(schoolDisbursement);
                    }
                }

                // Calculate Roll-up fields
                CalculateRollupFieldRequest calcuateRollupRequest = new CalculateRollupFieldRequest
                {
                    Target = postImageEntity.ToEntityReference(),
                    FieldName = "isfs_totalpaymentamount"
                };
                localContext.OrganizationService.Execute(calcuateRollupRequest);

                calcuateRollupRequest.FieldName = "isfs_authoritycount";
                localContext.OrganizationService.Execute(calcuateRollupRequest);

                calcuateRollupRequest.FieldName = "isfs_schoolcount";
                localContext.OrganizationService.Execute(calcuateRollupRequest);
            }
        }
    }
}
