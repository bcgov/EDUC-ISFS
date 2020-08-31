using System;
using System.Globalization;
using System.ServiceModel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace MoE_ISFS.ISFSPlugins
{

    /// <summary>
    /// PostEnrolmentDetailUpdate Plugin.
    /// Fires when the following attributes are updated:
    /// isfs_adjustedenrolmentnumber
    /// </summary>    
    public class PostEnrolmentDetailUpdate: PluginBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PostEnrolmentDetailUpdate"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public PostEnrolmentDetailUpdate(string unsecure, string secure)
            : base(typeof(PostEnrolmentDetailUpdate))
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

            UpdateSchoolDisbursement(localContext, postImageEntity);
        }

        private void UpdateSchoolDisbursement(LocalPluginContext localContext, Entity postImageEntity)
        {
            if (postImageEntity != null && postImageEntity.LogicalName == "isfs_isenrolmentdetail")
            {
                var adjustedEnrolmentNumber = postImageEntity.GetAttributeValue<decimal?>("isfs_adjustedenrolmentnumber");

                // List Active School Disbursement Details 
                var querySchoolDisbursementDetail = new QueryExpression("isfs_schooldisbursementdetail");
                querySchoolDisbursementDetail.ColumnSet.AllColumns = true;
                querySchoolDisbursementDetail.Criteria.AddCondition("isfs_isenrolmentdetail", ConditionOperator.Equal, postImageEntity.Id);
                querySchoolDisbursementDetail.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

                var schoolDisbursementDetails = localContext.OrganizationService.RetrieveMultiple(querySchoolDisbursementDetail);
                foreach (var schoolDisbursementDetail in schoolDisbursementDetails.Entities)
                {
                    schoolDisbursementDetail["isfs_enrolmentnumber"] = (adjustedEnrolmentNumber != null) ? adjustedEnrolmentNumber.Value : postImageEntity.GetAttributeValue<decimal>("isfs_enrolmentnumber");
                    localContext.OrganizationService.Update(schoolDisbursementDetail);

                    var schoolDisbursement = postImageEntity.GetAttributeValue<EntityReference>("isfs_schooldisbursement");

                    // Calculate Roll-up fields
                    //CalculateRollupFieldRequest calcuateRollupRequest = new CalculateRollupFieldRequest
                    //{
                    //    Target = schoolDisbursement,
                    //    FieldName = "isfs_disbursementamount"
                    //};

                    //localContext.OrganizationService.Execute(calcuateRollupRequest);
                }
            }
        }
    }
}
