
using System;
using System.Collections.Generic;
using System.ServiceModel;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using System.Text;

namespace MoE_ISFS.ISFSPlugins
{

    /// <summary>
    /// PostDisbursementCreationCreate Plugin.
    /// </summary>    
    public class PostDisbursementCreationCreate: PluginBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PostDisbursementCreationCreate"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public PostDisbursementCreationCreate(string unsecure, string secure)
            : base(typeof(PostDisbursementCreationCreate))
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

            CreateSchoolDisbursement(localContext, postImageEntity);
        }
        private void CreateSchoolDisbursement(LocalPluginContext localContext, Entity disbursementCreation)
        {
            if (disbursementCreation != null && disbursementCreation.LogicalName == "isfs_disbursementcreation")
            {
                var log = new StringBuilder("");

                int selectedSchoolCount = 0;
                int restrictedSchoolCount = 0;
                int removedSchoolDisbursementCount = 0;
                int skippedSchoolDisbursementCount = 0;
                int createdSchoolDisbursmentntCount = 0;
                var errorLogs = new List<string>();

                try
                {
                    log.Append(string.Format("{0}: Info - Creating School Disbursements started ...\r\n", DateTime.Now));

                    var fundingSchedule = disbursementCreation.GetAttributeValue<EntityReference>("isfs_fundingschedule");
                    if (fundingSchedule == null) return;

                    var fundingScheduleEntity = localContext.OrganizationService.Retrieve(fundingSchedule.LogicalName, fundingSchedule.Id, new ColumnSet(true));
                    var disbursementDate = fundingScheduleEntity.GetAttributeValue<DateTime>("isfs_disbursementdate");
                    var scheduleState = fundingScheduleEntity.GetAttributeValue<OptionSetValue>("statecode");
                    if (scheduleState.Value == 1) // Inactive - Paid/Canceled
                    {
                        log.Append(string.Format("{0}: Info - Cannot create School Disbursement for Paid/Canceled Funding Schedule.\r\n", DateTime.Now.ToLocalTime()));
                        return;
                    }

                    var schoolYear = fundingScheduleEntity.GetAttributeValue<EntityReference>("isfs_schoolyear");

                    var grantProgram = disbursementCreation.GetAttributeValue<EntityReference>("isfs_grantprogram");
                    if (grantProgram == null) return;

                    // List Funding Restrictions
                    var queryFundingRestriction = new QueryExpression("isfs_fundingrestriction");
                    queryFundingRestriction.ColumnSet.AllColumns = true;
                    queryFundingRestriction.Criteria.AddCondition("isfs_grantprogram", ConditionOperator.Equal, grantProgram.Id);
                    queryFundingRestriction.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // Active
                    var fundingRestrictions = localContext.OrganizationService.RetrieveMultiple(queryFundingRestriction);
                    var restrictedSchools = new List<EntityReference>();
                    foreach (var fundingRestriction in fundingRestrictions.Entities)
                    {
                        var startDate = fundingRestriction.GetAttributeValue<DateTime>("isfs_startdate");
                        var endDate = fundingRestriction.GetAttributeValue<DateTime>("isfs_enddate");
                        if (startDate <= disbursementDate && disbursementDate <= endDate)
                        {
                            restrictedSchools.Add(fundingRestriction.GetAttributeValue<EntityReference>("isfs_school"));
                        }
                    }

                    // Select Enrolment Collection
                    var selectedEnrolmentCollection = disbursementCreation.GetAttributeValue<EntityReference>("isfs_enrolmentcollection");
                    var enrolmentCollection = localContext.OrganizationService.Retrieve(selectedEnrolmentCollection.LogicalName, selectedEnrolmentCollection.Id, new ColumnSet(true));
                    var collectionType = enrolmentCollection.GetAttributeValue<OptionSetValue>("isfs_collectiontype");
                    var isESAudit = collectionType.Value == 746910002;
                    var enrolmentSchoolYear = enrolmentCollection.GetAttributeValue<EntityReference>("isfs_schoolyear");
                    var isCrossSchoolYear = enrolmentSchoolYear.Id != schoolYear.Id;

                    // Select Funding Rate
                    var selectedFundingRateId = disbursementCreation.GetAttributeValue<EntityReference>("isfs_fundingrate");
                    var selectedFundingRate = localContext.OrganizationService.Retrieve(selectedFundingRateId.LogicalName, selectedFundingRateId.Id, new ColumnSet(true));

                    // List Schools
                    var querySchool = new QueryExpression("edu_school");
                    querySchool.ColumnSet.AllColumns = true;
                    var schools = localContext.OrganizationService.RetrieveMultiple(querySchool);
                    var schoolEntities = new Dictionary<EntityReference, Entity>();
                    foreach (var school in schools.Entities)
                    {
                        schoolEntities.Add(school.ToEntityReference(), school);
                    }

                    // List Enrolments
                    var queryEnrolment = new QueryExpression("isfs_isenrolment");
                    queryEnrolment.ColumnSet.AllColumns = true;
                    queryEnrolment.Criteria.AddCondition("isfs_enrolmentcollection", ConditionOperator.Equal, enrolmentCollection.Id);
                    var enrolments = localContext.OrganizationService.RetrieveMultiple(queryEnrolment);
                    var enrolmentEntities = new Dictionary<Guid, Entity>();
                    foreach (var enrolment in enrolments.Entities)
                    {
                        enrolmentEntities.Add(enrolment.Id, enrolment);
                    }

                    var selectedEnrolmentIds = disbursementCreation.GetAttributeValue<string>("isfs_selectedenrolmentids");
                    if (string.IsNullOrWhiteSpace(selectedEnrolmentIds)) return;

                    var enrolmentIds = selectedEnrolmentIds.Split(';').ToList();
                    selectedSchoolCount = enrolmentIds.Count();
                    log.Append(string.Format("{0}: Info - Selected {1} school enrolment(s). \r\n", DateTime.Now, selectedSchoolCount));

                    var replacedExistingSchoolDisbursement = disbursementCreation.GetAttributeValue<bool>("isfs_replaceexistingdisbursements");
                    // Existing School Disbursements
                    var querySchoolDisbursement = new QueryExpression("isfs_schooldisbursement");
                    querySchoolDisbursement.ColumnSet.AllColumns = true;
                    querySchoolDisbursement.Criteria.AddCondition("isfs_grantprogram", ConditionOperator.Equal, grantProgram.Id);
                    querySchoolDisbursement.Criteria.AddCondition("isfs_fundingschedule", ConditionOperator.Equal, fundingSchedule.Id);
                    var schoolDisbursements = localContext.OrganizationService.RetrieveMultiple(querySchoolDisbursement);
                    foreach (var schoolDisbursement in schoolDisbursements.Entities)
                    {
                        var enrolmentId = schoolDisbursement.GetAttributeValue<EntityReference>("isfs_isenrolment");
                        if (enrolmentId != null && enrolmentIds.Contains(enrolmentId.Id.ToString()))
                        {
                            var school = enrolmentEntities[enrolmentId.Id].GetAttributeValue<EntityReference>("isfs_school");
                            if (restrictedSchools.Contains(school))
                            {
                                removedSchoolDisbursementCount++;
                                localContext.OrganizationService.Delete(schoolDisbursement.LogicalName, schoolDisbursement.Id);
                                log.Append(string.Format("{0}: Info - Existing School Disbursement has been removed because of restricted funding for {1}.\r\n",
                                                            DateTime.Now, schoolDisbursement.GetAttributeValue<EntityReference>("isfs_school").Name));
                                continue;
                            }

                            if (replacedExistingSchoolDisbursement == true)
                            {
                                removedSchoolDisbursementCount++;
                                localContext.OrganizationService.Delete(schoolDisbursement.LogicalName, schoolDisbursement.Id);
                                log.Append(string.Format("{0}: Info - Existing School Disbursement has been removed for {1}.\r\n",
                                                            DateTime.Now, schoolDisbursement.GetAttributeValue<EntityReference>("isfs_school").Name));
                            }
                            else
                            {
                                skippedSchoolDisbursementCount++;
                                enrolmentIds.Remove(enrolmentId.Id.ToString());
                                log.Append(string.Format("{0}: Info - Existing School Disbursement has been kept for {1}.\r\n",
                                                             DateTime.Now, schoolDisbursement.GetAttributeValue<EntityReference>("isfs_school").Name));
                            }
                        }
                    }

                    foreach (var enrolmentIdString in enrolmentIds.ToArray())
                    {
                        var enrolmentId = new Guid(enrolmentIdString);
                        if (enrolmentEntities.Keys.Contains(enrolmentId))
                        {
                            var school = enrolmentEntities[enrolmentId].GetAttributeValue<EntityReference>("isfs_school");
                            if (restrictedSchools.Contains(school))
                            {
                                restrictedSchoolCount++;
                                enrolmentIds.Remove(enrolmentIdString);
                                log.Append(string.Format("{0}: Info - Disbursement creation is restricted for {1}.\r\n",
                                                             DateTime.Now, school.Name));
                            }
                        }
                    }

                    if (enrolmentIds.Count() == 0)
                    {
                        return;
                    }

                    // List IS Enrolment Details
                    var enrolmentDetailState = 0;
                    var queryEnrolmentDetail = new QueryExpression("isfs_isenrolmentdetail");
                    queryEnrolmentDetail.ColumnSet.AllColumns = true;
                    queryEnrolmentDetail.AddOrder("isfs_name", OrderType.Ascending);
                    queryEnrolmentDetail.Criteria.AddCondition("statecode", ConditionOperator.Equal, enrolmentDetailState);

                    var linkISEnrolment = queryEnrolmentDetail.AddLink("isfs_isenrolment", "isfs_isenrolment", "isfs_isenrolmentid");
                    linkISEnrolment.EntityAlias = "af";
                    linkISEnrolment.Columns.AddColumns("isfs_school");
                    linkISEnrolment.LinkCriteria.AddCondition("isfs_enrolmentcollection", ConditionOperator.Equal, selectedEnrolmentCollection.Id);

                    var linkESGroupSubCategory = queryEnrolmentDetail.AddLink("isfs_esgroupsubcategory", "isfs_esgroupsubcategory", "isfs_esgroupsubcategoryid", JoinOperator.LeftOuter);
                    linkESGroupSubCategory.EntityAlias = "ae";
                    linkESGroupSubCategory.Columns.AddColumns("isfs_esgroup", "isfs_isschoolsubcategory", "isfs_rollupcategory");

                    var enrolmentDetails = localContext.OrganizationService.RetrieveMultiple(queryEnrolmentDetail);

                    var processedEnrolmentDetials = new List<Entity>();

                    // List Funding Rate Details
                    var queryFundingRateDetail = new QueryExpression("isfs_fundingratedetail");
                    queryFundingRateDetail.ColumnSet.AllColumns = true;
                    queryFundingRateDetail.AddOrder("isfs_isschoolsubcategory", OrderType.Descending);
                    queryFundingRateDetail.Criteria.AddCondition("isfs_fundingrate", ConditionOperator.Equal, selectedFundingRate.Id);
                    var fundingRateDetails = localContext.OrganizationService.RetrieveMultiple(queryFundingRateDetail);

                    // List ES Group
                    var queryESGroup = new QueryExpression("isfs_esgroup");
                    queryESGroup.ColumnSet.AllColumns = true;
                    queryESGroup.Criteria.AddCondition("isfs_schoolyear", ConditionOperator.Equal, schoolYear.Id);
                    var ESGroups = localContext.OrganizationService.RetrieveMultiple(queryESGroup);
                    var previousESGroups = new Dictionary<Guid, EntityReference>();
                    var currentESGroups = new Dictionary<Guid, EntityReference>();
                    foreach (var esGroup in ESGroups.Entities)
                    {
                        var previousESGroup = esGroup.GetAttributeValue<EntityReference>("isfs_previoussyesgroup");
                        if (previousESGroup != null)
                        {
                            previousESGroups.Add(esGroup.Id, previousESGroup);
                            currentESGroups.Add(previousESGroup.Id, esGroup.ToEntityReference());
                        }
                    }

                    // List ES Group SubCategory
                    var queryESGroupSubCategory = new QueryExpression("isfs_esgroupsubcategory");
                    queryESGroupSubCategory.ColumnSet.AllColumns = true;
                    queryESGroupSubCategory.Criteria.AddCondition("isfs_schoolyear", ConditionOperator.Equal, schoolYear.Id);
                    var ESGroupSubCategories = localContext.OrganizationService.RetrieveMultiple(queryESGroupSubCategory);
                    var previousESGroupSubCategories = new Dictionary<Guid, EntityReference>();
                    var currentESGroupSubCategories = new Dictionary<Guid, EntityReference>();
                    foreach (var esGroupSubCategory in ESGroupSubCategories.Entities)
                    {
                        var previousESGroupSubCategory = esGroupSubCategory.GetAttributeValue<EntityReference>("isfs_previoussyesgroupsubcategory");
                        if (previousESGroupSubCategory != null)
                        {
                            previousESGroupSubCategories.Add(esGroupSubCategory.Id, previousESGroupSubCategory);
                            currentESGroupSubCategories.Add(previousESGroupSubCategory.Id, esGroupSubCategory.ToEntityReference());
                        }
                    }

                    // List Funding Schedule Details
                    var queryFundingScheduleDetail = new QueryExpression("isfs_fundingscheduledetail");
                    queryFundingScheduleDetail.ColumnSet.AllColumns = true;
                    queryFundingScheduleDetail.Criteria.AddCondition("isfs_fundingschedule", ConditionOperator.Equal, fundingSchedule.Id);
                    var fundingScheduleDetails = localContext.OrganizationService.RetrieveMultiple(queryFundingScheduleDetail);

                    var createdSchoolDisbursements = new Dictionary<string, Entity>();

                    foreach (var fundingScheduleDetail in fundingScheduleDetails.Entities)
                    {
                        var ytdDisbursementPercentage = fundingScheduleDetail.GetAttributeValue<decimal>("isfs_ytddisbursementpercentage");
                        if (ytdDisbursementPercentage == 0) continue;

                        var scheduleDetailSchoolCategory = fundingScheduleDetail.GetAttributeValue<EntityReference>("isfs_schoolsubcategory");
                        var scheduleDetailESGroup = fundingScheduleDetail.GetAttributeValue<EntityReference>("isfs_esgroup");
                        var scheduleDetailESSubGroup = fundingScheduleDetail.GetAttributeValue<EntityReference>("isfs_esgroupsubcategory");
                        var changedEnrolmentNumberSetting = fundingScheduleDetail.GetAttributeValue<bool>("isfs_changedenrolmentnumber");
                        var enrolmentNetIncreaseSetting = fundingScheduleDetail.GetAttributeValue<bool>("isfs_enrolmentnetincrease");

                        if (isCrossSchoolYear && scheduleDetailESGroup != null)
                        {
                            if (previousESGroups.Keys.Contains(scheduleDetailESGroup.Id))
                            {
                                scheduleDetailESGroup = previousESGroups[scheduleDetailESGroup.Id];
                            }
                            else
                            {
                                errorLogs.Add(string.Format("{0}: Error - The Previous SY ES Group is missing for ES Group: {1}.\r\n", DateTime.Now, scheduleDetailESGroup.Name));
                                return;
                            }
                        }

                        if (isCrossSchoolYear && scheduleDetailESSubGroup != null)
                        {
                            if (previousESGroupSubCategories.Keys.Contains(scheduleDetailESSubGroup.Id))
                            {
                                scheduleDetailESSubGroup = previousESGroupSubCategories[scheduleDetailESSubGroup.Id];
                            }
                            else
                            {
                                errorLogs.Add(string.Format("{0}: Error - The Previous SY ES Group Sub-Category is missing for ES Group Sub-Category: {1}.\r\n", DateTime.Now, scheduleDetailESSubGroup.Name));
                                return;
                            }
                        }

                        var fundingDetailRef = fundingScheduleDetail.GetAttributeValue<EntityReference>("isfs_fundingdetail");
                        Entity appliedFundingDetail = localContext.OrganizationService.Retrieve(fundingDetailRef.LogicalName, fundingDetailRef.Id, new ColumnSet(true));
                        if (appliedFundingDetail == null)
                        {
                            errorLogs.Add(string.Format("{0}: Error - No Funding Detail.\r\n", DateTime.Now));
                            return;
                        }

                        var sameSchoolSubCategoryOnly = appliedFundingDetail.GetAttributeValue<bool>("isfs_schoolsubcategoryonly");

                        foreach (var enrolmentDetail in enrolmentDetails.Entities)
                        {
                            if (processedEnrolmentDetials.Contains(enrolmentDetail)) continue;

                            var enrolmentId = enrolmentDetail.GetAttributeValue<EntityReference>("isfs_isenrolment").Id;

                            if (!enrolmentIds.Contains(enrolmentId.ToString())) continue;

                            var enrolmentDetailSchoolSubCategory = enrolmentDetail.GetAttributeValue<EntityReference>("isfs_schoolsubcategory");
                            var enrolmentDetailESGroup = enrolmentDetail.GetAttributeValue<EntityReference>("isfs_esgroup");
                            EntityReference enrolmentDetailESSubGroup = isESAudit ? enrolmentDetail.GetAttributeValue<EntityReference>("isfs_mappedesgroupsubcategory") : enrolmentDetail.GetAttributeValue<EntityReference>("isfs_esgroupsubcategory");

                            if (scheduleDetailSchoolCategory.Id != enrolmentDetailSchoolSubCategory.Id ||
                                scheduleDetailESGroup.Id != enrolmentDetailESGroup.Id)
                                continue;
                            else
                            {
                                if (scheduleDetailESSubGroup != null)
                                {
                                    if (enrolmentDetailESSubGroup == null)
                                        continue;
                                    else if (scheduleDetailESSubGroup.Id != enrolmentDetailESSubGroup.Id)
                                        continue;
                                }
                            }


                            //if (esGroupDisbursements.ContainsKey(enrolmentDetailESGroup))
                            {
                                var school = (EntityReference)enrolmentDetail.GetAttributeValue<AliasedValue>("af.isfs_school").Value;

                                if (sameSchoolSubCategoryOnly == true)
                                {
                                    var schoolSubCategory = schoolEntities[school].GetAttributeValue<EntityReference>("isfs_schoolsubcategory");
                                    if (schoolSubCategory == null || schoolSubCategory.Id != enrolmentDetailSchoolSubCategory.Id)
                                        continue;
                                }

                                var rateType = appliedFundingDetail.GetAttributeValue<OptionSetValue>("isfs_fundingratetype");
                                var groupType = appliedFundingDetail.GetAttributeValue<OptionSetValue>("isfs_fundinggrouptype");


                                // Get Funding Rate Detail
                                Entity appliedFundingRateDetail = null;
                                var isESGroupMatched = false;
                                var isESGroupSubCategoryMatched = false;
                                var isSchoolSCategoryMatched = false;
                                var isRateTypeMatched = false;
                                var isFundingGroupMatched = false;

                                foreach (var fundingRateDetail in fundingRateDetails.Entities)
                                {
                                    isESGroupMatched = false;
                                    isESGroupSubCategoryMatched = false;
                                    isSchoolSCategoryMatched = false;
                                    isRateTypeMatched = false;
                                    isFundingGroupMatched = false;

                                    isESGroupMatched = fundingRateDetail.GetAttributeValue<EntityReference>("isfs_esgroup").Equals(enrolmentDetailESGroup);

                                    var fundingRateESSubCategory = fundingRateDetail.GetAttributeValue<EntityReference>("isfs_esgroupsubcategory");
                                    var fundingDetailESSubCategory = appliedFundingDetail.GetAttributeValue<EntityReference>("isfs_esgroupsubcategory");
                                    if (isCrossSchoolYear && fundingDetailESSubCategory != null)
                                        fundingDetailESSubCategory = previousESGroupSubCategories[fundingDetailESSubCategory.Id];
                                    isESGroupSubCategoryMatched = fundingDetailESSubCategory == null ? true : fundingDetailESSubCategory.Equals(fundingRateESSubCategory);

                                    var fundingSchoolSubCategory = fundingRateDetail.GetAttributeValue<EntityReference>("isfs_isschoolsubcategory");
                                    isSchoolSCategoryMatched = fundingSchoolSubCategory == null ? true : fundingSchoolSubCategory.Equals(enrolmentDetailSchoolSubCategory);

                                    var fundingDistrict = fundingRateDetail.GetAttributeValue<EntityReference>("isfs_schooldistrict");
                                    var fundingSchool = fundingRateDetail.GetAttributeValue<EntityReference>("isfs_school");
                                    switch (rateType.Value)
                                    {
                                        case 746910000: // Flat
                                            {
                                                if (fundingDistrict == null && fundingSchool == null)
                                                    isRateTypeMatched = true;
                                            }
                                            break;
                                        case 746910001: // District 
                                            {
                                                if (fundingDistrict != null && fundingDistrict.Equals(schoolEntities[school].GetAttributeValue<EntityReference>("edu_schooldistrict")))
                                                    isRateTypeMatched = true;
                                            }
                                            break;
                                        case 746910002: // School 
                                            {
                                                if (fundingSchool != null && fundingSchool.Equals(school))
                                                    isRateTypeMatched = true;
                                            }
                                            break;
                                        default: break;
                                    }

                                    var fundingRateGroup = fundingRateDetail.GetAttributeValue<EntityReference>("isfs_fundinggroup");

                                    switch (groupType.Value)
                                    {
                                        case 746910000: // N/A
                                            {
                                                if (fundingRateGroup == null)
                                                    isFundingGroupMatched = true;
                                            }
                                            break;
                                        case 746910001: // School
                                            {
                                                var fundingGroup = enrolmentEntities[enrolmentId].GetAttributeValue<EntityReference>("isfs_fundinggroup");
                                                if (fundingGroup != null && fundingGroup.Equals(fundingRateGroup))
                                                    isFundingGroupMatched = true;
                                            }
                                            break;
                                        case 746910002: // By ES Group
                                            {
                                                var fundingGroup = enrolmentDetail.GetAttributeValue<EntityReference>("isfs_fundinggroup");
                                                if (fundingGroup != null && fundingGroup.Equals(fundingRateGroup))
                                                    isFundingGroupMatched = true;
                                            }
                                            break;
                                        case 746910003: // By ES Group Sub-Category
                                            {
                                                var fundingGroup = enrolmentDetail.GetAttributeValue<EntityReference>("isfs_fundinggroup");
                                                if (fundingGroup != null && fundingGroup.Equals(fundingRateGroup) &&
                                                    enrolmentDetailESSubGroup != null && enrolmentDetailESSubGroup.Equals(fundingRateESSubCategory))
                                                    isFundingGroupMatched = true;
                                            }
                                            break;
                                        default: break;
                                    }

                                    if (isESGroupMatched && isESGroupSubCategoryMatched && isSchoolSCategoryMatched && isFundingGroupMatched && isRateTypeMatched)
                                    {
                                        appliedFundingRateDetail = fundingRateDetail;
                                        break;
                                    }
                                }

                                if (appliedFundingRateDetail == null)
                                {
                                    //						errorLogs.Add(string.Format("{0}: Error - Couldn't find Funding Rate Detail for {1}: ESGroupMatched: {2}, SchoolCategoryMatched: {3}, FundingGroupMatched:{4}, RateTypeMatched: {5}.\r\n",
                                    //													DateTime.Now, enrolmentDetail["isfs_name"], isESGroupMatched, isSchoolSCategoryMatched, isFundingGroupMatched, isRateTypeMatched));
                                    errorLogs.Add(string.Format("{0}: Error - Couldn't find Funding Rate Detail for {1}. \r\n",
                                        DateTime.Now, enrolmentDetail["isfs_name"]));

                                    continue;
                                }

                                // Get School Disbursement
                                var disbursementName = string.Format("{0} {1}", schoolEntities[school].GetAttributeValue<string>("edu_mincode"), fundingScheduleEntity.GetAttributeValue<string>("isfs_name"));
                                if (!createdSchoolDisbursements.ContainsKey(disbursementName))
                                {
                                    // Create School Disbursement
                                    Entity newSchoolDisbursement = new Entity("isfs_schooldisbursement");

                                    newSchoolDisbursement["isfs_disbursementname"] = disbursementName;
                                    newSchoolDisbursement["isfs_authority"] = schoolEntities[school].GetAttributeValue<EntityReference>("isfs_authority");
                                    newSchoolDisbursement["isfs_district"] = schoolEntities[school].GetAttributeValue<EntityReference>("edu_schooldistrict");
                                    newSchoolDisbursement["isfs_school"] = school;
                                    newSchoolDisbursement["isfs_fundingschedule"] = fundingSchedule;
                                    newSchoolDisbursement["isfs_fundingrate"] = selectedFundingRate.ToEntityReference();
                                    newSchoolDisbursement["isfs_grantprogram"] = grantProgram;
                                    newSchoolDisbursement["isfs_isenrolment"] = enrolmentDetail.GetAttributeValue<EntityReference>("isfs_isenrolment");
                                    newSchoolDisbursement["isfs_schoolyear"] = schoolYear;

                                    newSchoolDisbursement["isfs_disbursementcreation"] = disbursementCreation.ToEntityReference();

                                    newSchoolDisbursement.Id = localContext.OrganizationService.Create(newSchoolDisbursement);

                                    createdSchoolDisbursements.Add(disbursementName, newSchoolDisbursement);

                                    createdSchoolDisbursmentntCount++;

                                    log.Append(string.Format("{0}: Info - New School Disbursement: {1} has been created for {2}.\r\n", DateTime.Now, disbursementName, school.Name));

                                }

                                Entity schoolDisbursement = createdSchoolDisbursements[disbursementName];


                                // Create School Disbursement Detail

                                Entity schoolDisbursementDetail = new Entity("isfs_schooldisbursementdetail");

                                // Inputs

                                var enrolmentNumber = GetEnrolmentNumber(enrolmentDetail, changedEnrolmentNumberSetting, enrolmentNetIncreaseSetting);
                                if (isESAudit == true)
                                    enrolmentNumber = enrolmentNumber + GetAggregateEnrolmentNumber(enrolmentDetail, enrolmentDetailESSubGroup, enrolmentDetails, school, changedEnrolmentNumberSetting, enrolmentNetIncreaseSetting, processedEnrolmentDetials);

                                var rollupESGroupSubCategory = appliedFundingDetail.GetAttributeValue<bool>("isfs_rollupesgroupsubcategory");
                                if (rollupESGroupSubCategory == true)
                                    enrolmentNumber = enrolmentNumber + GetRollupEnrolmentNumber(appliedFundingDetail, enrolmentDetails, school, changedEnrolmentNumberSetting, enrolmentNetIncreaseSetting);

                                var adjustedEnrolmentRatio = fundingScheduleDetail.GetAttributeValue<decimal?>("isfs_adjustenrolmentratio");
                                if (adjustedEnrolmentRatio.HasValue)
                                    enrolmentNumber = enrolmentNumber * adjustedEnrolmentRatio.Value;

                                var fundingRate = appliedFundingRateDetail.GetAttributeValue<Money>("isfs_isschoolrate");
                                var groupPercentage = appliedFundingRateDetail.GetAttributeValue<decimal?>("isfs_grouppercentage");

                                Entity paidSchoolDisbursementDetail = null;

                                var newDisbursement = fundingScheduleDetail.GetAttributeValue<bool>("isfs_newdisbursement");
                                if (isCrossSchoolYear && newDisbursement == false)
                                {
                                    errorLogs.Add(string.Format("{0}: Error - Cross School Year Disbursement has to be New Disbursements for Funding Schedule Detail: {1}. \r\n",
                                        DateTime.Now, fundingScheduleDetail["isfs_name"]));
                                    return;
                                }

                                var previousFundingScheduleDetail = fundingScheduleDetail.GetAttributeValue<EntityReference>("isfs_previousfundingscheduledetail");
                                if (newDisbursement == false)
                                {
                                    if (previousFundingScheduleDetail == null)
                                    {
                                        errorLogs.Add(string.Format("{0}: Error - Previous Funding Schedule Detail is missing for {1}. \r\n",
                                        DateTime.Now, fundingScheduleDetail["isfs_name"]));
                                        continue;
                                    }

                                    var Paid = 2;
                                    var querySchoolDisburementDetail = new QueryExpression("isfs_schooldisbursementdetail");
                                    querySchoolDisburementDetail.TopCount = 1;
                                    querySchoolDisburementDetail.ColumnSet.AllColumns = true;
                                    querySchoolDisburementDetail.AddOrder("isfs_disbursementdate", OrderType.Descending);
                                    querySchoolDisburementDetail.Criteria.AddCondition("isfs_fundingscheduledetail", ConditionOperator.Equal, previousFundingScheduleDetail.Id);
                                    querySchoolDisburementDetail.Criteria.AddCondition("isfs_isschoolsubcategory", ConditionOperator.Equal, scheduleDetailSchoolCategory.Id);
                                    querySchoolDisburementDetail.Criteria.AddCondition("isfs_esgroup", ConditionOperator.Equal, scheduleDetailESGroup.Id);
                                    if (enrolmentDetailESSubGroup != null)
                                        querySchoolDisburementDetail.Criteria.AddCondition("isfs_esgroupsubcategory", ConditionOperator.Equal, enrolmentDetailESSubGroup.Id);
                                    querySchoolDisburementDetail.Criteria.AddCondition("statuscode", ConditionOperator.Equal, Paid);
                                    querySchoolDisburementDetail.Criteria.AddCondition("isfs_disbursementdate", ConditionOperator.OnOrBefore, fundingScheduleEntity.GetAttributeValue<DateTime>("isfs_disbursementdate"));

                                    var linkSchoolDisbursement = querySchoolDisburementDetail.AddLink("isfs_schooldisbursement", "isfs_schooldisbursement", "isfs_schooldisbursementid");
                                    linkSchoolDisbursement.EntityAlias = "ad";
                                    linkSchoolDisbursement.LinkCriteria.AddCondition("isfs_grantprogram", ConditionOperator.Equal, grantProgram.Id);
                                    linkSchoolDisbursement.LinkCriteria.AddCondition("isfs_school", ConditionOperator.Equal, school.Id);

                                    var paidSchoolDisbursementDetails = localContext.OrganizationService.RetrieveMultiple(querySchoolDisburementDetail);

                                    if (paidSchoolDisbursementDetails != null && paidSchoolDisbursementDetails.Entities.Count == 1)
                                    {
                                        paidSchoolDisbursementDetail = paidSchoolDisbursementDetails[0];
                                    }
                                }

                                if (enrolmentNumber == 0 && paidSchoolDisbursementDetail == null)
                                {
                                    continue;
                                }

                                var previousYTDDisbursement = new Money(0);
                                var previousYTDPayment = new Money(0);
                                var previousYTDBalance = new Money(0);

                                if (paidSchoolDisbursementDetail != null)
                                {
                                    previousYTDDisbursement = paidSchoolDisbursementDetail.GetAttributeValue<Money>("isfs_ytddisbursementamount");
                                    previousYTDPayment = paidSchoolDisbursementDetail.GetAttributeValue<Money>("isfs_ytdpaymentamount");
                                    previousYTDBalance = paidSchoolDisbursementDetail.GetAttributeValue<Money>("isfs_ytdbalanceamount");
                                }

                                // Calculations
                                var ytdDisbursement = ytdDisbursementPercentage * enrolmentNumber * fundingRate.Value;

                                var disbursement = ytdDisbursement - previousYTDDisbursement.Value;

                                var adjustBalance = fundingScheduleDetail.GetAttributeValue<bool>("isfs_adjustbalance");
                                var adjustedDisbursement = adjustBalance ? disbursement + previousYTDBalance.Value : disbursement;

                                var allowNegativePayment = fundingScheduleDetail.GetAttributeValue<bool>("isfs_allownegativepayment");
                                var payment = allowNegativePayment ? adjustedDisbursement : adjustedDisbursement > 0 ? adjustedDisbursement : 0;
                                var ytdPayment = payment + previousYTDPayment.Value;

                                var ytdBalance = ytdDisbursement - ytdPayment;

                                // Outputs
                                if (enrolmentDetailESSubGroup != null)
                                    schoolDisbursementDetail["isfs_name"] = string.Format("{0} {1} Grant", disbursementName, enrolmentDetailESSubGroup.Name);
                                else
                                    schoolDisbursementDetail["isfs_name"] = string.Format("{0} {1} Grant", disbursementName, enrolmentDetailESGroup.Name);

                                schoolDisbursementDetail["isfs_ytddisbursementpercentage"] = ytdDisbursementPercentage;

                                schoolDisbursementDetail["isfs_enrolmentnumber"] = enrolmentNumber;
                                schoolDisbursementDetail["isfs_rollupesgroupsubcategory"] = rollupESGroupSubCategory;

                                schoolDisbursementDetail["isfs_fundingrate"] = fundingRate;

                                if (groupPercentage.HasValue)
                                    schoolDisbursementDetail["isfs_grouppercentage"] = groupPercentage.Value;
                                if (adjustedEnrolmentRatio.HasValue)
                                    schoolDisbursementDetail["isfs_adjustenrolmentratio"] = adjustedEnrolmentRatio.Value;

                                schoolDisbursementDetail["isfs_ytddisbursementamount"] = new Money(ytdDisbursement);
                                schoolDisbursementDetail["isfs_previousytddisbursement"] = paidSchoolDisbursementDetail == null ? null : previousYTDDisbursement;
                                schoolDisbursementDetail["isfs_disbursementamount"] = new Money(adjustedDisbursement);
                                schoolDisbursementDetail["isfs_balanceadjusted"] = adjustBalance;
                                schoolDisbursementDetail["isfs_paymentamount"] = new Money(payment);
                                schoolDisbursementDetail["isfs_ytdpaymentamount"] = new Money(ytdPayment);
                                schoolDisbursementDetail["isfs_previousytdpayment"] = paidSchoolDisbursementDetail == null ? null : previousYTDPayment;
                                schoolDisbursementDetail["isfs_previousytdbalance"] = paidSchoolDisbursementDetail == null ? null : previousYTDBalance;

                                schoolDisbursementDetail["isfs_ytdbalanceamount"] = new Money(ytdBalance);
                                schoolDisbursementDetail["isfs_isschoolsubcategory"] = enrolmentDetailSchoolSubCategory;
                                schoolDisbursementDetail["isfs_enrolmentnumbertype"] = enrolmentDetail.GetAttributeValue<OptionSetValue>("isfs_enrolmentnumbertype");
                                schoolDisbursementDetail["isfs_esgroup"] = isCrossSchoolYear == false ? enrolmentDetailESGroup : currentESGroups[enrolmentDetailESGroup.Id];
                                if (enrolmentDetailESSubGroup != null)
                                    schoolDisbursementDetail["isfs_esgroupsubcategory"] = isCrossSchoolYear == false ? enrolmentDetailESSubGroup : currentESGroupSubCategories[enrolmentDetailESSubGroup.Id];
                                schoolDisbursementDetail["isfs_fundinggroup"] = appliedFundingRateDetail.GetAttributeValue<EntityReference>("isfs_fundinggroup");

                                schoolDisbursementDetail["isfs_disbursementdate"] = fundingScheduleEntity.GetAttributeValue<DateTime>("isfs_disbursementdate");

                                schoolDisbursementDetail["isfs_schoolyear"] = schoolYear;

                                // Related details
                                schoolDisbursementDetail["isfs_previousdisbursementdetail"] = paidSchoolDisbursementDetail?.ToEntityReference();
                                schoolDisbursementDetail["isfs_fundingratedetail"] = appliedFundingRateDetail.ToEntityReference();
                                schoolDisbursementDetail["isfs_fundingscheduledetail"] = fundingScheduleDetail.ToEntityReference();
                                schoolDisbursementDetail["isfs_isenrolmentdetail"] = enrolmentDetail.ToEntityReference();
                                schoolDisbursementDetail["isfs_fundingdetail"] = appliedFundingDetail.ToEntityReference();
                                schoolDisbursementDetail["isfs_schooldisbursement"] = schoolDisbursement.ToEntityReference();

                                schoolDisbursementDetail.Id = localContext.OrganizationService.Create(schoolDisbursementDetail);
                            }
                        }
                    }

                    foreach (var schoolDisbursement in createdSchoolDisbursements.Values)
                    {
                        // Calculate Roll-up fields
                        CalculateRollupFieldRequest calcuateRollupRequest = new CalculateRollupFieldRequest
                        {
                            Target = schoolDisbursement.ToEntityReference(),
                            FieldName = "isfs_disbursementamount"
                        };

                        localContext.OrganizationService.Execute(calcuateRollupRequest);

                        calcuateRollupRequest.FieldName = "isfs_ytddisbursementamount";
                        localContext.OrganizationService.Execute(calcuateRollupRequest);

                        calcuateRollupRequest.FieldName = "isfs_ytdbalance";
                        localContext.OrganizationService.Execute(calcuateRollupRequest);
                    }
                }
                finally
                {
                    if (errorLogs.Count() == 0)
                        log.Append(string.Format("{0}: Info - Creating School Disbursements completed.\r\n", DateTime.Now.ToLocalTime()));
                    else
                    {
                        log.Append(string.Format("{0}: Error - Creating School Disbursements completed with the following error(s):\r\n", DateTime.Now.ToLocalTime()));
                        foreach (var err in errorLogs)
                        {
                            log.Append(err);
                        }
                    }

                    log.Append(string.Format("{0}: Info - Selected Schools: {1}; Restricted Schools: {5}; Removed School Disbursements: {2}; Skipped School Disbursements: {3}; Created School Disbursements: {4}.\r\n",
                                    DateTime.Now, selectedSchoolCount, removedSchoolDisbursementCount, skippedSchoolDisbursementCount, createdSchoolDisbursmentntCount, restrictedSchoolCount));

                    disbursementCreation["isfs_creationcompletedon"] = DateTime.UtcNow;
                    disbursementCreation["isfs_creationlog"] = log.ToString();
                    localContext.OrganizationService.Update(disbursementCreation);
                }
            }
        }
        private decimal GetEnrolmentNumber(Entity enrolmentDetail, bool changedEnrolmentNumberSetting, bool enrolmentNetIncreaseSetting)
        {
            var adjustedEnrolmentNumber = enrolmentDetail.GetAttributeValue<decimal?>("isfs_adjustedenrolmentnumber");
            var enrolmentNumber = (adjustedEnrolmentNumber != null) ? adjustedEnrolmentNumber.Value : enrolmentDetail.GetAttributeValue<decimal>("isfs_enrolmentnumber");
            if (changedEnrolmentNumberSetting == true)
            {
                var changedEnrolmentNumber = enrolmentDetail.GetAttributeValue<decimal?>("isfs_enrollmentnumberdifferencefromprevious");
                enrolmentNumber = (changedEnrolmentNumber == null) ? 0 : changedEnrolmentNumber.Value;
                if (enrolmentNetIncreaseSetting == true && enrolmentNumber < 0)
                    enrolmentNumber = 0;
            }

            return enrolmentNumber;
        }
        private decimal GetRollupEnrolmentNumber(Entity appliedFundingDetail, EntityCollection enrolmentDetails, EntityReference school, bool changedEnrolmentNumberSetting, bool enrolmentNetIncreaseSetting)
        {
            decimal rollupEnrolmentNumber = 0;

            if (appliedFundingDetail != null)
            {
                var fundingDetailESSubCategory = appliedFundingDetail.GetAttributeValue<EntityReference>("isfs_esgroupsubcategory");

                foreach (var enrolmentDetail in enrolmentDetails.Entities)
                {
                    var schoolRef = (EntityReference)enrolmentDetail.GetAttributeValue<AliasedValue>("af.isfs_school").Value;
                    var rollupCategory = enrolmentDetail.GetAttributeValue<AliasedValue>("ae.isfs_rollupcategory");

                    if (schoolRef == null || rollupCategory == null) continue;

                    if (schoolRef.Id == school.Id && (rollupCategory.Value as EntityReference).Id == fundingDetailESSubCategory.Id)
                    {
                        rollupEnrolmentNumber = rollupEnrolmentNumber + GetEnrolmentNumber(enrolmentDetail, changedEnrolmentNumberSetting, enrolmentNetIncreaseSetting);
                    }
                }
            }

            return rollupEnrolmentNumber;
        }
        private decimal GetAggregateEnrolmentNumber(Entity firstEnrolmentDetail, EntityReference esGroupSubcategory, EntityCollection enrolmentDetails, EntityReference school, bool changedEnrolmentNumberSetting, bool enrolmentNetIncreaseSetting, List<Entity> processedEnrolmentDetials)
        {
            decimal aggregateEnrolmentNumber = 0;

            if (processedEnrolmentDetials.Contains(firstEnrolmentDetail))
                return aggregateEnrolmentNumber;
            else
                processedEnrolmentDetials.Add(firstEnrolmentDetail);

            foreach (var enrolmentDetail in enrolmentDetails.Entities)
            {
                if (processedEnrolmentDetials.Contains(enrolmentDetail)) continue;

                if (enrolmentDetail.Id == firstEnrolmentDetail.Id) continue;

                var schoolRef = (EntityReference)enrolmentDetail.GetAttributeValue<AliasedValue>("af.isfs_school").Value;
                var aggregateESGroupSubcategory = enrolmentDetail.GetAttributeValue<EntityReference>("isfs_mappedesgroupsubcategory");

                if (schoolRef == null || aggregateESGroupSubcategory == null) continue;

                if (schoolRef.Id == school.Id && aggregateESGroupSubcategory.Id == esGroupSubcategory.Id)
                {
                    aggregateEnrolmentNumber = aggregateEnrolmentNumber + GetEnrolmentNumber(enrolmentDetail, changedEnrolmentNumberSetting, enrolmentNetIncreaseSetting);
                    processedEnrolmentDetials.Add(enrolmentDetail);
                }
            }

            return aggregateEnrolmentNumber;
        }
    }
}
