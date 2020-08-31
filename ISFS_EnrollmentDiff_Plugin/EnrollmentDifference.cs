using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using System.Web.Configuration;
using System.Xml;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace MoE_ISFS.ISFSEnrollmentDiff
{
    public class EnrollmentDifference : PluginBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PostAuthorityPaymentUpdate"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public EnrollmentDifference(string unsecure, string secure)
            : base(typeof(EnrollmentDifference))
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
            localContext.Trace("EnrollmentDifference Plugin: Starting");

            if (localContext == null) throw new InvalidPluginExecutionException(OperationStatus.Failed, "Missing localContext");

            string entityName = localContext.PluginExecutionContext.PrimaryEntityName.ToLower();
            Guid collectionId;

            if (entityName.Equals("isfs_enrolmentcollection"))
            {
                EntityReference previousCollectionRef = null;

                if (!localContext.PluginExecutionContext.InputParameters.Contains("Target"))
                    throw new InvalidPluginExecutionException(OperationStatus.Failed, "Missing Target");

                localContext.Trace("EnrollmentDifference Plugin: " + localContext.PluginExecutionContext.MessageName.ToLower() + "|isfs_enrolmentcollection");

                if (localContext.PluginExecutionContext.MessageName.ToLower().Equals("delete"))
                {
                    EntityReference enrolmentRef = (EntityReference)localContext.PluginExecutionContext.InputParameters["Target"];

                    localContext.Trace("EnrollmentDifference Plugin: UpdateCurrentEnrolments()");

                    UpdateCurrentEnrolments(localContext, enrolmentRef);

                    return;
                }


                Entity collection = (Entity)localContext.PluginExecutionContext.InputParameters["Target"];
                collectionId = localContext.PluginExecutionContext.PrimaryEntityId;

                if (collection.Attributes.Contains("isfs_previouscollection"))
                {
                    previousCollectionRef = collection.GetAttributeValue<EntityReference>("isfs_previouscollection");

                    if (!collection.Attributes.Contains("isfs_schoolyear"))
                    {
                        collection = localContext.OrganizationService.Retrieve(collection.LogicalName, collection.Id, new ColumnSet(new string[] { "isfs_previouscollection", "isfs_schoolyear" }));
                    }

                    Entity previousCollection = null;
                    if (previousCollectionRef != null) previousCollection = localContext.OrganizationService.Retrieve(previousCollectionRef.LogicalName, previousCollectionRef.Id, new ColumnSet(new string[] { "isfs_schoolyear" }));

                    if (previousCollectionRef == null || collection.GetAttributeValue<EntityReference>("isfs_schoolyear").Id.Equals(previousCollection.GetAttributeValue<EntityReference>("isfs_schoolyear").Id))
                    {
                        SetCalculatingDeltas(localContext, collection.Id, true);

                        if (previousCollectionRef != null && !previousCollectionRef.Id.Equals(Guid.Empty))
                        {
                            localContext.Trace("EnrollmentDifference Plugin: FlagEnrolmentsForRecalc()");

                            FlagEnrolmentsForRecalc(localContext, collection.Id);
                        }
                        else
                        {
                            localContext.Trace("EnrollmentDifference Plugin: RemoveAllDifferences()");

                            FlagEnrolmentsForRecalc(localContext, collection.Id);
                            //RemoveAllDifferences(localContext, localContext.PluginExecutionContext.PrimaryEntityId);
                            //SetCalculatingDeltas(localContext, collectionId, false);
                        }
                    }
                    else
                    {
                        localContext.Trace("EnrollmentDifference Plugin: Collection and previous collection school years don't match, exiting plugin.");
                    }
                }
                else
                {
                    localContext.Trace("EnrollmentDifference Plugin: Missing InputParameter 'isfs_previouscollection', exiting plugin.");
                }
            }
            else if (entityName.Equals("isfs_isenrolmentdetail"))
            {
                Entity enrolmentDetail = null;
                bool isDelete = false;
                bool ignoreCreate = false;

                localContext.Trace("EnrollmentDifference Plugin: " + localContext.PluginExecutionContext.MessageName.ToLower() + "|isfs_isenrolmentdetail");

                if (localContext.PluginExecutionContext.MessageName.ToLower().Equals("delete"))
                {
                    isDelete = true;
                    EntityReference enrolmentDetailRef = (EntityReference)localContext.PluginExecutionContext.InputParameters["Target"];
                    enrolmentDetail = localContext.OrganizationService.Retrieve(enrolmentDetailRef.LogicalName, enrolmentDetailRef.Id,
                        new ColumnSet(new string[] { "isfs_isenrolmentdetailid", "isfs_enrolmentnumber", "isfs_adjustedenrolmentnumber" }));
                }
                else
                {
                    enrolmentDetail = (Entity)localContext.PluginExecutionContext.InputParameters["Target"];
                }

                if (enrolmentDetail.Attributes.Contains("isfs_ignorecreate") && enrolmentDetail.GetAttributeValue<bool>("isfs_ignorecreate") == true) 
                    ignoreCreate = true;

                if (!ignoreCreate && (isDelete || enrolmentDetail.Attributes.Contains("isfs_enrolmentnumber") || enrolmentDetail.Attributes.Contains("isfs_adjustedenrolmentnumber")))
                {
                    localContext.Trace("EnrollmentDifference Plugin: RecalculateEnrollmentDetailDifference()");
                    RecalculateEnrollmentDetailDifference(localContext, enrolmentDetail, isDelete);
                }
                else
                {
                    localContext.Trace("EnrollmentDifference Plugin: Missing InputParameter 'isfs_enrolmentnumber' or 'isfs_adjustedenrolmentnumber', or parameter isfs_ignorecreate was specified. Exiting plugin.");
                }
            }
            else if (entityName.Equals("isfs_isenrolment"))
            {
                if (localContext.PluginExecutionContext.MessageName.ToLower().Equals("update"))
                {
                    localContext.Trace("EnrollmentDifference Plugin: update|isfs_isenrolment");

                    Entity enrolment = (Entity)localContext.PluginExecutionContext.InputParameters["Target"];

                    if (enrolment.Attributes.Contains("isfs_triggerdifferencecalc") && 
                        enrolment.GetAttributeValue<bool>("isfs_triggerdifferencecalc") == true)
                    {
                        Guid previousCollectionId = Guid.Empty;
                        if (enrolment.Attributes.Contains("isfs_previouscollection"))
                        {
                            previousCollectionId = enrolment.GetAttributeValue<EntityReference>("isfs_previouscollection").Id;
                        }
                        else
                        {
                            Entity collection = GetCollectionForEnrolment(localContext, enrolment.Id);
                            if (collection != null && collection.Attributes.Contains("isfs_previouscollection")) previousCollectionId = collection.GetAttributeValue<EntityReference>("isfs_previouscollection").Id;
                        }

                        if (!previousCollectionId.Equals(Guid.Empty))
                        {
                            localContext.Trace("EnrollmentDifference Plugin: RecalcEnrolment()");
                            RecalcEnrolment(localContext, enrolment);
                        }
                        else
                        {
                            localContext.Trace("EnrollmentDifference Plugin: RemoveEnrolmentDifferences()");
                            RemoveEnrolmentDifferences(localContext, enrolment);
                        }
                    }
                    else localContext.Trace("EnrollmentDifference Plugin: Missing InputParameter 'isfs_triggerdifferencecalc=true', exiting plugin.");
                }
                else if (localContext.PluginExecutionContext.MessageName.ToLower().Equals("delete"))
                {
                    localContext.Trace("EnrollmentDifference Plugin: delete|isfs_isenrolment");
                    EntityReference enrolmentRef = (EntityReference)localContext.PluginExecutionContext.InputParameters["Target"];

                    localContext.Trace("EnrollmentDifference Plugin: UpdateCurrentEnrolments()");
                    UpdateCurrentEnrolments(localContext, enrolmentRef);
                }
            }
            else throw new InvalidPluginExecutionException(OperationStatus.Failed, "Entity '" + entityName + "' not supported by EnrollmentDifference plugin.");
        }

        private Entity GetCollectionForEnrolment(LocalPluginContext localContext, Guid enrolmentId)
        {
            QueryExpression query = new QueryExpression("isfs_enrolmentcollection");
            query.TopCount = 1;
            query.ColumnSet.AddColumns(new string[] { "isfs_previouscollection" });

            LinkEntity enrolmentLink = query.AddLink("isfs_isenrolment", "isfs_enrolmentcollectionid", "isfs_enrolmentcollection", JoinOperator.Inner);
            enrolmentLink.LinkCriteria.AddCondition(new ConditionExpression("isfs_isenrolmentid", ConditionOperator.Equal, enrolmentId));

            return localContext.OrganizationService.RetrieveMultiple(query).Entities.FirstOrDefault();
        }

        private void RecalcEnrolment(LocalPluginContext localContext, Entity enrolment)
        {
            EntityReference collectionRef = null;
            int count = 0;

            EntityCollection enrolmentDetails = GetEnrolmentDetails(localContext, enrolment);

            if (enrolmentDetails != null && enrolmentDetails.Entities != null && enrolmentDetails.Entities.Count > 0)
            {
                ExecuteMultipleRequest executeMultipleRequest = null;
                decimal enrolmentNum;
                decimal previousEnrolmentNum;
                decimal enrolmentDelta;
                decimal currentDelta;
                bool hasCurrentDelta;

                Guid previousCollectionId = ((EntityReference)enrolmentDetails[0].GetAttributeValue<AliasedValue>("Col.isfs_previouscollection").Value).Id;
                EntityReference schoolRef = (EntityReference)enrolmentDetails[0].GetAttributeValue<AliasedValue>("Enr.isfs_school").Value;
                enrolment.Attributes["isfs_school"] = schoolRef;

                collectionRef = new EntityReference("isfs_enrolmentcollection", (Guid)enrolmentDetails[0].GetAttributeValue<AliasedValue>("Col.isfs_enrolmentcollectionid").Value);
                enrolment.Attributes["isfs_enrolmentcollectionid"] = collectionRef;

                EntityCollection previousEnrolmentDetails = GetPreviousEnrolmentDetails(localContext, previousCollectionId, enrolment);

                // for each enrolment detail
                foreach (Entity enrolmentDetail in enrolmentDetails.Entities)
                {
                    enrolmentNum = 0;
                    previousEnrolmentNum = 0;
                    enrolmentDelta = 0;
                    currentDelta = 0;
                    hasCurrentDelta = false;

                    Entity previousEnrolmentDetail = (from e in previousEnrolmentDetails.Entities
                        where
                            (!enrolmentDetail.Contains("isfs_schoolsubcategory") || e.GetAttributeValue<EntityReference>("isfs_schoolsubcategory").Id == enrolmentDetail.GetAttributeValue<EntityReference>("isfs_schoolsubcategory").Id) &&
                            (!enrolmentDetail.Contains("isfs_esgroup") || e.GetAttributeValue<EntityReference>("isfs_esgroup").Id == enrolmentDetail.GetAttributeValue<EntityReference>("isfs_esgroup").Id) &&
                            (!enrolmentDetail.Contains("isfs_esgroupsubcategory") || e.GetAttributeValue<EntityReference>("isfs_esgroupsubcategory").Id == enrolmentDetail.GetAttributeValue<EntityReference>("isfs_esgroupsubcategory").Id)
                        select e).FirstOrDefault();

                    if (previousEnrolmentDetail != null)
                    {
                        previousEnrolmentDetails.Entities.Remove(previousEnrolmentDetail);

                        // find the matching detail and calculate the difference
                        //Entity matchingEnrolmentDetail = FindMatchingEnrolmentDetail(localContext, enrolmentDetail, previousEnrolment.Id);

                        //if (matchingEnrolmentDetail != null)
                        //{
                        //hasPreviousEnrollment = true; // ISFS-740

                        if (previousEnrolmentDetail.Contains("isfs_adjustedenrolmentnumber"))
                            previousEnrolmentNum = previousEnrolmentDetail.GetAttributeValue<decimal>("isfs_adjustedenrolmentnumber");

                        if (previousEnrolmentNum == 0 && previousEnrolmentDetail.Contains("isfs_enrolmentnumber"))
                            previousEnrolmentNum = previousEnrolmentDetail.GetAttributeValue<decimal>("isfs_enrolmentnumber");
                        //}
                    }

                    if (enrolmentDetail.Contains("isfs_enrollmentnumberdifferencefromprevious"))
                    {
                        currentDelta = enrolmentDetail.GetAttributeValue<decimal>("isfs_enrollmentnumberdifferencefromprevious");
                        hasCurrentDelta = true;
                    }

                    if (enrolmentDetail.Contains("isfs_adjustedenrolmentnumber"))
                        enrolmentNum = enrolmentDetail.GetAttributeValue<decimal>("isfs_adjustedenrolmentnumber");

                    if (enrolmentNum == 0 && enrolmentDetail.Contains("isfs_enrolmentnumber"))
                        enrolmentNum = enrolmentDetail.GetAttributeValue<decimal>("isfs_enrolmentnumber");

                    enrolmentDelta = enrolmentNum - previousEnrolmentNum;

                    if ((enrolmentDelta != currentDelta) || !hasCurrentDelta)
                    {
                        count++;

                        if (executeMultipleRequest == null)
                        {
                            executeMultipleRequest = new ExecuteMultipleRequest()
                            {
                                // Assign settings that define execution behavior: continue on error, return responses.
                                Settings = new ExecuteMultipleSettings()
                                {
                                    ContinueOnError = true,
                                    ReturnResponses = false
                                },
                                // Create an empty organization request collection.
                                Requests = new OrganizationRequestCollection()
                            };
                        }

                        Entity updatedEnrolmentDetail = new Entity("isfs_isenrolmentdetail", enrolmentDetail.Id);

                        updatedEnrolmentDetail.Attributes.Add("isfs_enrollmentnumberdifferencefromprevious", enrolmentDelta);
                        updatedEnrolmentDetail.Attributes.Add("isfs_ignorecreate", false);

                        executeMultipleRequest.Requests.Add(new UpdateRequest() { Target = updatedEnrolmentDetail });
                    }
                }

                if (previousEnrolmentDetails != null && previousEnrolmentDetails.Entities != null && previousEnrolmentDetails.Entities.Count > 0)
                {
                    foreach (Entity previousEnrtDetail in previousEnrolmentDetails.Entities)
                    {
                        count++;

                        Entity currentEnrollment = CreateCurrentEnrolmentDetailFromPrevious(localContext, enrolment.Id, previousEnrtDetail);

                        executeMultipleRequest.Requests.Add(new CreateRequest() { Target = currentEnrollment });
                    }
                }

                if (count > 0)
                {
                    localContext.OrganizationService.Execute(executeMultipleRequest);
                } 
            }

            Entity updatedEnrolment = new Entity("isfs_isenrolment", enrolment.Id);
            updatedEnrolment.Attributes["isfs_triggerdifferencecalc"] = false;

            localContext.OrganizationService.Update(updatedEnrolment);

            // Check if other isfs_isenrolment in collection still recalculating
            if (collectionRef != null && !StillCalculating(localContext, collectionRef.Id))
            {
                SetCalculatingDeltas(localContext, collectionRef.Id, false);
            }
        }

        private Entity CreateCurrentEnrolmentDetailFromPrevious(LocalPluginContext localContext, Guid enrolmentId, Entity previousEnrolmentDetail)
        {
            decimal previousEnrolmentNum = 0;
            decimal enrolmentDelta = 0;

            if (previousEnrolmentDetail.Contains("isfs_adjustedenrolmentnumber"))
                previousEnrolmentNum = previousEnrolmentDetail.GetAttributeValue<decimal>("isfs_adjustedenrolmentnumber");

            if (previousEnrolmentNum == 0 && previousEnrolmentDetail.Contains("isfs_enrolmentnumber"))
                previousEnrolmentNum = previousEnrolmentDetail.GetAttributeValue<decimal>("isfs_enrolmentnumber");

            enrolmentDelta = 0 - previousEnrolmentNum;

            Entity currentEnrollment = new Entity("isfs_isenrolmentdetail");
            currentEnrollment.Attributes["isfs_isenrolment"] = new EntityReference("isfs_isenrolment", enrolmentId);
            currentEnrollment.Attributes["isfs_name"] = "Created by Enrollment Difference Calculation";
            currentEnrollment.Attributes["isfs_ignorecreate"] = true;
            currentEnrollment.Attributes["isfs_schoolyear"] = previousEnrolmentDetail.Attributes["isfs_schoolyear"];
            currentEnrollment.Attributes["isfs_enrolmentnumber"] = decimal.Parse("0");
            currentEnrollment.Attributes["isfs_enrollmentnumberdifferencefromprevious"] = enrolmentDelta;
            if (previousEnrolmentDetail.Attributes.Contains("isfs_schoolsubcategory")) currentEnrollment.Attributes["isfs_schoolsubcategory"] = previousEnrolmentDetail.Attributes["isfs_schoolsubcategory"];
            if (previousEnrolmentDetail.Attributes.Contains("isfs_esgroup")) currentEnrollment.Attributes["isfs_esgroup"] = previousEnrolmentDetail.Attributes["isfs_esgroup"];
            if (previousEnrolmentDetail.Attributes.Contains("isfs_esgroupsubcategory")) currentEnrollment.Attributes["isfs_esgroupsubcategory"] = previousEnrolmentDetail.Attributes["isfs_esgroupsubcategory"];
            if (previousEnrolmentDetail.Attributes.Contains("isfs_enrolmentnumbertype")) currentEnrollment.Attributes["isfs_enrolmentnumbertype"] = previousEnrolmentDetail.Attributes["isfs_enrolmentnumbertype"];
            if (previousEnrolmentDetail.Attributes.Contains("isfs_fundinggroup")) currentEnrollment.Attributes["isfs_fundinggroup"] = previousEnrolmentDetail.Attributes["isfs_fundinggroup"];
            if (previousEnrolmentDetail.Attributes.Contains("isfs_mappedesgroupsubcategory")) currentEnrollment.Attributes["isfs_mappedesgroupsubcategory"] = previousEnrolmentDetail.Attributes["isfs_mappedesgroupsubcategory"];

            return currentEnrollment;
        }

        private bool StillCalculating(LocalPluginContext localContext, Guid collectionId)
        {
            QueryExpression query = new QueryExpression("isfs_isenrolment");
            query.TopCount = 1;
            query.ColumnSet.AddColumns(new string[] { "isfs_isenrolmentid" });

            query.Criteria.AddCondition(new ConditionExpression("isfs_enrolmentcollection", ConditionOperator.Equal, collectionId));
            query.Criteria.AddCondition(new ConditionExpression("isfs_triggerdifferencecalc", ConditionOperator.Equal, true));

            Entity result = localContext.OrganizationService.RetrieveMultiple(query).Entities.FirstOrDefault();

            if (result == null)
                return false;
            else
                return true;
        }

        private void UpdateCurrentEnrolments(LocalPluginContext localContext, EntityReference entityRef)
        {
            int count = 0;
            EntityCollection previousEnrolmentDetails = GetPreviousEnrolments(localContext, entityRef);

            if (previousEnrolmentDetails != null && previousEnrolmentDetails.Entities != null && previousEnrolmentDetails.Entities.Count > 0)
            {
                ExecuteMultipleRequest executeMultipleRequest = null;

                foreach (Entity enrollmentDetail in previousEnrolmentDetails.Entities)
                {
                    count++;

                    if (executeMultipleRequest == null)
                    {
                        executeMultipleRequest = new ExecuteMultipleRequest()
                        {
                            // Assign settings that define execution behavior: continue on error, return responses.
                            Settings = new ExecuteMultipleSettings()
                            {
                                ContinueOnError = true,
                                ReturnResponses = false
                            },
                            // Create an empty organization request collection.
                            Requests = new OrganizationRequestCollection()
                        };
                    }

                    enrollmentDetail.Attributes["isfs_enrollmentnumberdifferencefromprevious"] = null;

                    executeMultipleRequest.Requests.Add(new UpdateRequest() { Target = enrollmentDetail });

                    if (count > 50)
                    {
                        localContext.OrganizationService.Execute(executeMultipleRequest);
                        executeMultipleRequest = null;
                        count = 0;
                    }
                }

                if (count > 0)
                {
                    localContext.OrganizationService.Execute(executeMultipleRequest);
                }
            }
        }

        private EntityCollection GetPreviousEnrolments(LocalPluginContext localContext, EntityReference entityRef)
        {
            Entity previousEnrolment = null;
            Guid collectionId = Guid.Empty;

            if (entityRef.LogicalName.ToLower().Equals("isfs_isenrolment"))
            {
                previousEnrolment = localContext.OrganizationService.Retrieve("isfs_isenrolment", entityRef.Id,
                    new ColumnSet(new string[] { "isfs_isenrolmentid", "isfs_school", "isfs_enrolmentcollection" }));

                collectionId = previousEnrolment.GetAttributeValue<EntityReference>("isfs_enrolmentcollection").Id;
            }
            else if (entityRef.LogicalName.ToLower().Equals("isfs_enrolmentcollection"))
            {
                collectionId = entityRef.Id;
            }
            else
            {
                throw new Exception("GetPreviousEnrolments doesn't support EnityReference of type " + entityRef.LogicalName);
            }

            QueryExpression query = new QueryExpression("isfs_isenrolmentdetail");
            query.ColumnSet.AddColumns(new string[] { "isfs_isenrolmentdetailid", "isfs_enrollmentnumberdifferencefromprevious" });

            query.Criteria.AddCondition(new ConditionExpression("isfs_enrollmentnumberdifferencefromprevious", ConditionOperator.NotNull));

            LinkEntity linkEnrolment = query.AddLink("isfs_isenrolment", "isfs_isenrolment", "isfs_isenrolmentid", JoinOperator.Inner);

            if (previousEnrolment != null)
                linkEnrolment.LinkCriteria.AddCondition(new ConditionExpression("isfs_school", ConditionOperator.Equal, previousEnrolment.GetAttributeValue<EntityReference>("isfs_school").Id));

            LinkEntity linkCollection = linkEnrolment.AddLink("isfs_enrolmentcollection", "isfs_enrolmentcollection", "isfs_enrolmentcollectionid", JoinOperator.Inner);
            linkCollection.LinkCriteria.AddCondition(new ConditionExpression("isfs_previouscollection", ConditionOperator.Equal, collectionId));

            return localContext.OrganizationService.RetrieveMultiple(query);
        }

        /// <summary>
        /// Flags Enrolments in a current collection for difference calculation, which will trigger their plugin
        /// </summary>
        /// <param name="localContext"></param>
        /// <param name="collectionId"></param>
        private void FlagEnrolmentsForRecalc(LocalPluginContext localContext, Guid collectionId) 
        {
            EntityCollection enrolments = GetEnrolments(localContext, collectionId);

            if (enrolments != null && enrolments.Entities != null && enrolments.Entities.Count > 0)
            {
                ExecuteMultipleRequest executeMultipleRequest = null;
                int count = 0;

                foreach (Entity enrolment in enrolments.Entities)
                {
                    count++;

                    if (executeMultipleRequest == null)
                    {
                        executeMultipleRequest = new ExecuteMultipleRequest()
                        {
                            // Assign settings that define execution behavior: continue on error, return responses.
                            Settings = new ExecuteMultipleSettings()
                            {
                                ContinueOnError = true,
                                ReturnResponses = false
                            },
                            // Create an empty organization request collection.
                            Requests = new OrganizationRequestCollection()
                        };
                    }

                    Entity updatedEnrolment = new Entity("isfs_isenrolment", enrolment.Id);
                    updatedEnrolment.Attributes.Add("isfs_triggerdifferencecalc", true);
                    executeMultipleRequest.Requests.Add(new UpdateRequest() { Target = updatedEnrolment });

                    if (count >= 50)
                    {
                        localContext.OrganizationService.Execute(executeMultipleRequest);
                        executeMultipleRequest = null;
                        count = 0;
                    }
                }

                if (count > 0)
                {
                    localContext.OrganizationService.Execute(executeMultipleRequest);
                }
            }
        }

        private void RemoveEnrolmentDifferences(LocalPluginContext localContext, Entity enrolment)
        {
            EntityReference collectionRef = null;
            EntityCollection enrolmentDetails = GetEnrolmentDetails(localContext, enrolment);

            if (enrolmentDetails != null && enrolmentDetails.Entities != null && enrolmentDetails.Entities.Count > 0)
            {
                ExecuteMultipleRequest executeMultipleRequest = null;
                int count = 0;

                collectionRef = new EntityReference("isfs_enrolmentcollection", (Guid)enrolmentDetails[0].GetAttributeValue<AliasedValue>("Col.isfs_enrolmentcollectionid").Value);

                foreach (Entity enrolmentDetail in enrolmentDetails.Entities)
                {
                    if (enrolmentDetail.Contains("isfs_enrollmentnumberdifferencefromprevious"))
                    {
                        count++;

                        if (executeMultipleRequest == null)
                        {
                            executeMultipleRequest = new ExecuteMultipleRequest()
                            {
                                // Assign settings that define execution behavior: continue on error, return responses.
                                Settings = new ExecuteMultipleSettings()
                                {
                                    ContinueOnError = true,
                                    ReturnResponses = false
                                },
                                // Create an empty organization request collection.
                                Requests = new OrganizationRequestCollection()
                            };
                        }

                        enrolmentDetail.Attributes["isfs_enrollmentnumberdifferencefromprevious"] = null;
                        enrolmentDetail.Attributes["isfs_ignorecreate"] = false;

                        executeMultipleRequest.Requests.Add(new UpdateRequest() { Target = enrolmentDetail });

                        if (count >= 75)
                        {
                            localContext.OrganizationService.Execute(executeMultipleRequest);
                            executeMultipleRequest = null;
                            count = 0;
                        }
                    }
                }

                if (count > 0)
                {
                    localContext.OrganizationService.Execute(executeMultipleRequest);
                }
            }

            Entity updatedEnrolment = new Entity("isfs_isenrolment", enrolment.Id);
            updatedEnrolment.Attributes["isfs_triggerdifferencecalc"] = false;

            localContext.OrganizationService.Update(updatedEnrolment);

            // Check if other isfs_isenrolment in collection still recalculating
            if (collectionRef != null && !StillCalculating(localContext, collectionRef.Id))
            {
                SetCalculatingDeltas(localContext, collectionRef.Id, false);
            }
        }

        private void RemoveAllDifferences(LocalPluginContext localContext, Guid collectionId)
        {
            EntityCollection entityDetails = GetAllEnrolmentDetails(localContext, collectionId);

            if (entityDetails != null && entityDetails.Entities != null && entityDetails.Entities.Count > 0)
            {
                ExecuteMultipleRequest executeMultipleRequest = null;
                int count = 0;

                foreach (Entity enrolmentDetail in entityDetails.Entities)
                {
                    if (enrolmentDetail.Contains("isfs_enrollmentnumberdifferencefromprevious"))
                    {
                        count++;

                        if (executeMultipleRequest == null)
                        {
                            executeMultipleRequest = new ExecuteMultipleRequest()
                            {
                                // Assign settings that define execution behavior: continue on error, return responses.
                                Settings = new ExecuteMultipleSettings()
                                {
                                    ContinueOnError = true,
                                    ReturnResponses = false
                                },
                                // Create an empty organization request collection.
                                Requests = new OrganizationRequestCollection()
                            };
                        }

                        enrolmentDetail.Attributes["isfs_enrollmentnumberdifferencefromprevious"] = null;

                        executeMultipleRequest.Requests.Add(new UpdateRequest() { Target = enrolmentDetail });

                        if (count >= 75)
                        {
                            localContext.OrganizationService.Execute(executeMultipleRequest);
                            executeMultipleRequest = null;
                            count = 0;
                        }
                    }
                }

                if (count > 0)
                {
                    localContext.OrganizationService.Execute(executeMultipleRequest);
                }
            }
        }

        private void RecalculateEnrollmentDetailDifference(LocalPluginContext localContext, Entity enrolmentDetail, bool isDelete)
        {
            EntityCollection presentCollections = null;
            Entity previousCollection = null;
            Guid previousCollectionId;

            Entity collection = GetCollectionForEnrolmentDetail(localContext, enrolmentDetail.Id);

            //Determine Current and Previous collections
            if (collection.Contains("isfs_previouscollection"))
            {
                previousCollectionId = collection.GetAttributeValue<EntityReference>("isfs_previouscollection").Id;

                previousCollection = GetPreviousCollection(localContext, collection, previousCollectionId);
                presentCollections = GetCurrentCollectionsForPreviousCollection(localContext, collection, previousCollectionId);
            }
            else
            {
                previousCollectionId = collection.Id;

                previousCollection = collection;
                presentCollections = GetCurrentCollectionsForPreviousCollection(localContext, previousCollection, previousCollectionId);
            }

            if (isDelete || (previousCollection != null && presentCollections != null && presentCollections.Entities != null && presentCollections.Entities.Count > 0))
                CalculateDifference(localContext, presentCollections, previousCollection, enrolmentDetail.Id, isDelete);

            if (!isDelete)
            { 
                if (presentCollections == null || presentCollections.Entities == null || presentCollections.Entities.Count == 0)
                {
                    presentCollections = GetAllCurrentCollectionsForPreviousCollection(localContext, previousCollection, previousCollectionId);
                }

                if (previousCollection != null && (presentCollections != null && presentCollections.Entities != null && presentCollections.Entities.Count > 0))
                {
                    foreach (Entity coll in presentCollections.Entities)
                    {
                        Guid enrolmentId = (Guid)coll.GetAttributeValue<AliasedValue>("enr.isfs_isenrolmentid").Value;

                        enrolmentDetail = GetEnrolmentDetail(localContext, enrolmentDetail.Id);

                        if (!HasMatchingEnrolmentDetail(localContext, coll, previousCollection))
                        {
                            Entity currentEnrolmentDetail = CreateCurrentEnrolmentDetailFromPrevious(localContext, enrolmentId, enrolmentDetail);

                            localContext.OrganizationService.Create(currentEnrolmentDetail);
                        }
                    }
                }
            }
        }

        private bool HasMatchingEnrolmentDetail(LocalPluginContext localContext, Entity coll, Entity previousCollection)
        {
            QueryExpression query = new QueryExpression("isfs_enrolmentcollection");
            query.ColumnSet.AddColumns(new string[] { "isfs_enrolmentcollectionid", "isfs_previouscollection" });

            query.Criteria.AddCondition(new ConditionExpression("isfs_previouscollection", ConditionOperator.Equal, coll.Id));

            LinkEntity linkEnrolment = query.AddLink("isfs_isenrolment", "isfs_enrolmentcollectionid", "isfs_enrolmentcollection", JoinOperator.Inner);
            linkEnrolment.EntityAlias = "enr";
            linkEnrolment.Columns = new ColumnSet(new string[] { "isfs_isenrolmentid", "isfs_school" });

            linkEnrolment.LinkCriteria.AddCondition(new ConditionExpression("isfs_school", ConditionOperator.Equal, ((EntityReference)(previousCollection.GetAttributeValue<AliasedValue>("enr.isfs_school").Value)).Id));

            LinkEntity linkEnrolmentDetail = linkEnrolment.AddLink("isfs_isenrolmentdetail", "isfs_isenrolmentid", "isfs_isenrolment", JoinOperator.Inner);
            linkEnrolmentDetail.EntityAlias = "enr_det";
            linkEnrolmentDetail.Columns = new ColumnSet(new string[] { "isfs_isenrolmentdetailid", "isfs_enrolmentnumber", "isfs_adjustedenrolmentnumber", "isfs_enrollmentnumberdifferencefromprevious", "isfs_schoolsubcategory", "isfs_esgroup", "isfs_esgroupsubcategory", "isfs_isenrolment" });

            if (previousCollection.Contains("enr_det.isfs_schoolsubcategory"))
                linkEnrolmentDetail.LinkCriteria.AddCondition(new ConditionExpression("isfs_schoolsubcategory", ConditionOperator.Equal, ((EntityReference)previousCollection.GetAttributeValue<AliasedValue>("enr_det.isfs_schoolsubcategory").Value).Id));

            if (previousCollection.Contains("enr_det.isfs_schoolsubcategory"))
                linkEnrolmentDetail.LinkCriteria.AddCondition(new ConditionExpression("isfs_esgroup", ConditionOperator.Equal, ((EntityReference)previousCollection.GetAttributeValue<AliasedValue>("enr_det.isfs_esgroup").Value).Id));

            if (previousCollection.Contains("enr_det.isfs_esgroupsubcategory"))
                linkEnrolmentDetail.LinkCriteria.AddCondition(new ConditionExpression("isfs_esgroupsubcategory", ConditionOperator.Equal, ((EntityReference)previousCollection.GetAttributeValue<AliasedValue>("enr_det.isfs_esgroupsubcategory").Value).Id));

            EntityCollection results = localContext.OrganizationService.RetrieveMultiple(query);

            if (results != null && results.Entities != null && results.Entities.Count > 0) return true;
            return false;
        }

        private Entity GetEnrolmentDetail(LocalPluginContext localContext, Guid enrolmentDetailId)
        {
            QueryExpression query = new QueryExpression("isfs_isenrolmentdetail");

            query.Criteria.AddCondition(new ConditionExpression("isfs_isenrolmentdetailid", ConditionOperator.Equal, enrolmentDetailId));

            query.ColumnSet.AddColumns(new string[] { "isfs_isenrolmentdetailid", "isfs_enrolmentnumber", "isfs_adjustedenrolmentnumber", "isfs_enrollmentnumberdifferencefromprevious", "isfs_schoolsubcategory", "isfs_esgroup", "isfs_esgroupsubcategory", "isfs_isenrolment", "isfs_schoolyear" });

            LinkEntity linkEnrolment = query.AddLink("isfs_isenrolment", "isfs_isenrolment", "isfs_isenrolmentid", JoinOperator.Inner);
            linkEnrolment.EntityAlias = "Enr";
            linkEnrolment.Columns.AddColumn("isfs_school");

            LinkEntity linkCollection = linkEnrolment.AddLink("isfs_enrolmentcollection", "isfs_enrolmentcollection", "isfs_enrolmentcollectionid", JoinOperator.Inner);
            linkCollection.EntityAlias = "Col";
            linkCollection.Columns.AddColumns(new string[] { "isfs_enrolmentcollectionid", "isfs_previouscollection" });

            return localContext.OrganizationService.RetrieveMultiple(query).Entities.FirstOrDefault();
        }

        private void CalculateDifference(LocalPluginContext localContext, EntityCollection presentCollections, Entity previousCollection, Guid enrolmentDetailId, bool isDelete)
        {
            foreach (Entity presentCollection in presentCollections.Entities)
            {
                decimal enrolmentNum = 0;
                decimal previousEnrolmentNum = 0;
                decimal enrolmentDelta = 0;
                decimal currentDelta = 0;
                bool hasCurrentDelta = false;

                if (!isDelete)
                {
                    if (previousCollection.Contains("enr_det.isfs_adjustedenrolmentnumber"))
                        previousEnrolmentNum = (decimal)previousCollection.GetAttributeValue<AliasedValue>("enr_det.isfs_adjustedenrolmentnumber").Value;

                    if (previousEnrolmentNum == 0 && previousCollection.Contains("enr_det.isfs_enrolmentnumber"))
                        previousEnrolmentNum = (decimal)previousCollection.GetAttributeValue<AliasedValue>("enr_det.isfs_enrolmentnumber").Value;

                    if (presentCollection.Contains("enr_det.isfs_enrollmentnumberdifferencefromprevious"))
                    {
                        currentDelta = (decimal)presentCollection.GetAttributeValue<AliasedValue>("enr_det.isfs_enrollmentnumberdifferencefromprevious").Value;
                        hasCurrentDelta = true;
                    }

                    if (presentCollection.Contains("enr_det.isfs_adjustedenrolmentnumber"))
                        enrolmentNum = (decimal)presentCollection.GetAttributeValue<AliasedValue>("enr_det.isfs_adjustedenrolmentnumber").Value;


                    if (enrolmentNum == 0 && presentCollection.Contains("enr_det.isfs_enrolmentnumber"))
                        enrolmentNum = (decimal)presentCollection.GetAttributeValue<AliasedValue>("enr_det.isfs_enrolmentnumber").Value;

                    enrolmentDelta = enrolmentNum - previousEnrolmentNum;
                }

                if (isDelete || (enrolmentDelta != currentDelta) || !hasCurrentDelta)
                {
                    Entity enrolmentDetail = new Entity("isfs_isenrolmentdetail", ((Guid)presentCollection.GetAttributeValue<AliasedValue>("enr_det.isfs_isenrolmentdetailid").Value));

                    if (!isDelete)
                        enrolmentDetail.Attributes.Add("isfs_enrollmentnumberdifferencefromprevious", enrolmentDelta);
                    else
                        enrolmentDetail.Attributes.Add("isfs_enrollmentnumberdifferencefromprevious", null);

                    localContext.OrganizationService.Update(enrolmentDetail);
                }
            }
        }

        private Entity GetMatchingEnrollmentDetail(LocalPluginContext localContext, Entity collection, Entity enrollmentDetail)
        {
            QueryExpression query = new QueryExpression("isfs_isenrolmentdetail");

            query.ColumnSet.AddColumns(new string[] { "isfs_isenrolmentdetailid", "isfs_enrolmentnumber", "isfs_adjustedenrolmentnumber", "isfs_enrollmentnumberdifferencefromprevious", "isfs_schoolsubcategory", "isfs_esgroup", "isfs_esgroupsubcategory", "isfs_isenrolment" });

            if (enrollmentDetail.Contains("isfs_schoolsubcategory"))
                query.Criteria.AddCondition(new ConditionExpression("isfs_schoolsubcategory", ConditionOperator.Equal, enrollmentDetail.GetAttributeValue<EntityReference>("isfs_schoolsubcategory").Id));

            if (enrollmentDetail.Contains("isfs_esgroup"))
                query.Criteria.AddCondition(new ConditionExpression("isfs_esgroup", ConditionOperator.Equal, enrollmentDetail.GetAttributeValue<EntityReference>("isfs_esgroup").Id));

            if (enrollmentDetail.Contains("isfs_esgroupsubcategory"))
                query.Criteria.AddCondition(new ConditionExpression("isfs_esgroupsubcategory", ConditionOperator.Equal, enrollmentDetail.GetAttributeValue<EntityReference>("isfs_esgroupsubcategory").Id));

            return localContext.OrganizationService.RetrieveMultiple(query).Entities.FirstOrDefault();
        }

        private Entity GetPreviousCollection(LocalPluginContext localContext, Entity collection, Guid previousCollectionId)
        {
            QueryExpression query = new QueryExpression("isfs_enrolmentcollection");
            query.ColumnSet.AddColumns(new string[] { "isfs_enrolmentcollectionid" });

            query.Criteria.AddCondition(new ConditionExpression("isfs_enrolmentcollectionid", ConditionOperator.Equal, previousCollectionId));

            LinkEntity linkEnrolment = query.AddLink("isfs_isenrolment", "isfs_enrolmentcollectionid", "isfs_enrolmentcollection", JoinOperator.Inner);
            linkEnrolment.EntityAlias = "enr";
            linkEnrolment.Columns = new ColumnSet(new string[] { "isfs_isenrolmentid", "isfs_school" });

            linkEnrolment.LinkCriteria.AddCondition(new ConditionExpression("isfs_school", ConditionOperator.Equal, ((EntityReference)(collection.GetAttributeValue<AliasedValue>("enr.isfs_school").Value)).Id));

            LinkEntity linkEnrolmentDetail = linkEnrolment.AddLink("isfs_isenrolmentdetail", "isfs_isenrolmentid", "isfs_isenrolment", JoinOperator.Inner);
            linkEnrolmentDetail.EntityAlias = "enr_det";
            linkEnrolmentDetail.Columns = new ColumnSet(new string[] { "isfs_isenrolmentdetailid", "isfs_enrolmentnumber", "isfs_adjustedenrolmentnumber", "isfs_enrollmentnumberdifferencefromprevious", "isfs_schoolsubcategory", "isfs_esgroup", "isfs_esgroupsubcategory", "isfs_isenrolment" });

            if (collection.Contains("enr_det.isfs_schoolsubcategory"))
                linkEnrolmentDetail.LinkCriteria.AddCondition(new ConditionExpression("isfs_schoolsubcategory", ConditionOperator.Equal, ((EntityReference)collection.GetAttributeValue<AliasedValue>("enr_det.isfs_schoolsubcategory").Value).Id));

            if (collection.Contains("enr_det.isfs_esgroup"))
                linkEnrolmentDetail.LinkCriteria.AddCondition(new ConditionExpression("isfs_esgroup", ConditionOperator.Equal, ((EntityReference)collection.GetAttributeValue<AliasedValue>("enr_det.isfs_esgroup").Value).Id));

            if (collection.Contains("enr_det.isfs_esgroupsubcategory"))
                linkEnrolmentDetail.LinkCriteria.AddCondition(new ConditionExpression("isfs_esgroupsubcategory", ConditionOperator.Equal, ((EntityReference)collection.GetAttributeValue<AliasedValue>("enr_det.isfs_esgroupsubcategory").Value).Id));

            return localContext.OrganizationService.RetrieveMultiple(query).Entities.FirstOrDefault();
        }

        private EntityCollection GetCurrentCollectionsForPreviousCollection(LocalPluginContext localContext, Entity previousCollection, Guid previousCollectionId)
        {
            QueryExpression query = new QueryExpression("isfs_enrolmentcollection");
            query.ColumnSet.AddColumns(new string[] { "isfs_enrolmentcollectionid", "isfs_previouscollection" });

            query.Criteria.AddCondition(new ConditionExpression("isfs_previouscollection", ConditionOperator.Equal, previousCollectionId));

            LinkEntity linkEnrolment = query.AddLink("isfs_isenrolment", "isfs_enrolmentcollectionid", "isfs_enrolmentcollection", JoinOperator.Inner);
            linkEnrolment.EntityAlias = "enr";
            linkEnrolment.Columns = new ColumnSet(new string[] { "isfs_isenrolmentid", "isfs_school" });

            linkEnrolment.LinkCriteria.AddCondition(new ConditionExpression("isfs_school", ConditionOperator.Equal, ((EntityReference)(previousCollection.GetAttributeValue<AliasedValue>("enr.isfs_school").Value)).Id));

            LinkEntity linkEnrolmentDetail = linkEnrolment.AddLink("isfs_isenrolmentdetail", "isfs_isenrolmentid", "isfs_isenrolment", JoinOperator.Inner);
            linkEnrolmentDetail.EntityAlias = "enr_det";
            linkEnrolmentDetail.Columns = new ColumnSet(new string[] { "isfs_isenrolmentdetailid", "isfs_enrolmentnumber", "isfs_adjustedenrolmentnumber", "isfs_enrollmentnumberdifferencefromprevious", "isfs_schoolsubcategory", "isfs_esgroup", "isfs_esgroupsubcategory", "isfs_isenrolment" });

            if (previousCollection.Contains("enr_det.isfs_schoolsubcategory"))
                linkEnrolmentDetail.LinkCriteria.AddCondition(new ConditionExpression("isfs_schoolsubcategory", ConditionOperator.Equal, ((EntityReference)previousCollection.GetAttributeValue<AliasedValue>("enr_det.isfs_schoolsubcategory").Value).Id));

            if (previousCollection.Contains("enr_det.isfs_esgroup"))
                linkEnrolmentDetail.LinkCriteria.AddCondition(new ConditionExpression("isfs_esgroup", ConditionOperator.Equal, ((EntityReference)previousCollection.GetAttributeValue<AliasedValue>("enr_det.isfs_esgroup").Value).Id));

            if (previousCollection.Contains("enr_det.isfs_esgroupsubcategory"))
                linkEnrolmentDetail.LinkCriteria.AddCondition(new ConditionExpression("isfs_esgroupsubcategory", ConditionOperator.Equal, ((EntityReference)previousCollection.GetAttributeValue<AliasedValue>("enr_det.isfs_esgroupsubcategory").Value).Id));

            return localContext.OrganizationService.RetrieveMultiple(query);
        }

        private EntityCollection GetAllCurrentCollectionsForPreviousCollection(LocalPluginContext localContext, Entity previousCollection, Guid previousCollectionId)
        {
            QueryExpression query = new QueryExpression("isfs_enrolmentcollection");
            query.ColumnSet.AddColumns(new string[] { "isfs_enrolmentcollectionid", "isfs_previouscollection" });

            query.Criteria.AddCondition(new ConditionExpression("isfs_previouscollection", ConditionOperator.Equal, previousCollectionId));

            LinkEntity linkEnrolment = query.AddLink("isfs_isenrolment", "isfs_enrolmentcollectionid", "isfs_enrolmentcollection", JoinOperator.Inner);
            linkEnrolment.EntityAlias = "enr";
            linkEnrolment.Columns = new ColumnSet(new string[] { "isfs_isenrolmentid", "isfs_school" });

            linkEnrolment.LinkCriteria.AddCondition(new ConditionExpression("isfs_school", ConditionOperator.Equal, ((EntityReference)(previousCollection.GetAttributeValue<AliasedValue>("enr.isfs_school").Value)).Id));

            return localContext.OrganizationService.RetrieveMultiple(query);
        }

        private EntityCollection GetEnrolments(LocalPluginContext localContext, Guid collectionId)
        {
            QueryExpression query = new QueryExpression("isfs_isenrolment");

            query.Criteria.AddCondition(new ConditionExpression("isfs_enrolmentcollection", ConditionOperator.Equal, collectionId));

            query.ColumnSet.AddColumns(new string[] { "isfs_isenrolmentid", "isfs_school", "isfs_enrolmentcollection" });

            return localContext.OrganizationService.RetrieveMultiple(query);
        }
        
        private EntityCollection GetEnrolmentDetails(LocalPluginContext localContext, Entity enrolment)
        {
            QueryExpression query = new QueryExpression("isfs_isenrolmentdetail");

            query.Criteria.AddCondition(new ConditionExpression("isfs_isenrolment", ConditionOperator.Equal, enrolment.Id));

            query.ColumnSet.AddColumns(new string[] { "isfs_isenrolmentdetailid", "isfs_enrolmentnumber", "isfs_adjustedenrolmentnumber", "isfs_enrollmentnumberdifferencefromprevious", "isfs_schoolsubcategory", "isfs_esgroup", "isfs_esgroupsubcategory", "isfs_isenrolment", "isfs_schoolyear" });

            LinkEntity linkEnrolment = query.AddLink("isfs_isenrolment", "isfs_isenrolment", "isfs_isenrolmentid", JoinOperator.Inner);
            linkEnrolment.EntityAlias = "Enr";
            linkEnrolment.Columns.AddColumn("isfs_school");

            LinkEntity linkCollection = linkEnrolment.AddLink("isfs_enrolmentcollection", "isfs_enrolmentcollection", "isfs_enrolmentcollectionid", JoinOperator.Inner);
            linkCollection.EntityAlias = "Col";
            linkCollection.Columns.AddColumns(new string[] { "isfs_enrolmentcollectionid", "isfs_previouscollection" });

            return localContext.OrganizationService.RetrieveMultiple(query);
        }

        private EntityCollection GetPreviousEnrolmentDetails(LocalPluginContext localContext, Guid previousCollectionId, Entity currentEnrolment)
        {
            QueryExpression query = new QueryExpression("isfs_isenrolmentdetail");

            EntityReference schoolRef = currentEnrolment.GetAttributeValue<EntityReference>("isfs_school");

            query.ColumnSet.AddColumns(new string[] { "isfs_isenrolmentdetailid", "isfs_enrolmentnumber", "isfs_adjustedenrolmentnumber", "isfs_enrollmentnumberdifferencefromprevious", "isfs_schoolsubcategory", "isfs_esgroup", "isfs_esgroupsubcategory", "isfs_isenrolment", "isfs_schoolyear", "isfs_enrolmentnumbertype", "isfs_fundinggroup", "isfs_mappedesgroupsubcategory" });

            LinkEntity linkEnrolment = query.AddLink("isfs_isenrolment", "isfs_isenrolment", "isfs_isenrolmentid", JoinOperator.Inner);
            linkEnrolment.EntityAlias = "enr";
            linkEnrolment.Columns = new ColumnSet(new string[] { "isfs_isenrolmentid", "isfs_school" });
            linkEnrolment.LinkCriteria.AddCondition(new ConditionExpression("isfs_enrolmentcollection", ConditionOperator.Equal, previousCollectionId));
            linkEnrolment.LinkCriteria.AddCondition(new ConditionExpression("isfs_school", ConditionOperator.Equal, schoolRef.Id));

            return localContext.OrganizationService.RetrieveMultiple(query);
        }

        private EntityCollection GetAllEnrolmentDetails(LocalPluginContext localContext, Guid collectionId)
        {
            QueryExpression query = new QueryExpression("isfs_isenrolmentdetail");

            query.ColumnSet.AddColumns(new string[] { "isfs_isenrolmentdetailid", "isfs_enrollmentnumberdifferencefromprevious" });

            query.Criteria.AddCondition(new ConditionExpression("isfs_enrollmentnumberdifferencefromprevious", ConditionOperator.NotNull));

            LinkEntity link = query.AddLink("isfs_isenrolment", "isfs_isenrolment", "isfs_isenrolmentid", JoinOperator.LeftOuter);

            link.LinkCriteria.AddCondition("isfs_enrolmentcollection", ConditionOperator.Equal, collectionId);

            return localContext.OrganizationService.RetrieveMultiple(query);
        }

        private Entity GetPreviousEnrolment(LocalPluginContext localContext, Entity enrolment, Guid previousCollectionId)
        {
            EntityReference schoolRef = enrolment.GetAttributeValue<EntityReference>("isfs_school");

            QueryExpression query = new QueryExpression("isfs_isenrolment");

            query.Criteria.AddCondition(new ConditionExpression("isfs_enrolmentcollection", ConditionOperator.Equal, previousCollectionId));
            query.Criteria.AddCondition(new ConditionExpression("isfs_school", ConditionOperator.Equal, schoolRef.Id));

            query.ColumnSet.AddColumns(new string[] { "isfs_isenrolmentid", "isfs_school", "isfs_enrolmentcollection" });

            return localContext.OrganizationService.RetrieveMultiple(query).Entities.FirstOrDefault();
        }

        private Entity FindMatchingEnrolmentDetail(LocalPluginContext localContext, Entity enrollmentDetail, Guid previousEnrollmentId)
        {
            QueryExpression query = new QueryExpression("isfs_isenrolmentdetail");

            query.ColumnSet.AddColumns(new string[] { "isfs_isenrolmentdetailid", "isfs_enrolmentnumber", "isfs_adjustedenrolmentnumber", "isfs_enrollmentnumberdifferencefromprevious", "isfs_schoolsubcategory", "isfs_esgroup", "isfs_esgroupsubcategory", "isfs_isenrolment" });

            query.Criteria.AddCondition(new ConditionExpression("isfs_isenrolment", ConditionOperator.Equal, previousEnrollmentId));

            if (enrollmentDetail.Contains("isfs_schoolsubcategory"))
                query.Criteria.AddCondition(new ConditionExpression("isfs_schoolsubcategory", ConditionOperator.Equal, enrollmentDetail.GetAttributeValue<EntityReference>("isfs_schoolsubcategory").Id));

            if (enrollmentDetail.Contains("isfs_esgroup"))
                query.Criteria.AddCondition(new ConditionExpression("isfs_esgroup", ConditionOperator.Equal, enrollmentDetail.GetAttributeValue<EntityReference>("isfs_esgroup").Id));

            if (enrollmentDetail.Contains("isfs_esgroupsubcategory"))
                query.Criteria.AddCondition(new ConditionExpression("isfs_esgroupsubcategory", ConditionOperator.Equal, enrollmentDetail.GetAttributeValue<EntityReference>("isfs_esgroupsubcategory").Id));

            return localContext.OrganizationService.RetrieveMultiple(query).Entities.FirstOrDefault();
        }

        private Entity GetCollectionForEnrolmentDetail(LocalPluginContext localContext, Guid enrolmentDetailId)
        {
            QueryExpression query = new QueryExpression("isfs_enrolmentcollection");

            query.ColumnSet = new ColumnSet(new string[] { "isfs_enrolmentcollectionid", "isfs_previouscollection" });

            LinkEntity linkEnrolment = query.AddLink("isfs_isenrolment", "isfs_enrolmentcollectionid", "isfs_enrolmentcollection", JoinOperator.Inner);
            linkEnrolment.EntityAlias = "enr";
            linkEnrolment.Columns = new ColumnSet(new string[] { "isfs_isenrolmentid", "isfs_school" });

            LinkEntity linkEnrolmentDetail = linkEnrolment.AddLink("isfs_isenrolmentdetail", "isfs_isenrolmentid", "isfs_isenrolment", JoinOperator.Inner);
            linkEnrolmentDetail.EntityAlias = "enr_det";
            linkEnrolmentDetail.Columns = new ColumnSet(new string[] { "isfs_isenrolmentdetailid", "isfs_enrolmentnumber", "isfs_adjustedenrolmentnumber", "isfs_enrollmentnumberdifferencefromprevious", "isfs_schoolsubcategory", "isfs_esgroup", "isfs_esgroupsubcategory", "isfs_isenrolment" });

            linkEnrolmentDetail.LinkCriteria.AddCondition(new ConditionExpression("isfs_isenrolmentdetailid", ConditionOperator.Equal, enrolmentDetailId));

            return localContext.OrganizationService.RetrieveMultiple(query).Entities.FirstOrDefault();
        }

        private void SetCalculatingDeltas(LocalPluginContext localContext, Guid collectionId, bool isCalculating)
        {
            if (!collectionId.Equals(Guid.Empty))
            {
                Entity collection = new Entity("isfs_enrolmentcollection", collectionId);
                collection.Attributes.Add("isfs_iscalculatingenrollmentdeltas", isCalculating);

                localContext.OrganizationService.Update(collection);
            }
        }
    }
}
