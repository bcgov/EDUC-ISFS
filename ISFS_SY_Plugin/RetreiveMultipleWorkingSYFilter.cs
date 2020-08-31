using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace MoE_ISFS.ISFSPlugins
{
    /// <summary>
    /// RetreiveMultiple plugin used to filter all views with School Year field using user's current working school years
    /// </summary>
    public class RetreiveMultipleWorkingSYFilter: PluginBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PostAuthorityPaymentUpdate"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public RetreiveMultipleWorkingSYFilter(string unsecure, string secure)
            : base(typeof(RetreiveMultipleWorkingSYFilter))
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
            localContext.Trace("RetreiveMultipleWorkingSYFilter Plugin: Starting");

            if (localContext == null) throw new InvalidPluginExecutionException(OperationStatus.Failed, "Missing localContext.");

            FilterViewWithWorkingSY(localContext);
        }

        private void FilterViewWithWorkingSY(LocalPluginContext localContext)
        {
            IPluginExecutionContext context = localContext.PluginExecutionContext;

            if (context.InputParameters["Query"] == null)
            {
                localContext.Trace("RetreiveMultipleWorkingSYFilter Plugin: Missing InputParameter 'Query', exiting with error.");
                throw new InvalidPluginExecutionException(OperationStatus.Failed, "Missing InputParameter 'Query'.");
            }


            // STEP 0: Make sure the query is a QueryExpression or FetchExpression object
            #region STEP0
            QueryExpression qExpression = null;
            FetchExpression fExpression = null;

            if (localContext.PluginExecutionContext.InputParameters["Query"].GetType() == typeof(QueryExpression))
            {
                qExpression = (QueryExpression)localContext.PluginExecutionContext.InputParameters["Query"];
            }
            else if (localContext.PluginExecutionContext.InputParameters["Query"].GetType() == typeof(FetchExpression))
            {
                fExpression = (FetchExpression)localContext.PluginExecutionContext.InputParameters["Query"];
            }
            else if (localContext.PluginExecutionContext.InputParameters["Query"].GetType() == typeof(QueryByAttribute))
            {
                localContext.Trace("RetreiveMultipleWorkingSYFilter Plugin: QueryByAttribute query, exiting.");
                return;
            }
            else
            {
                string message = "Unexpected Query type: " + localContext.PluginExecutionContext.InputParameters["Query"].GetType().ToString();
                localContext.Trace("RetreiveMultipleWorkingSYFilter Plugin: " + message + ", exiting with error.");
                throw new InvalidPluginExecutionException(OperationStatus.Failed, message);
            }
            #endregion

            // STEP 1: Check if the view has the SY field in it. Return if it doesn't as we don't need to append filter
            #region STEP1
            if ((qExpression != null && !qExpression.ColumnSet.Columns.Contains("isfs_schoolyear")) || (fExpression != null && fExpression.Query.IndexOf("<attribute name=\"isfs_schoolyear\"") == -1))
            {
                localContext.Trace("RetreiveMultipleWorkingSYFilter Plugin: isfs_schoolyearid field not in view, exiting.");
                return;
            }

            localContext.Trace("RetreiveMultipleWorkingSYFilter Plugin: isfs_schoolyearid found in view.");
            #endregion


            // STEP 2: Check if isfs_ignoresyfilter field is in query. Don't modify query if it is.
            #region STEP2
            bool skip = false;
            if (qExpression != null)
            {
                if (qExpression.Criteria.Filters != null && qExpression.Criteria.Filters.Count > 0)
                {
                    skip = QueryExpressionHasIgnoreFilter(qExpression.Criteria.Filters);
                }
                else if (qExpression.Criteria.Conditions != null && qExpression.Criteria.Conditions.Count > 0)
                {
                    skip = QueryConditionHasIgnoreFilter(qExpression.Criteria.Conditions);
                }
            }
            else if (fExpression != null && fExpression.Query.IndexOf("isfs_ignoresyfilter") > -1) skip = true;

            if (skip == true)
            {
                localContext.Trace("RetreiveMultipleWorkingSYFilter Plugin: isfs_ignoresyfilter field query, exiting.");
                return;
            }

            localContext.Trace("RetreiveMultipleWorkingSYFilter Plugin: isfs_ignoresyfilter field not found in query, continuing");
            #endregion


            // STEP 3: Check if the user has working SY(s) set
            #region STEP3
            EntityCollection usersWorkingSYs = GetUsersWorkingSYs(localContext);
            if (usersWorkingSYs == null || usersWorkingSYs.Entities == null || usersWorkingSYs.Entities.Count == 0)
            {
                localContext.Trace("RetreiveMultipleWorkingSYFilter Plugin: No user working years found, exiting.");
                return;
            }

            localContext.Trace("RetreiveMultipleWorkingSYFilter Plugin: User working years found: " + usersWorkingSYs.Entities.Count);
            #endregion


            // STEP 4: If view has SY field and user has working SY(s) set, filter the view to the user's working SY(s)
            #region STEP4
            if (qExpression != null)
            {
                FilterExpression workingSYFilter = new FilterExpression(LogicalOperator.Or);
                foreach (Entity entity in usersWorkingSYs.Entities)
                {
                    if (entity.Attributes.Contains("isfs_schoolyearid"))
                    {
                        workingSYFilter.AddCondition(new ConditionExpression("isfs_schoolyear", ConditionOperator.Equal, entity.GetAttributeValue<EntityReference>("isfs_schoolyearid").Id));
                    }
                }

                if (qExpression != null) qExpression.Criteria.AddFilter(workingSYFilter);
            }
            else if (fExpression != null)
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(fExpression.Query);

                XmlNode subFilter = doc.CreateElement("filter");
                XmlAttribute attr = doc.CreateAttribute("type");
                attr.Value = "or";
                subFilter.Attributes.Append(attr);

                foreach (Entity entity in usersWorkingSYs.Entities)
                {
                    if (entity.Attributes.Contains("isfs_schoolyearid"))
                    {
                        XmlNode condition = doc.CreateElement("condition");

                        attr = doc.CreateAttribute("attribute");
                        attr.Value = "isfs_schoolyear";
                        condition.Attributes.Append(attr);

                        attr = doc.CreateAttribute("operator");
                        attr.Value = "eq";
                        condition.Attributes.Append(attr);

                        attr = doc.CreateAttribute("value");
                        attr.Value = entity.GetAttributeValue<EntityReference>("isfs_schoolyearid").Id.ToString();
                        condition.Attributes.Append(attr);

                        subFilter.AppendChild(condition);
                    }
                }

                XmlNode filter = doc.SelectSingleNode("/fetch/entity/filter[1]");
                if (filter != null)
                {
                    if (filter.Attributes["type"] == null || filter.Attributes["type"].Value.ToLower().Equals("and")) filter.AppendChild(subFilter);
                    else
                    {
                        // Existing filter is an or filter, we need to add a parent "and" filter
                        XmlNode parent = filter.ParentNode;
                        parent.RemoveChild(filter);

                        XmlNode mainFilter = doc.CreateElement("filter");
                        attr = doc.CreateAttribute("type");
                        attr.Value = "and";
                        mainFilter.Attributes.Append(attr);

                        mainFilter.AppendChild(filter);
                        mainFilter.AppendChild(subFilter);

                        parent.AppendChild(mainFilter);
                    }
                }
                else doc.SelectSingleNode("/fetch/entity").AppendChild(subFilter); 

                fExpression.Query = doc.OuterXml;
            }
            #endregion


            localContext.Trace("RetreiveMultipleWorkingSYFilter Plugin: User working years added to query filter, exiting.");
        }


        /// <summary>
        /// User to recursively check if isfs_ignoresyfilter field is used in query
        /// </summary>
        /// <param name="filters"></param>
        /// <returns></returns>
        private bool QueryExpressionHasIgnoreFilter(DataCollection<FilterExpression> filters)
        {
            bool found = false;

            if (filters != null && filters.Count > 0)
            {
                foreach (FilterExpression filter in filters)
                {
                    if (filter.Conditions != null && filter.Conditions.Count > 0)
                        found = QueryConditionHasIgnoreFilter(filter.Conditions);

                    if (!found && filter.Filters != null && filter.Filters.Count > 0)
                        found = QueryExpressionHasIgnoreFilter(filter.Filters);

                    if (found == true)
                        break;
                }
            }

            return found;
        }


        /// <summary>
        /// Used to check if isfs_ignoresyfilter filter condition exists in query
        /// </summary>
        /// <param name="conditions"></param>
        /// <returns></returns>
        private bool QueryConditionHasIgnoreFilter(DataCollection<ConditionExpression> conditions)
        {
            if (conditions != null && conditions.Count > 0)
            {
                foreach (ConditionExpression condition in conditions)
                {
                    if (condition.AttributeName.ToLower().Equals("isfs_ignoresyfilter")) return true;
                }
            }

            return false;
        }


        /// <summary>
        /// Get's the user's working school years
        /// </summary>
        /// <param name="localContext"></param>
        /// <returns></returns>
        private EntityCollection GetUsersWorkingSYs(LocalPluginContext localContext)
        {
            QueryExpression query = new QueryExpression("isfs_workingschoolyear");
            query.ColumnSet.AddColumn("isfs_schoolyearid");
            query.Criteria.AddCondition(new ConditionExpression("isfs_userid", ConditionOperator.Equal, localContext.PluginExecutionContext.UserId));

            return localContext.OrganizationService.RetrieveMultiple(query);
        }
    }
}
