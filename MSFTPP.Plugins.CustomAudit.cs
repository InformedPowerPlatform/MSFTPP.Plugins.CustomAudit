using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Security.Policy;
using System.Text;

namespace MSFTPP.Plugin.CustomAudit
{
    // Note: Things that will need to be re-worked
    // 1. How to get the prefix of the entity being audited to map to the lookup for that entity? (line 100)
    // 2. A possible configuration option per entity if the customer wants to audit a case change to a field (Line 276)
    //      e.g. John Smith -> JOHN SMITH (Do they want to audit if the values are the same, just the case changed)

    public class MSFTPPPluginCustomAudit : IPlugin
    {
        #region Secure/Unsecure Configuration Setup
        private string _secureConfig = null;
        private string _unsecureConfig = null;
        private string _inUseConfig = null;

        public MSFTPPPluginCustomAudit(string unsecureConfig, string secureConfig)
        {
            _secureConfig = secureConfig;
            _unsecureConfig = unsecureConfig;
        }
        #endregion
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            //Declaring and assigning this in case we decide to use the other config setting.
            _inUseConfig = _unsecureConfig;
            try
            {

                bool traceOn = false;
                if (!string.IsNullOrWhiteSpace(_inUseConfig) && _inUseConfig.ToLower().IndexOf("traceon=true") >= 0)
                {
                    traceOn = true;
                    showVariableInTrace(tracer, traceOn, "TraceON!", "...duh!");
                }

                string message = context.MessageName;
                showVariableInTrace(tracer, traceOn, "Message", message);

                // Get the pre, post, and Target images
                Entity preImageEntity = (context.PreEntityImages != null && context.PreEntityImages.Contains("Target")) ? context.PreEntityImages["Target"] : null;
                Entity postImageEntity = (context.PostEntityImages != null && context.PostEntityImages.Contains("Target")) ? context.PostEntityImages["Target"] : null;
                Entity targetEntity = (context.InputParameters != null && context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity) ? (Entity)context.InputParameters["Target"] : null;

                // If neither of these are present, return with message
                if (postImageEntity == null)
                {
                    showVariableInTrace(tracer, traceOn, "Post Image Entity is NULL, cannot log audit. Check Plugin Registration configuration", string.Empty);
                    return;
                }
                if (targetEntity == null)
                {
                    showVariableInTrace(tracer, traceOn, "Target Image Entity NULL, cannot log audit. Check Plugin Registration", string.Empty);
                    return;
                }
                bool createAudit;
                string oldValue;
                string newValue;
                string fieldName;
                string friendlyFieldName;
                string fieldLabelName;
                string recordName;
                string entityName = postImageEntity.LogicalName.ToLower();

                if (message.ToLower() != "create" && preImageEntity != null)
                {
                    showAttributes(tracer, traceOn, preImageEntity, "preImageEntity");
                }

                showAttributes(tracer, traceOn, postImageEntity, "postImageEntity");
                showAttributes(tracer, traceOn, targetEntity, "targetEntity");


                // ===========================================================
                // Setup the auditEntity variable to be logged
                // TODO: Might want to make this a configurable value?
                Entity auditEntity = new Entity("msftpp_customaudit");
                auditEntity["msftpp_tablename"] = getDisplayNameForEntity(service, entityName);
                auditEntity["msftpp_action"] = message;

                // ===========================================================
                // Setup additional Fields for every entry
                Entity systemUser = retrieveSingleEntityRecordData(new string[] { "systemuserid", "fullname" }, ((EntityReference)targetEntity["modifiedby"]).LogicalName.ToString(), ((EntityReference)targetEntity["modifiedby"]).Id, service);

                // recordName = systemUser["fullname"].ToString();
                recordName = getDisplayNameForEntity(service, entityName) + " :: " + message + " :: ";

                //auditEntity["subject"] = recordName;
                //showVariableInTrace(tracer, traceOn, "recordName ", recordName);

                // --------------------------------------------------------------
                // TODO: This (entity prefix) might have to be pulled from a setting? or in plugin config?
                string entityPrefix = "msftpp_";

                //// Setup entityreference to link to regarding record
                //if (postImageEntity.LogicalName.ToLower().StartsWith(entityPrefix))
                //{
                //    auditEntity[postImageEntity.LogicalName.ToLower() + "id"] = postImageEntity.ToEntityReference();
                //}
                //else
                //{
                //    auditEntity[entityPrefix + postImageEntity.LogicalName.ToLower() + "id"] = postImageEntity.ToEntityReference();
                //}

                auditEntity["regardingobjectid"] = postImageEntity.ToEntityReference();

                // =============================================================================
                // Here we have to check for entries where the 'new' value is blank
                // Basically, the 'old' value will only be in the preImage, 
                // but the 'new' value will NOT be in the postImage
                if (preImageEntity != null)
                {
                    foreach (var field in preImageEntity.Attributes.OrderBy(f => f.Key))
                    {
                        // This is an invalid field to check for audit, skip it
                        if (field.Key.ToString().Contains("onbehalfby")
                            || field.Key.ToString().Contains("yomi"))
                        {
                            continue;
                        }

                        oldValue = null;
                        newValue = null;
                        fieldName = field.Key.ToString();
                        showVariableInTrace(tracer, traceOn, "preImage field.Key", fieldName);
                        showVariableInTrace(tracer, traceOn, "preImage field.Value", field.Value.ToString());

                        // if there are NO postImage attributes, OR the postImage does NOT contain this field
                        // then it must be one we want to log.
                        if (postImageEntity.Attributes == null || !postImageEntity.Attributes.Contains(fieldName))
                        {

                            // the field.Value will be set to the 'type' if its not a string or number
                            // so we have to deal with each type
                            switch (field.Value.ToString())
                            {
                                case "Microsoft.Xrm.Sdk.OptionSetValue":
                                    oldValue = getOptionSetLabelFromOptionSetValue(service, entityName, fieldName, ((OptionSetValue)preImageEntity[fieldName]).Value);
                                    break;

                                case "Microsoft.Xrm.Sdk.EntityReference":
                                    oldValue = ((EntityReference)preImageEntity[fieldName]).Name;
                                    break;

                                case "Microsoft.Xrm.Sdk.Money":
                                    oldValue = ((Money)preImageEntity[fieldName]).Value.ToString();
                                    break;

                                default:
                                    oldValue = string.IsNullOrWhiteSpace(preImageEntity[fieldName].ToString()) ? null : preImageEntity[fieldName].ToString();
                                    break;
                            }

                            auditEntity["msftpp_oldvalue"] = oldValue != null && oldValue.Length > 3999 ? oldValue.Substring(0, 4000) : oldValue;
                            auditEntity["msftpp_newvalue"] = newValue != null && newValue.Length > 3999 ? newValue.Substring(0, 4000) : newValue;

                            // Set the Field Name to the 'friendly' label
                            friendlyFieldName = getDisplayLabelFromLogicalName(service, postImageEntity.LogicalName.ToLower(), fieldName);
                            auditEntity["msftpp_columnname"] = friendlyFieldName;
                            recordName = getDisplayNameForEntity(service, entityName) + " :: " + message + " :: " + friendlyFieldName;
                            auditEntity["subject"] = recordName;
                            showVariableInTrace(tracer, traceOn, "recordName ", recordName);
                            Guid newActivityId = service.Create(auditEntity);
                            EntityReference e1 = new EntityReference("msftpp_customaudit", newActivityId);
                            showVariableInTrace(tracer, traceOn, " *********************** Deactivate Record!! ", e1.Id.ToString());
                            deactivateRecord(service, e1);
                        }

                    }
                }
                else
                {
                    showVariableInTrace(tracer, traceOn, "preImage entity is NULL ", string.Empty);
                }

                // Null out values to be safe
                auditEntity["msftpp_columnname"] = null;
                auditEntity["msftpp_oldvalue"] = null;
                auditEntity["msftpp_newvalue"] = null;
                auditEntity["subject"] = null;

                // =============================================================================
                // Work on post image fields
                showVariableInTrace(tracer, traceOn, "postImage fields\r\n=============================", string.Empty);
                foreach (var field in postImageEntity.Attributes.OrderBy(f => f.Key))
                {
                    auditEntity["msftpp_columnname"] = null;
                    auditEntity["msftpp_oldvalue"] = null;
                    auditEntity["msftpp_newvalue"] = null;
                    auditEntity["subject"] = null;

                    showVariableInTrace(tracer, traceOn, "postImage field.Key", field.Key.ToString());
                    showVariableInTrace(tracer, traceOn, "postImage field.Value", field.Value.ToString());

                    // This is an invalid field to check for audit, skip it
                    if (field.Key.ToString().Contains("onbehalfby")
                        || field.Key.ToString().Contains("yomi"))
                    {
                        continue;
                    }

                    oldValue = null;
                    newValue = null;
                    fieldName = field.Key.ToString();
                    fieldLabelName = getDisplayLabelFromLogicalName(service, postImageEntity.LogicalName.ToLower(), fieldName);


                    // Set the default to potentially be used later in the logic
                    createAudit = true;

                    // the field.Value will be set to the 'type' if its not a string or number
                    // so we have to deal with each type
                    switch (field.Value.ToString())
                    {
                        case "Microsoft.Xrm.Sdk.OptionSetValue":
                            if (preImageEntity != null && preImageEntity.Contains(fieldName))
                            {
                                oldValue = getOptionSetLabelFromOptionSetValue(service, entityName, fieldName, ((OptionSetValue)preImageEntity[fieldName]).Value);
                            }
                            else
                            {
                                oldValue = null;
                            }

                            newValue = getOptionSetLabelFromOptionSetValue(service, entityName, fieldName, ((OptionSetValue)postImageEntity[fieldName]).Value, tracer, traceOn);
                            break;

                        case "Microsoft.Xrm.Sdk.EntityReference":
                            if (preImageEntity != null && preImageEntity.Contains(fieldName))
                            {
                                oldValue = ((EntityReference)preImageEntity[fieldName]).Name;
                            }
                            else
                            {
                                oldValue = null;
                            }

                            newValue = ((EntityReference)postImageEntity[fieldName]).Name;
                            break;

                        case "Microsoft.Xrm.Sdk.Money":
                            if (preImageEntity != null && preImageEntity.Contains(fieldName))
                            {
                                oldValue = ((Money)preImageEntity[fieldName]).Value.ToString();
                            }
                            else
                            {
                                oldValue = null;
                            }

                            newValue = ((Money)postImageEntity[fieldName]).Value.ToString();
                            break;


                        default:
                            if (preImageEntity != null && preImageEntity.Contains(fieldName))
                            {
                                oldValue = string.IsNullOrWhiteSpace(preImageEntity.Attributes[fieldName].ToString()) ? null : preImageEntity.Attributes[fieldName].ToString();
                            }
                            else
                            {
                                oldValue = null;
                            }

                            newValue = string.IsNullOrWhiteSpace(field.Value.ToString()) ? " " : field.Value.ToString();
                            break;
                    }

                    // If they match, we are not logging it ... continue;
                    if (oldValue == newValue)
                    {
                        showVariableInTrace(tracer, traceOn, " Values Match, Skipping!!", string.Empty);
                        continue;

                    }

                    // TODO: This might be a configuration setting option?
                    // If only case changes, we are not logging it ... continue;
                    if (oldValue.ToLower() == newValue.ToLower())
                    {
                        showVariableInTrace(tracer, traceOn, " Values Match, only case Changed, Skipping!!", string.Empty);
                        continue;
                    }

                    // Set the Field Name to the 'friendly' label
                    friendlyFieldName = getDisplayLabelFromLogicalName(service, postImageEntity.LogicalName.ToLower(), fieldName);
                    auditEntity["msftpp_columnname"] = friendlyFieldName;
                    recordName = getDisplayNameForEntity(service, entityName) + " :: " + message + " :: " + friendlyFieldName;
                    auditEntity["subject"] = recordName;
                    showVariableInTrace(tracer, traceOn, "recordName ", recordName);

                    // Old/New only hold 4000 characters, so we have to trim if it's longer than 4k
                    auditEntity["msftpp_oldvalue"] = oldValue != null && oldValue.Length > 3999 ? oldValue.Substring(0, 4000) : oldValue;
                    auditEntity["msftpp_newvalue"] = newValue != null && newValue.Length > 3999 ? newValue.Substring(0, 4000) : newValue;

                    if (createAudit)
                    {
                        auditEntity["msftpp_columnname"] = getDisplayLabelFromLogicalName(service, postImageEntity.LogicalName.ToLower(), fieldName);
                        showVariableInTrace(tracer, traceOn, " *********************** CREATE!! ", string.Empty);
                        Guid newActivityId2 = service.Create(auditEntity);
                        EntityReference e1 = new EntityReference("msftpp_customaudit", newActivityId2);
                        showVariableInTrace(tracer, traceOn, " *********************** Deactivate Record!! ", e1.Id.ToString());
                        deactivateRecord(service, e1);
                    }
                }




            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }

        private static void showAttributes(ITracingService tracer, bool traceOn, Entity e, string eName)
        {
            StringBuilder sb = new StringBuilder();
            showVariableInTrace(tracer, traceOn, eName + " Attributes count?", string.IsNullOrWhiteSpace(e.Attributes.Count.ToString()) ? "NULL Count" : "Count :: " + e.Attributes.Count.ToString());
            foreach (var field in e.Attributes)
            {

                if (field.Key != null && field.Key.ToString().Contains("onbehalfby"))
                {
                    sb.AppendFormat("Skipping!! ", " ");
                    continue;
                }
                if (field.Key != null)
                {
                    sb.AppendFormat("Key = {0}", string.IsNullOrWhiteSpace(field.Key.ToString()) ? "NULL" : field.Key.ToString());
                }
                else
                {
                    sb.AppendFormat("Key = {0}", "NULL");
                }

                if (field.Value != null)
                {
                    sb.AppendFormat(", Value = {0}", string.IsNullOrWhiteSpace(field.Value.ToString()) ? "NULL" : field.Value.ToString());
                }
                else
                {
                    sb.AppendFormat(", Value = {0}", "NULL");
                }

                showVariableInTrace(tracer, traceOn, "     ", sb.ToString());
                sb.Clear();
            }
            showVariableInTrace(tracer, traceOn, "==============================", string.Empty);

        }
        private static void showFormattedValues(ITracingService tracer, bool traceOn, Entity e, string eName)
        {
            StringBuilder sb = new StringBuilder();
            showVariableInTrace(tracer, traceOn, eName + " FormattedValues count?", string.IsNullOrWhiteSpace(e.FormattedValues.Count.ToString()) ? "NULL Count" : "Count :: " + e.FormattedValues.Count.ToString());
            foreach (var field in e.FormattedValues)
            {
                if (field.Key != null && field.Key.ToString().Contains("onbehalfby"))
                {
                    continue;
                }

                if (field.Key != null)
                {
                    sb.AppendFormat("Key = {0}", string.IsNullOrWhiteSpace(field.Key.ToString()) ? "NULL" : field.Key.ToString());
                }
                else
                {
                    sb.AppendFormat("Key = {0}", "NULL");
                }

                if (field.Value != null)
                {
                    sb.AppendFormat(", Value = {0} ", string.IsNullOrWhiteSpace(field.Value.ToString()) ? "NULL" : field.Value.ToString());
                }
                else
                {
                    sb.AppendFormat(", Value = {0}", "NULL");
                }

                showVariableInTrace(tracer, traceOn, "     ", sb.ToString());
                sb.Clear();
            }
            showVariableInTrace(tracer, traceOn, "==============================", string.Empty);
        }

        public static void showVariableInTrace(ITracingService tracer, bool isTraceOn, string varName, string varValue)
        {
            if (isTraceOn)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendFormat("{0} -- ", DateTime.Now.ToString("HH.mm.ss.ffffff"));
                sb.AppendFormat("{0}", varName);
                if (!string.IsNullOrWhiteSpace(varValue))
                {
                    sb.AppendFormat(" :: {0}", varValue);
                }
                // sb.AppendLine();
                tracer.Trace(sb.ToString());
            }
        }

        public static Entity retrieveSingleEntityRecordData(string[] fieldArray, string entityLogicalName, Guid entityId, IOrganizationService service)
        {
            RetrieveRequest requestRecord = new RetrieveRequest();
            requestRecord.ColumnSet = new ColumnSet(fieldArray);
            requestRecord.Target = new EntityReference(entityLogicalName, entityId);
            Entity returnEntity = (Entity)((RetrieveResponse)service.Execute(requestRecord)).Entity;
            return returnEntity;
        }

        public static string getOptionSetLabelFromOptionSetValue(IOrganizationService service, string entityName, string schemaFieldName, int OptionSetValue, ITracingService tracer = null, bool traceOn = false)
        {
            string returnValue = string.Empty;

            RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = schemaFieldName,
                RetrieveAsIfPublished = true
            };
            RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
            OptionMetadata[] optionList = null;

            // StateAttributeMetadata
            switch (schemaFieldName)
            {
                case "statuscode":
                    StatusAttributeMetadata a = (StatusAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
                    optionList = a.OptionSet.Options.ToArray();
                    break;

                case "statecode":
                    StateAttributeMetadata b = (StateAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;

                    optionList = b.OptionSet.Options.ToArray();
                    break;

                default:
                    PicklistAttributeMetadata c = (PicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
                    optionList = c.OptionSet.Options.ToArray();

                    break;

            }
            foreach (OptionMetadata o in optionList)
            {
                if (o.Value == OptionSetValue)
                {
                    returnValue = o.Label.LocalizedLabels.FirstOrDefault(e => e.LanguageCode == 1033).Label.ToString().Trim();
                    break;
                }
            }


            return returnValue;

        }

        public static string getDisplayLabelFromLogicalName(IOrganizationService service, string entityName, string schemaFieldName)
        {
            string returnValue = string.Empty;

            RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = schemaFieldName,
                RetrieveAsIfPublished = true
            };
            RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
            AttributeMetadata retrievedAttributeMetadata = (AttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
            returnValue = retrievedAttributeMetadata.DisplayName.UserLocalizedLabel.Label;

            return returnValue;

        }

        public static string getDisplayNameForEntity(IOrganizationService service, string entityName)
        {
            string returnValue = string.Empty;

            RetrieveEntityRequest req = new RetrieveEntityRequest
            {
                LogicalName = entityName,
                EntityFilters = EntityFilters.Entity,
                RetrieveAsIfPublished = true
            };
            RetrieveEntityResponse resp = (RetrieveEntityResponse)service.Execute(req);
            EntityMetadata m = (EntityMetadata)resp.EntityMetadata;
            returnValue = m.DisplayName.UserLocalizedLabel.Label.ToString();

            return returnValue;

        }

        public static void deactivateRecord(IOrganizationService service, EntityReference entityRef)
        {

            SetStateRequest state = new SetStateRequest();
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);
            state.EntityMoniker = entityRef;
            SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
        }
    }
}
