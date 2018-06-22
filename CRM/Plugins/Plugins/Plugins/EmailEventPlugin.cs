using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;


namespace BlinkUX.Plugins
{
    public partial class EmailEvent : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext) serviceProvider.GetService(typeof(IPluginExecutionContext));


            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity entity = (Entity)context.InputParameters["Target"];
                
                try
                {
                    //fire only on the ClickDimensions email event plugin
                    if (entity.LogicalName != "cdi_emailevent")
                    {
                        throw new Exception("Plugin not registered properly. Please check configuration.");
                    }

                    SetNPSSurveyLookup(entity, context, tracingService);
                }
                catch(Exception ex)
                {
                    tracingService.Trace(ex.Message);
                    throw new Exception(ex.Message);
                }
            }
        }

        /// <summary>
        /// Get the URL passed in from the NPS Survey. The NPS Survey PK is a query string parameter
        /// Retrieve the NPS Survey ID and set it on the Email Event record
        /// This plugin is designed for create & update
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="context"></param>
        /// <param name="tracingService"></param>
        private void SetNPSSurveyLookup(Entity entity, IPluginExecutionContext context, ITracingService tracingService)
        {
            Guid nps_surveyid;
            string surveyid = string.Empty;
            string cdi_url = string.Empty;

            // make sure the URL is passed in
            if (entity.Attributes.Contains("cdi_url"))
            {
                cdi_url = entity.Attributes["cdi_url"].ToString();

                tracingService.Trace("cdi_url:" + entity.Attributes["cdi_url"].ToString());

                //look specifically for the key value set
                if (cdi_url.Contains("cp_npssurveyid="))
                {
                    tracingService.Trace("cp_npssurveyid found");

                    //find the place in the URL where the key is
                    //remove the key and get the GUID
                    surveyid = cdi_url.Substring(cdi_url.IndexOf("cp_npssurveyid="), 51).Remove(0, 15);

                    tracingService.Trace("cp_npssurveyid:" + surveyid);

                    //if there is value retrieved and it is GUID 
                    //then set the lookup on the entity
                    if (Guid.TryParse(surveyid, out nps_surveyid))
                    {
                        entity["new_nspsurveyid"] = new EntityReference("cp_npssurvey", nps_surveyid);
                        tracingService.Trace("new_nspsurveyid set");
                    }
                }
            }
        }
    }
}