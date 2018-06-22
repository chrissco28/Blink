using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System.ServiceModel.Description;



namespace BlinkUX.Plugins
{
    public partial class SurveyAnswerPlugin : IPlugin
    {
        private OrganizationServiceProxy _serviceProxy;
        private IOrganizationService _service;
        private ITracingService _tracingService;
        private IOrganizationServiceFactory _serviceFactory;

        public void Execute(IServiceProvider serviceProvider)
        {
            Entity image;
         
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            
            
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            _serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            _service = _serviceFactory.CreateOrganizationService(context.UserId);


            //make sure we don't get into a loop
            if (context.Depth > 1)
                return;

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity entity = (Entity)context.InputParameters["Target"];

                //get the images. for anything other than delete, get the post updat
                if (context.MessageName.ToString().ToLower() == "delete")
                {
                    image = (Entity)context.PreEntityImages["image"];
                }
                else
                {
                    image = (Entity)context.PostEntityImages["image"];
                }

                try
                {
                    

                    //fire only on the ClickDimensions email event plugin
                    if (entity.LogicalName != "cdi_surveyanswer")
                    {
                        throw new Exception("Plugin not registered properly. Please check configuration.");
                    }

                    SetNumberofAnswers(entity, image);
                    
                }
                catch (Exception ex)
                {
                    _tracingService.Trace(ex.Message);
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
        private void SetNumberofAnswers(Entity entity, Entity image)
        {
            int responseCount = 0;
            ColumnSet cols = new ColumnSet(true);

            string cdi_postedsurveyid = image.GetAttributeValue<EntityReference>("cdi_postedsurveyid").Id.ToString();

            //get the atttribute from the map
            //make sure the posted survey look up is in the map
            if (!image.Attributes.Contains("cdi_postedsurveyid"))
            {
                throw new Exception("missing posted survey field in image");
            }

            string fetchXML = string.Format(@"<fetch mapping='logical' aggregate='true' version='1.0'>
                            <entity name='cdi_surveyanswer'>
                                <attribute name ='cdi_surveyanswerid' alias='responseCount' aggregate='count' />
                                <filter><condition attribute='cdi_postedsurveyid' operator='eq' value='{0}' /></filter></entity></fetch>", cdi_postedsurveyid);
            _tracingService.Trace(fetchXML);

            EntityCollection responseCountCollection = _service.RetrieveMultiple(new FetchExpression(fetchXML));

            foreach (var c in responseCountCollection.Entities)
            {
                //_tracingService.Trace("Response Count found: " + responseCount.ToString());
                responseCount = ((int)((AliasedValue)c["responseCount"]).Value);
            }

            //update the posted survey with the latest answer count
            Entity cdi_postedsurvey = _service.Retrieve("cdi_postedsurvey", Guid.Parse(cdi_postedsurveyid), cols);
            cdi_postedsurvey["bnk_responsecount"] = responseCount;
            _service.Update(cdi_postedsurvey);

            _tracingService.Trace("cdi_postedsurvey updated with response count");
        }
    }
}
 