using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System.ServiceModel;
using System.Reflection;
using System.Runtime.Serialization;
using System;
using System.Collections.Generic;

namespace EmailEnhancementWeb.Services
{
    public class AppEventReceiver : IRemoteEventService
    {
        /// <summary>
        /// Handles app events that occur after the app is installed or upgraded, or when app is being uninstalled.
        /// </summary>
        /// <param name="properties">Holds information about the app event.</param>
        /// <returns>Holds information returned from the app event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult result = new SPRemoteEventResult();

            switch (properties.EventType)
            {
                case SPRemoteEventType.AppInstalled: HandleAppInstalled(properties);
                    break;

                //  case SPRemoteEventType.app: HandleAppUninstalled(properties); break;

                //   case SPRemoteEventType.ItemAdded: HandleItemAdded(properties); break;
            }
            if (properties.EventType == SPRemoteEventType.AppUninstalling)
            {
                using (ClientContext clientContext = TokenHelper.CreateAppEventClientContext(properties, false))
                {
                    var list = clientContext.Web.Lists.GetByTitle("Email Template");
                    clientContext.Load(list);
                    clientContext.ExecuteQuery();
                    EventReceiverDefinitionCollection erdc = list.EventReceivers;
                    clientContext.Load(erdc);
                    clientContext.ExecuteQuery();
                    List<EventReceiverDefinition> toDelete = new List<EventReceiverDefinition>();
                    foreach (EventReceiverDefinition erd in erdc)
                    {
                        if (erd.ReceiverName == "EmailTemplateEventReceiver")
                        {
                            toDelete.Add(erd);
                        }
                    }
                    //Delete the remote event receiver from the list, when the app gets uninstalled
                    foreach (EventReceiverDefinition item in toDelete)
                    {
                        item.DeleteObject();
                        clientContext.ExecuteQuery();
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// This method is a required placeholder, but is not used by app events.
        /// </summary>
        /// <param name="properties">Unused.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            throw new NotImplementedException();
        }

        private void HandleAppInstalled(SPRemoteEventProperties properties)
        {

            using (ClientContext clientContext = TokenHelper.CreateAppEventClientContext(properties, false))
            {
                if (clientContext != null)
                {
                    bool rerExists = false;
                    List myList = clientContext.Web.Lists.GetByTitle("Email Template");
                    clientContext.Load(myList, p => p.EventReceivers);
                    clientContext.ExecuteQuery();

                    foreach (var rer in myList.EventReceivers)
                    {
                        if (rer.ReceiverName == "EmailTemplateEventReceiver")
                        {
                            rerExists = true;
                        }
                    }

                    if (!rerExists)
                    {
                        EventReceiverDefinitionCreationInformation receiver = new EventReceiverDefinitionCreationInformation();
                        receiver.EventType = EventReceiverType.ItemAdded;
                        receiver.ReceiverName = "EmailTemplateEventReceiver";
                        receiver.ReceiverClass ="EmailTemplateEventReceiver";
                        receiver.ReceiverAssembly = Assembly.GetExecutingAssembly().FullName;
                        OperationContext op = OperationContext.Current;
                        string str = op.RequestContext.RequestMessage.Headers.To.ToString();
                        string remoterersvc = str.Replace(str.Substring(str.IndexOf("AppEventReceiver")), "EmailTemplateEventReceiver.svc");
                        receiver.ReceiverUrl = remoterersvc;
                        receiver.Synchronization = EventReceiverSynchronization.Synchronous;
                        myList.EventReceivers.Add(receiver);
                        //receiver = new EventReceiverDefinitionCreationInformation();
                        //receiver.EventType = EventReceiverType.ItemUpdated;
                        //clientContext.ExecuteQuery();
                        //receiver.ReceiverName = "EmailTemplateEventReceiver";
                        //receiver.ReceiverClass = "EmailTemplateEventReceiver";
                        //receiver.ReceiverAssembly = Assembly.GetExecutingAssembly().FullName;
                        //op = OperationContext.Current;
                        //str = op.RequestContext.RequestMessage.Headers.To.ToString();
                        //remoterersvc = str.Replace(str.Substring(str.IndexOf("AppEventReceiver")), "EmailTemplateEventReceiver.svc");
                        //receiver.ReceiverUrl = remoterersvc;
                        //receiver.Synchronization = EventReceiverSynchronization.Synchronous;
                        //myList.EventReceivers.Add(receiver);
                        //receiver = new EventReceiverDefinitionCreationInformation();
                        //receiver.EventType = EventReceiverType.ItemAdding;
                        //clientContext.ExecuteQuery();
                        //receiver.ReceiverName = "EmailTemplateEventReceiver";
                        //receiver.ReceiverClass = "EmailTemplateEventReceiver";
                        //receiver.ReceiverAssembly = Assembly.GetExecutingAssembly().FullName;
                        //op = OperationContext.Current;
                        //str = op.RequestContext.RequestMessage.Headers.To.ToString();
                        //remoterersvc = str.Replace(str.Substring(str.IndexOf("AppEventReceiver")), "EmailTemplateEventReceiver.svc");
                        //receiver.ReceiverUrl = remoterersvc;
                        //receiver.Synchronization = EventReceiverSynchronization.Synchronous;
                        //myList.EventReceivers.Add(receiver); 
                        clientContext.ExecuteQuery();
                    }

                }
            }
        }

    }
}
