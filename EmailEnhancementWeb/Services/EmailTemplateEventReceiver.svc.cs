using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;

namespace EmailEnhancementWeb.Services
{
    public class EmailTemplateEventReceiver : IRemoteEventService
    {
        
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult result = new SPRemoteEventResult();

            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    clientContext.Load(clientContext.Web);
                    clientContext.ExecuteQuery();


                    if (properties.EventType == SPRemoteEventType.ItemAdded)
                    {

                        itemaddevent(properties);
                    }

                    else if (properties.EventType == SPRemoteEventType.ItemUpdated)
                    {
                        itemaddevent(properties);

                    }
                }
            }

            return result;
        }

        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            throw new NotImplementedException();
        }

        public static void itemaddevent(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {


                if (clientContext != null)
                {
                    clientContext.Load(clientContext.Web);
                    clientContext.ExecuteQuery();
                    List questionChoice = clientContext.Web.Lists.GetByTitle("Question Choice");
                    List nominations = clientContext.Web.Lists.GetByTitle("Nomination");
                    List test = clientContext.Web.Lists.GetByTitle("Test");

                    string eventlist = properties.ItemEventProperties.ListTitle;
                    ListItem item = clientContext.Web.Lists.GetByTitle("Email Template").GetItemById(
                    properties.ItemEventProperties.ListItemId);
                    clientContext.Load(test);
                    clientContext.Load(item);
                    clientContext.ExecuteQuery();
                    FieldLookupValue group = (FieldLookupValue)item["Choice_x0020_ID"];
                    string choiceID = group.LookupValue;
                    string templateType = Convert.ToString(item["Template_x0020_Type"]);
                    string body = Convert.ToString(item["Body"]);
                    string ImageUrl = Convert.ToString(item["Image_x0020_Path"]);
                    string subject = Convert.ToString(item["Subject"]);

                    ListItemCreationInformation cInfo = new ListItemCreationInformation();
                    ListItem newItem = test.AddItem(cInfo);
                    string text = choiceID + body;
                    newItem["Title"] = text;
                    newItem.Update();
                    clientContext.ExecuteQuery();


                    //        CamlQuery query = new CamlQuery();
                    //        query.ViewXml = string.Format("<View><Query>" +
                    //                                    "<Where>" +
                    //                                            "<Eq><FieldRef Name='Title' />" +
                    //                                            "<Value Type='Text'>{0}</Value></Eq>" +
                    //                                        "</Where></Query><RowLimit>500</RowLimit></View>", choiceID);

                    //        Microsoft.SharePoint.Client.ListItemCollection spItems = questionChoice.GetItems(query);

                    //        clientContext.Load(spItems);
                    //        clientContext.ExecuteQuery();

                    //        foreach (ListItem spItem in spItems)
                    //        {
                    //            string choiceEN = Convert.ToString(spItem["Choice_x0020_EN"]);
                    //            updateNominations(clientContext, nominations, templateType, choiceEN, body, subject, ImageUrl);

                    //        }
                }
            }
        }
    }
}
