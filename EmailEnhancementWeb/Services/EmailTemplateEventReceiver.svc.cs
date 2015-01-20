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



                    /* Karthik code
                    clientContext.Load(clientContext.Web);
                    clientContext.ExecuteQuery();
                    List imageLibrary = clientContext.Web.Lists.GetByTitle("Test");
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem oListItem = imageLibrary.GetItemById(14);
           
                    oListItem["Title"] = "tcs:" ;
                    oListItem.Update();
                    clientContext.ExecuteQuery();
                     * */


                    clientContext.Load(clientContext.Web);
                    clientContext.ExecuteQuery();
                    List questionChoice = clientContext.Web.Lists.GetByTitle("Question Choice");
                    List nominations = clientContext.Web.Lists.GetByTitle("Nomination");

                    string eventlist = properties.ItemEventProperties.ListTitle;
                    ListItem item = clientContext.Web.Lists.GetByTitle("Email Template").GetItemById(
                    properties.ItemEventProperties.ListItemId);
                    clientContext.Load(item);
                    clientContext.ExecuteQuery();
                    FieldLookupValue group = (FieldLookupValue)item["Choice_x0020_ID"];
                    string choiceID = group.LookupValue;
                    string templateType = Convert.ToString(item["Template_x0020_Type"]);
                    string body = Convert.ToString(item["Body"]);
                    string ImageUrl = Convert.ToString(item["Image_x0020_Path"]);
                    string subject = Convert.ToString(item["Subject"]);

                    //ListItemCreationInformation cInfo = new ListItemCreationInformation();
                    //ListItem newItem = test.AddItem(cInfo);
                    //string text = choiceID + body;
                    //newItem["Title"] = text;
                    //newItem.Update();
                    //clientContext.ExecuteQuery();


                    CamlQuery query = new CamlQuery();
                    query.ViewXml = string.Format("<View><Query>" +
                                                "<Where>" +
                                                        "<Eq><FieldRef Name='Title' />" +
                                                        "<Value Type='Text'>{0}</Value></Eq>" +
                                                    "</Where></Query><RowLimit>500</RowLimit></View>", choiceID);

                    Microsoft.SharePoint.Client.ListItemCollection spItems = questionChoice.GetItems(query);

                    clientContext.Load(spItems);
                    clientContext.ExecuteQuery();

                    foreach (ListItem spItem in spItems)
                    {
                        string choiceEN = Convert.ToString(spItem["Choice_x0020_EN"]);
                        updateNominations(clientContext, nominations, templateType, choiceEN, body, subject, ImageUrl);

                    }
                }
            }
        }

        public static void updateNominations(ClientContext clientContext, List nomination, string templateType, string choiceEN, string body, string subject, string ImageUrl)
        {
            try
            {
                CamlQuery nominations = new CamlQuery();
                nominations.ViewXml = string.Format("<View><Query>" +
                                            "<Where>" +
                                                    "<Eq><FieldRef Name='Business_x0020_Unit' />" +
                                                    "<Value Type='Text'>{0}</Value></Eq>" +
                                                "</Where></Query><RowLimit>500</RowLimit></View>", choiceEN);

                Microsoft.SharePoint.Client.ListItemCollection nItems = nomination.GetItems(nominations);
                clientContext.Load(nItems);
                clientContext.ExecuteQuery();

                foreach (ListItem nomItem in nItems)
                {
                    List<FieldUserValue> To = new List<FieldUserValue>();
                    List<FieldUserValue> CC = new List<FieldUserValue>();
                    FieldUserValue mgr = new FieldUserValue();
                    FieldUserValue n = new FieldUserValue();
                    FieldUserValue[] coords = null;
                    FieldUserValue[] nom = null;
                    string nominationId = nomItem.Id.ToString();
                    string formattedBody = string.Empty;

                    //updates the nomination list based on the email templateType value passed
                    switch (templateType)
                    {
                        case "Manager":

                            //formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            formattedBody = body;
                            nomItem["Manager_x0020_Email"] = formattedBody;
                            mgr = (FieldUserValue)nomItem["Approving_x0020_Manager"];
                            mgr = (FieldUserValue)nomItem["Approving_x0020_Manager"];
                            To.Add(mgr);



                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);

                            break;

                        case "ManagerRetractNotify":
                            //formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            formattedBody = body;
                            nomItem["Mgr_x0020_Draft_x0020_Email"] = formattedBody;
                            n = (FieldUserValue)nomItem["Approving_x0020_Manager"];
                            if (n != null)
                            {
                                To.Add(n);
                            }
                            else
                            {
                                return;
                            }

                            n = (FieldUserValue)nomItem["Nominator"];
                            CC.Add(n);

                            n = (FieldUserValue)nomItem["Submitter"];
                            CC.Add(n);
                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);
                            break;

                        case "ManagerRejected":
                            //formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            formattedBody = body;
                            nomItem["Mgr_x0020_Reject_x0020_Email"] = formattedBody;
                            n = (FieldUserValue)nomItem["Nominator"];
                            To.Add(n);

                            n = (FieldUserValue)nomItem["Approving_x0020_Manager"];
                            CC.Add(n);

                            n = (FieldUserValue)nomItem["Submitter"];
                            CC.Add(n);

                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);
                            break;

                        case "ManagerReminder":
                            //formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            formattedBody = body;
                            nomItem["Mgr_x0020_Remind_x0020_Email"] = formattedBody;
                            mgr = (FieldUserValue)nomItem["Approving_x0020_Manager"];
                            To.Add(mgr);

                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);

                            DateTime start = (DateTime)nomItem["Review_x0020_Start"];
                            break;

                        case "Submitted":
                            //formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            nomItem["Submitted_x0020_Email"] = formattedBody;
                            n = (FieldUserValue)nomItem["Nominator"];
                            To.Add(n);

                            n = (FieldUserValue)nomItem["Submitter"];
                            CC.Add(n);

                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);
                            break;

                        case "NomineeFailed":
                            //formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            formattedBody = body;
                            nomItem["Failure_x0020_Email"] = formattedBody;
                            nom = (FieldUserValue[])nomItem["Nominees"];
                            To.AddRange(nom);

                            n = (FieldUserValue)nomItem["Nominator"];
                            CC.Add(n);

                            n = (FieldUserValue)nomItem["Submitter"];
                            CC.Add(n);

                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);
                            break;

                        case "NomineeSelected":
                            //formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            formattedBody = body;
                            nomItem["Success_x0020_Email"] = formattedBody;
                            nom = (FieldUserValue[])nomItem["Nominees"];
                            To.AddRange(nom);

                            n = (FieldUserValue)nomItem["Nominator"];
                            CC.Add(n);

                            n = (FieldUserValue)nomItem["Submitter"];
                            CC.Add(n);
                            break;

                        case "Reviewer":
                            //formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            formattedBody = body;
                            nomItem["Reviewer_x0020_Email"] = formattedBody;
                            nom = (FieldUserValue[])nomItem["Reviewers"];
                            To.AddRange(nom);

                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);
                            break;

                        case "ReviewerReminder":
                            //formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            formattedBody = body;
                            nomItem["Reminder_x0020_Email"] = formattedBody;
                            nom = (FieldUserValue[])nomItem["Reviewers"];
                            To.AddRange(nom);

                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);
                            break;

                        case "NominatorNotify":
                            //formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            formattedBody = body;
                            nomItem["Nominator_x0020_Email"] = formattedBody;

                            n = (FieldUserValue)nomItem["Nominator"];
                            To.Add(n);

                            n = (FieldUserValue)nomItem["Submitter"];
                            CC.Add(n);

                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);
                            break;
                    }

                    nomItem.Update();
                    clientContext.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                // Karthik code
                clientContext.Load(clientContext.Web);
                clientContext.ExecuteQuery();
                List imageLibrary = clientContext.Web.Lists.GetByTitle("Test");
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = imageLibrary.GetItemById(14);

                oListItem["Title"] = "tcs:" + ex.ToString();
                oListItem.Update();
                clientContext.ExecuteQuery();

            }

            //updateEmailSendList(clientContext, nominationId, nomItem, subject, formattedBody, To, CC, templateType);
        }
    }
}