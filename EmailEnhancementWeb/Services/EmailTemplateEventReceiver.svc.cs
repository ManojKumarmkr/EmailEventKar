using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Configuration;
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
                        try
                        {
                            itemaddevent(properties);
                        }
                        catch (Exception ex)
                        {
                            Trace.TraceError("Failed to call itemaddevent method", ex);
                            Console.Error.WriteLine("Failed to itemaddevent method" + ex);
                        }

                    }

                    else if (properties.EventType == SPRemoteEventType.ItemUpdated)
                    {
                        try
                        {
                            itemaddevent(properties);
                        }

                        catch (Exception ex)
                        {
                            Trace.TraceError("Failed to call itemaddevent method", ex);
                            Console.Error.WriteLine("Failed to itemaddevent method" + ex);
                        }
                    }
                }
            }

            return result;
        }

        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            throw new NotImplementedException();
        }

        //called on itemupdated and itemadded event
        public static void itemaddevent(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {

                if (clientContext != null)
                {

                    try
                    {

                        clientContext.Load(clientContext.Web);
                        clientContext.ExecuteQuery();
                        List questionChoice = clientContext.Web.Lists.GetByTitle("Question Choice");
                        List nominations = clientContext.Web.Lists.GetByTitle("Nomination");

                        //string eventlist = properties.ItemEventProperties.ListTitle;
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
                        bool sendMail = Convert.ToBoolean(item["Send_x0020_Mail"]);

                        CamlQuery query = new CamlQuery();
                        query.ViewXml = string.Format("<View><Query>" +
                                                    "<Where>" +
                                                            "<Eq><FieldRef Name='Title' />" +
                                                            "<Value Type='Text'>{0}</Value></Eq>" +
                                                        "</Where></Query><RowLimit>500</RowLimit></View>", choiceID);

                        Trace.TraceInformation("Query Question choice list based on BU/BG query:" + query.ViewXml);
                        Console.Out.WriteLine("Query Question choice list based on BU/BG query:" + query.ViewXml);
                        Microsoft.SharePoint.Client.ListItemCollection spItems = questionChoice.GetItems(query);

                        clientContext.Load(spItems);
                        clientContext.ExecuteQuery();

                        foreach (ListItem spItem in spItems)
                        {
                            string choiceEN = Convert.ToString(spItem["Choice_x0020_EN"]);
                            Trace.TraceInformation("BU/BG value to query nominations:" + choiceEN);
                            Console.Out.WriteLine("BU/BG value to query nominations:" + choiceEN);
                            updateNominations(clientContext, nominations, templateType, choiceEN, body, subject, ImageUrl, sendMail);

                        }
                    }
                    catch (Exception ex)
                    {
                        Trace.TraceError("Failed to update nomination and emailsend list", ex);
                        Console.Error.WriteLine("Failed to update nomination and emailsend list" + ex);
                    }
                }
            }
        }

        //updates the nomination list based on BU/BG and template type
        public static void updateNominations(ClientContext clientContext, List nomination, string templateType, string choiceEN, string body, string subject, string ImageUrl, bool sendMail)
        {
            try
            {
                CamlQuery nominations = new CamlQuery();
                nominations.ViewXml = string.Format("<View><Query>" +
                                            "<Where>" +
                                                    "<Eq><FieldRef Name='Business_x0020_Unit' />" +
                                                    "<Value Type='Text'>{0}</Value></Eq>" +
                                                "</Where></Query><RowLimit>500</RowLimit></View>", choiceEN);

                Trace.TraceInformation("Query nominations list based on BU/BG query:" + nominations.ViewXml);
                Console.Out.WriteLine("Query nominations list based on BU/BG query:" + nominations.ViewXml);
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

                    Trace.TraceInformation("NominationID =" +nominationId+ "Template Type =" +templateType);
                    Console.Out.WriteLine("NominationID =" + nominationId + "Template Type =" + templateType);

                    //updates the nomination list based on the email templateType value passed
                    switch (templateType)
                    {
                        case "Manager":

                            formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            nomItem["Manager_x0020_Email"] = formattedBody;
                            mgr = (FieldUserValue)nomItem["Approving_x0020_Manager"];
                            To.Add(mgr);

                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);

                            break;

                        case "Manager Retract Notify":
                            formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
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

                        case "Manager Rejected":
                            formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
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

                        case "Manager Reminder":
                            formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            nomItem["Mgr_x0020_Remind_x0020_Email"] = formattedBody;
                            mgr = (FieldUserValue)nomItem["Approving_x0020_Manager"];
                            To.Add(mgr);

                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);

                            break;

                        case "Submitted":
                            formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            nomItem["Submitted_x0020_Email"] = formattedBody;
                            n = (FieldUserValue)nomItem["Nominator"];
                            To.Add(n);

                            n = (FieldUserValue)nomItem["Submitter"];
                            CC.Add(n);

                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);
                            break;

                        case "Nominee Failed":
                            formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
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

                        case "Nominee Selected":
                            formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            nomItem["Success_x0020_Email"] = formattedBody;
                            nom = (FieldUserValue[])nomItem["Nominees"];
                            To.AddRange(nom);

                            n = (FieldUserValue)nomItem["Nominator"];
                            CC.Add(n);

                            n = (FieldUserValue)nomItem["Submitter"];
                            CC.Add(n);
                            break;

                        case "Reviewer":
                            formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            nomItem["Reviewer_x0020_Email"] = formattedBody;
                            nom = (FieldUserValue[])nomItem["Reviewers"];
                            To.AddRange(nom);

                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);
                            break;

                        case "Reviewer Reminder":
                            formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
                            nomItem["Reminder_x0020_Email"] = formattedBody;
                            nom = (FieldUserValue[])nomItem["Reviewers"];
                            To.AddRange(nom);

                            coords = (FieldUserValue[])nomItem["Coordinators"];
                            if (coords != null)
                                CC.AddRange(coords);
                            break;

                        case "Nominator Notify":
                            formattedBody = ExpandEmailBody(body, nomItem, ImageUrl);
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
                    updateEmailSendList(clientContext, nominationId, nomItem, subject, formattedBody, To, CC, templateType, sendMail);
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError("Failed to update nomination list", ex);
                Console.Error.WriteLine("Failed to update nomination list" + ex);
            }

            
        }

        //updates the emailsend list based on the nominationID
        public static void updateEmailSendList(ClientContext clientContext, string nominationId, ListItem nomItem, string subject, string body, List<FieldUserValue> To, List<FieldUserValue> CC, string templateType, bool sendMail)
        {
            try
            {
                List emailSend = clientContext.Web.Lists.GetByTitle("EmailSend");

                string status = Convert.ToString(nomItem["Submission_x0020_Status"]);
                string team = Convert.ToString(nomItem["Team_x0020_Name"]);
                StringBuilder builder = new StringBuilder(subject);
                builder.Replace("{0}", team);
                subject = builder.ToString();

                CamlQuery query;

                query = new CamlQuery();
                query.ViewXml = string.Format("<View><Query>" +
                                "<Where>" +
                                        "<Eq><FieldRef Name='Nomination' LookupId='TRUE'/>" +
                                        "<Value Type='Lookup'>{0}</Value></Eq>" +
                                "</Where></Query><ViewFields><FieldRef Name='ID' /><RowLimit>500</RowLimit></ViewFields></View>", nominationId);

                Trace.TraceInformation("Query emailsend list based on nominationID query:" + query.ViewXml);
                Console.Out.WriteLine("Query emailsend list based on nominationID query:" + query.ViewXml);
                Microsoft.SharePoint.Client.ListItemCollection mails = emailSend.GetItems(query);
                clientContext.Load(mails);
                clientContext.ExecuteQuery();


                Trace.TraceInformation("Submission Status =" + status);
                Console.Out.WriteLine("Submission Status =" + status);
                Trace.TraceInformation("NominationID =" + nominationId + "Template Type =" + templateType+ "Team Name =" +team);
                Console.Out.WriteLine("NominationID =" + nominationId + "Template Type =" + templateType + "Team Name =" + team);

                foreach (ListItem mail in mails)
                {
                    var update = false;
                    string sendType = Convert.ToString(mail["Send_x0020_Type"]);

                    if (status == "Draft" && templateType == "Manager Retract Notify")
                    {
                        update = true;
                    }
                    if (status == "Submitted" && templateType == "Submitted")
                    {
                        update = true;
                    }
                    if (status == "WaitingManagerApproval" && templateType == "Manager" && sendType == "Manager")
                    {
                        update = true;
                    }
                    if (status == "InReview" && templateType == "Reviewer" && sendType == "Reviewer")
                    {
                        update = true;
                    }
                    if (status == "Completed" && templateType == "Nominee Selected")
                    {
                        update = true;
                    }
                    if (status == "NomineeFailed" && templateType == "Rejected")
                    {
                        update = true;
                    }
                    if (status == "NominatorNotify" && templateType == "Nominator Notify")
                    {
                        update = true;
                    }
                    if (status == "ManagerRejected" && templateType == "Manager Rejected")
                    {
                        update = true;
                    }

                    if (update)
                    {
                        mail["Send_x0020_Mail"] = sendMail;
                        mail["Subject"] = subject;
                        mail["Body"] = body;
                        mail["To"] = To;
                        mail["CC"] = CC;
                        mail.Update();
                        clientContext.ExecuteQuery();
                    }

                }
            }
            catch (Exception ex)
            {
                Trace.TraceError("Failed to update emailsend list", ex);
                Console.Error.WriteLine("Failed to update emailsend list" + ex);
            }

        }

        //updates the nomination data in the body
        public static string ExpandEmailBody(string text, ListItem nomItem, string imageUrl)
        {
            try
            {
                const string NOMINATOR = "${NOMINATOR}";
                const string NOMINATION_SUMMARY = "${SUMMARY}";
                const string NOMINATION_SIGNATURE = "${SIGNATURE}";
                const string NOMINATION_TEAMNAME = "${TEAMNAME}";
                const string NOMINATION_SUBMITDATE = "${SUBMITDATE}";
                const string NOMINATION_URL = "${NOMINATION}";

                string nominatorUserID = Convert.ToString(((FieldUserValue)nomItem["Nominator"]).LookupValue);
                string TeamName = Convert.ToString(nomItem["Team_x0020_Name"]);
                string SubmittedDate = Convert.ToString((DateTime)nomItem["Submitted_x0020_Date"]);

                string url = string.Format("{0}/Nomination Summary/{1}-{1}.pdf", ConfigurationManager.AppSettings["SiteUrl"], nomItem.Id.ToString());
                url = url.Replace(" ", "%20");

                string tag = string.Format("<a href='{0}'>{1}</a>", url, "Click here for nomination summary");

                string img = string.Format("<img src='{0}'/>", imageUrl);

                text = text.Replace(NOMINATOR, nominatorUserID);
                text = text.Replace(Escape(NOMINATOR), nominatorUserID);

                text = text.Replace(NOMINATION_SUMMARY, tag);
                text = text.Replace(Escape(NOMINATION_SUMMARY), tag);

                text = text.Replace(NOMINATION_SIGNATURE, img);
                text = text.Replace(Escape(NOMINATION_SIGNATURE), img);

                text = text.Replace(NOMINATION_TEAMNAME, TeamName);
                text = text.Replace(Escape(NOMINATION_TEAMNAME), TeamName);


                text = text.Replace(NOMINATION_SUBMITDATE, SubmittedDate);
                text = text.Replace(Escape(NOMINATION_SUBMITDATE), SubmittedDate);

                text = text.Replace(NOMINATION_URL, "<nomination.url>");
                text = text.Replace(Escape(NOMINATION_URL), "<nomination.url>");

            }
            catch (Exception ex)
            {
                Trace.TraceError("Failed to format the email body", ex);
                Console.Error.WriteLine("Failed to format the email body" + ex);
            }
            return text;
        }

        private static string Escape(string token)
        {
            try
            {
                token = token.Replace("{", "&#123;");
                token = token.Replace("}", "&#125;");
            }
            catch (Exception ex)
            {
                Trace.TraceError("Failed to replace character", ex);
                Console.Error.WriteLine("Failed to replace character" + ex);
            }
            return token;
        }
    }
}