using ConnectorTestDataHandler.Models;
using ConnectorTestDataHandler.ServiceClients;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel.DataAnnotations;
using ConnectorTestDataHandler.Helpers;

namespace ConnectorTestDataHandler.TestDataHandlers
{
	public class Office365MailTestDataHandler
	{
		//Service Client
		Office365MailServiceClient mailServiceClient;

		/// <summary>
		/// Initializes a new instance of <see cref="Office365MailTestDataHandler"/>.
		/// </summary>
		/// <param name="connString">Connection string as a key value pair.</param>
		public Office365MailTestDataHandler(Dictionary<string, string> connString)
		{
			this.mailServiceClient = new Office365MailServiceClient(connString);
		}

        /// <summary>
        /// Gets all the mails for the logged in user.
        /// </summary>
        /// <returns></returns>
		public List<Office365MailModel> GetAllMails()
        {
            List<Office365MailModel> mailModelList = new List<Office365MailModel>();
            List<Message> mails = this.mailServiceClient.GetAllMails();

            foreach(Message mail in mails)
            {
                //Filtering out calendar invitations
                if (!mail.ODataType.Equals("#microsoft.graph.eventMessageRequest"))
                {
                    Office365MailModel mailModel = new Office365MailModel()
                    {
                        Subject = mail.Subject ?? string.Empty,
                        Body = mail.Body?.Content ?? string.Empty,
                        SenderEmail = mail.Sender?.EmailAddress?.Address ?? string.Empty,
                        //skipping non-email formatted and duplicated emails added by MW
                        ToRecipients = EmailUtilities.GetValidDistinctEmailList(mail.ToRecipients.Select(x => x.EmailAddress.Address)).ToList()
                                                        .Select(x => EmailUtilities.GetMailNickName(x)),
                        ParentFolderName = this.mailServiceClient.GetMailFolderNameById(mail.ParentFolderId)
                    };

                    mailModelList.Add(mailModel);
                }
            }

            return mailModelList;
        }

        /// <summary>
        /// Gets all the events for the logged in user.
        /// </summary>
        /// <returns></returns>
        public List<Office365CalendarEventModel> GetAllCalenderEvents()
		{
			List<Office365CalendarEventModel> eventModelList = new List<Office365CalendarEventModel>();
			List<Event> events = this.mailServiceClient.GetAllCalenderEvents();

			foreach (Event evnt in events)
			{
				Office365CalendarEventModel eventModel = new Office365CalendarEventModel()
				{
					Subject = evnt.Subject ?? string.Empty,
					Body = evnt.Body?.Content ?? string.Empty,
					Organizer = evnt.Organizer?.EmailAddress?.Address ?? string.Empty,
					//skipping non-email formatted and duplicated emails added by MW
					Attendees = EmailUtilities.GetValidDistinctEmailList(evnt.Attendees.Select(x => x.EmailAddress.Address))
												.Select(x => EmailUtilities.GetMailNickName(x)),
				};
				
				eventModelList.Add(eventModel);
			}

			return eventModelList;
		}

		/// <summary>
		/// Gets all the contacts for the logged in user.
		/// </summary>
		/// <returns></returns>
		public List<Office365ContactModel> GetAllContacts()
		{
			List<Office365ContactModel> contactsModelList = new List<Office365ContactModel>();
			List<Contact> contacts = this.mailServiceClient.GetAllContacts();

			foreach (Contact contact in contacts)
			{
				Office365ContactModel contactModel = new Office365ContactModel()
				{
					DisplayName = contact.DisplayName
				};
				contactsModelList.Add(contactModel);
			}

			return contactsModelList;
		}

		/// <summary>
		/// Cleans up the content in the logged account.
		/// </summary>
		public void Cleanup()
		{
			//TO-DO: Delete mail folders
			this.mailServiceClient.DeleteMails();
			this.mailServiceClient.DeleteEvents();
			this.mailServiceClient.DeleteContacts();
		}
	}
}
