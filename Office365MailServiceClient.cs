using ConnectorTestDataHandler.Authenticators;
using Microsoft.Graph;
using System;
using System.Collections.Generic;

namespace ConnectorTestDataHandler.ServiceClients
{
	/// <summary>
	/// Represents an instance of <see cref="Office365MailServiceClient"/>.
	/// </summary>
	public class Office365MailServiceClient
	{
		private Dictionary<string, string> connString;
		private GraphServiceClient graphClientV1;
		private MSGraphAuthenticator graphAuthenticatorV1;

		/// <summary>
		/// Gets the Graph Client for V1.
		/// </summary>
		public GraphServiceClient GraphClientV1 { get => this.graphClientV1; }

		/// <summary>
		/// Initializes a new instance of <see cref="Office365MailServiceClient"/>.
		/// </summary>
		/// <param name="connectionParams">Connection string as a key value pair.</param>
		public Office365MailServiceClient(Dictionary<string, string> connectionParams)
		{
			this.graphAuthenticatorV1 = new MSGraphAuthenticator(connectionParams, MSGraphAuthenticator.GraphAPIVersion.v1);
			this.graphClientV1 = this.graphAuthenticatorV1.GetAuthenticatedClient();
		}

		/// <summary>
		/// Gets the mail folder name by id.
		/// </summary>
		/// <param name="parentFolderId">Parent Folder Id.</param>
		/// <returns></returns>
		public string GetMailFolderNameById(string parentFolderId)
		 => graphClientV1.Me.MailFolders[parentFolderId].Request().GetAsync().Result?.DisplayName ?? string.Empty;

        /// <summary>
        /// Gets all the mails for the logged in user.
        /// </summary>
        /// <returns></returns>
        //public List<Message> GetAllMails()
        //{
        //    List<Message> mails = new List<Message>();
        //    //var result = graphClientV1.Me.Messages.Request().GetAsync();
        //    //IUserMessagesCollectionPage currentPage = result.Result;
        //    IUserMessagesCollectionPage currentPage = graphClientV1.Me.Messages.Request().GetAsync().Result;
        //    while (currentPage != null)
        //    {
        //        foreach (var mail in currentPage)
        //        {
        //            try
        //            {
        //                mails.Add(mail);
        //            }
        //            catch (Exception) { }
        //        }
        //        currentPage = (currentPage.NextPageRequest != null) ? currentPage.NextPageRequest.GetAsync().Result : null;
        //    }
        //    return mails;
        //}


        public List<Message> GetAllMails()
        {
            List<Message> mails = new List<Message>();

            try
            {
                IUserMessagesCollectionPage currentPage = graphClientV1.Me.Messages.Request().GetAsync().Result;

                while (currentPage != null)
                {
                    foreach (var mail in currentPage)
                    {
                        try
                        {
                            mails.Add(mail);
                        }
                        catch (Exception ex)
                        {
                            // Log the exception (ex)
                            Console.WriteLine($"Error adding mail: {ex.Message}");
                        }
                    }

                    currentPage = (currentPage.NextPageRequest != null) ? currentPage.NextPageRequest.GetAsync().Result : null;
                }
            }
            catch (NullReferenceException ex)
            {
                // Log the exception
                Console.WriteLine($"Null reference exception: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Log the exception
                Console.WriteLine($"General exception: {ex.Message}");
            }

            return mails;
        }


        //public List<Message> GetAllMails()
        //{
        //    List<Message> mails = new List<Message>();

        //    var messages = graphClientV1.Me.Messages
        //        .Request()
        //        .Select(e => new
        //        {
        //            e.Sender,
        //            e.Subject,
        //            e.BodyPreview,
        //            e.ReceivedDateTime
        //        })
        //        .Top(1000)
        //        .GetAsync().GetAwaiter().GetResult();

        //    var pageIterator = PageIterator<Message>
        //        .CreatePageIterator(
        //            graphClientV1,
        //            messages,
        //            // Callback executed for each item in the collection
        //            (m) =>
        //            {
        //                try
        //                {
        //                    mails.Add(m);
        //                }
        //                catch (Exception)
        //                {
        //                    // Handle exceptions if necessary 
        //                }
        //                return true;
        //            }
        //        );

        //    pageIterator.IterateAsync().GetAwaiter().GetResult();

        //    return mails;
        //}


        /// <summary>
        /// Gets all the events for the logged in user.
        /// </summary>
        /// <returns></returns>
        public List<Event> GetAllCalenderEvents()
		{
			List<Event> events = new List<Event>();
			IUserEventsCollectionPage currentPage = graphClientV1.Me.Events.Request().GetAsync().Result;
			while (currentPage != null)
			{
				foreach (var evnt in currentPage)
				{
					try
					{
						events.Add(evnt);
					}
					catch (Exception) { }
				}
				currentPage = (currentPage.NextPageRequest != null) ? currentPage.NextPageRequest.GetAsync().Result : null;
			}
			return events;
		}


		/// <summary>
		/// Gets all the contacts for the logged in user.
		/// </summary>
		/// <returns></returns>
		public List<Contact> GetAllContacts()
		{
			List<Contact> contacts = new List<Contact>();
			IUserContactsCollectionPage currentPage = graphClientV1.Me.Contacts.Request().GetAsync().Result;
			while (currentPage != null)
			{
				foreach (var contact in currentPage)
				{
					try
					{
						contacts.Add(contact);
					}
					catch (Exception) { }
				}
				currentPage = (currentPage.NextPageRequest != null) ? currentPage.NextPageRequest.GetAsync().Result : null;
			}
			return contacts;
		}

		/// <summary>
		/// Deletes all the emails of the logged in account. 
		/// </summary>
		public async void DeleteMails()
		{
			List<Message> mails = GetAllMails();

			foreach (Message mail in mails)
			{
				await graphClientV1.Me.Messages[mail.Id].Request().DeleteAsync();
			}
		}

		/// <summary>
		/// Deletes all the events of the logged in account. 
		/// </summary>
		public async void DeleteEvents()
		{
			List<Event> events = GetAllCalenderEvents();

			foreach (Event evnt in events)
			{
				await graphClientV1.Me.Events[evnt.Id].Request().DeleteAsync();
			}
		}

		/// <summary>
		/// Deletes all the contacts of the logged in account. 
		/// </summary>
		public async void DeleteContacts()
		{
			List<Contact> contacts = GetAllContacts();

			foreach (Contact contact in contacts)
			{
				await graphClientV1.Me.Contacts[contact.Id].Request().DeleteAsync();
			}
		}
	}
}