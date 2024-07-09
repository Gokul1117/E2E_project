using BitTitan.Migration;
using BitTitan.Migration.Configurations;
using BitTitan.Migration.Connectors.Projects;
using ConnectorTestDataHandler.Models;
using ConnectorTestDataHandler.TestDataHandlers;
using EndToEndTestMigrationHelper;
using Newtonsoft.Json;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MigrationWizTestAutomation.CommonUtilities;
using System.Reflection;
using EndToEndTestMigrationHelper.Responses;
using NUnit.Framework.Legacy;

namespace MigrationTests
{
	[TestFixture]
	[Category("EndToEndMigrationTests")]
	[Category("EndToEndOffice365ToOffice365MigrationTests")]
	public class Office365ToOffice365MigrationTests
	{
		private Dictionary<string, string> sourceConnString;
		private Dictionary<string, string> destinationConnString;
		private Office365MailTestDataHandler sourceTestDataHandler;
		private Office365MailTestDataHandler destTestDataHandler;
		private ExchangeConfiguration sourceEndPointConfig;
		private ExchangeConfiguration destinationEndpointConfig;
        private string[] advancedOptions;
        private IMigrationHelper migrationHelper;
        MailboxItemTypes[] defaultItemTypes = new MailboxItemTypes[] { MailboxItemTypes.Mail, MailboxItemTypes.Calendar, MailboxItemTypes.Contact };
		

		[OneTimeSetUp]
		public void Setup()
		{
			//These parameters might get changed if the authentication mode changes
			this.sourceConnString = new Dictionary<string, string>(){
				{ "ClientId" , Consts.Office365SourceClientId },
				{ "TenantId" , Consts.Office365SourceTenantId },
				{ "ClientSecret", Consts.Office365SourceClientSecret },
				{ "AdminUserName", Consts.Office365SourceMigratingUserName},
				{ "AdminPassword", Consts.Office365SourceMigratingUserPassword} };

			//These parameters might get changed if the authentication mode changes
			this.destinationConnString = new Dictionary<string, string>(){
				{ "ClientId" , Consts.Office365DestinationClientId },
				{ "TenantId" , Consts.Office365DestinationTenantId },
				{ "ClientSecret", Consts.Office365DestinationClientSecret },
				{ "AdminUserName", Consts.Office365DestinationMigratingUserName },
				{ "AdminPassword", Consts.Office365DestinationMigratingUserPassword }};

			//Source and Destination Test Data Handlers
			this.sourceTestDataHandler = new Office365MailTestDataHandler(this.sourceConnString);
			this.destTestDataHandler = new Office365MailTestDataHandler(this.destinationConnString);

			//Office 365 Source Endpoint configuration
			this.sourceEndPointConfig = new ExchangeConfiguration
			{
				UseAdministrativeCredentials = true,
				AdministrativeUsername = Consts.Office365SourceAdminUserName,
				AdministrativePassword = Consts.Office365SourceAdminUserPassword,
			};

			//Office 365 Destination Endpoint configuration
			this.destinationEndpointConfig = new ExchangeConfiguration
			{
				UseAdministrativeCredentials = true,
				AdministrativeUsername = Consts.Office365DestinationAdminUserName,
				AdministrativePassword = Consts.Office365DestinationAdminUserPassword,
			};

            //Instantiate Migration Helper 

           // advancedOptions = new String[] {"ModernAuthClientIdExport=" + Consts.Office365SourceClientId, "ModernAuthTenantIdExport=" + Consts.Office365SourceTenantId,
           // "ModernAuthClientIdImport=" + Consts.Office365DestinationClientId, "ModernAuthTenantIdImport=" + Consts.Office365DestinationTenantId};

            this.migrationHelper = new BaseMigrationHelper(Consts.BittianAPIBaseURL, Consts.MigwizAPIBaseURL);
		}

		[OneTimeTearDown]
		public void OneTimeTearDown()
		{
			//Cleaning up MW entities
			migrationHelper.Cleanup();

			//Cleaning up Destination Mailbox
			destTestDataHandler.Cleanup();
		}

		[Test]
		[TestCaseSource("Office365ToOffice365FullMigrationTestCases")]
		public void Office365ToOffice365FullMigrationTest(Office365MailboxModel expectedMailBox)
		{	
			//Arrange 
			//Retrieve source data
			List<Office365MailModel> sourceMails = this.sourceTestDataHandler.GetAllMails();
			List<Office365CalendarEventModel> sourceEvents = this.sourceTestDataHandler.GetAllCalenderEvents();
			List<Office365ContactModel> sourceContacts = this.sourceTestDataHandler.GetAllContacts();

			//Deleting destination mailbox data before the migration
			destTestDataHandler.Cleanup();

            //Compare source account data with the expected data
            //Compare source and expected mails			
            ClassicAssert.IsTrue(expectedMailBox.Mails.All(sourceMails.Contains), $"Source mailbox {Consts.Office365SourceMigratingUserName} does not have expected data setup.");

            //Compare source and expected events			
            ClassicAssert.IsTrue(expectedMailBox.Events.All(sourceEvents.Contains), $"Source calendar events for the user { Consts.Office365SourceMigratingUserName} does not have expected data setup.");

            //Compare source and expected contacts			
            ClassicAssert.IsTrue(expectedMailBox.Contacts.All(sourceContacts.Contains), $"Source contacts for the user {Consts.Office365SourceMigratingUserName} does not have expected data setup.");

			//Act - Perform Migration
			Guid projectItemId = this.migrationHelper.PerformEndToEndMigration(Consts.MigrationWizUser,
				Consts.MigrationWizUserPassword,
				this.sourceEndPointConfig,
				this.destinationEndpointConfig,
				MailboxConnectorTypes.ExchangeOnline2,
				MailboxConnectorTypes.ExchangeOnline2,
				Consts.Office365SourceMigratingUserName,
				Consts.Office365DestinationMigratingUserName,
				MailboxQueueTypes.Full,
				defaultItemTypes,
				ProjectTypes.Mailbox);

            //Assert
            //Validate the migration status
            ClassicAssert.AreEqual(MailboxQueueStatus.Completed, this.migrationHelper.GetMigrationStatus(projectItemId));

            //Validate migrations errors
            ClassicAssert.AreEqual(0, this.migrationHelper.GetProjectItemErros(projectItemId).FindAll(x => x.Severity.Equals(ErrorSeverityLevel.Error)).Count);

			//Get destination mails after the migration
			List<Office365MailModel> destinationMails = this.destTestDataHandler.GetAllMails();
            ClassicAssert.NotNull(destinationMails, $"Failed to obtain destination mailbox {Consts.Office365DestinationMigratingUserName} after the migration.");

			//Get destination events after the migration
			List<Office365CalendarEventModel> destinationEvents = this.destTestDataHandler.GetAllCalenderEvents();
            ClassicAssert.NotNull(destinationEvents, $"Failed to obtain destination calendar for {Consts.Office365DestinationMigratingUserName} after the migration.");

			//Get destination contacts after the migration
			List<Office365ContactModel> destinationContacts = this.destTestDataHandler.GetAllContacts();
            ClassicAssert.NotNull(destinationContacts, $"Failed to obtain destination contact list for {Consts.Office365DestinationMigratingUserName} after the migration.");

			//Comparing source and destination items and list down any unmatched source items
			List<Office365MailModel> unmatchedSourceMails = sourceMails.Except(destinationMails).ToList();
			List<Office365CalendarEventModel> unmatchedSourceEvents = sourceEvents.Except(destinationEvents).ToList();
			List<Office365ContactModel> unmatchedSourceContacts = sourceContacts.Except(destinationContacts).ToList();
			
			Assert.Multiple(() =>
			{
                ClassicAssert.AreEqual(sourceMails.Count, destinationMails.Count, $"The no of emails at the source account {Consts.Office365SourceMigratingUserName} " +
					$"is not equal to the no of emails at the destination account {Consts.Office365DestinationMigratingUserName}");

                ClassicAssert.IsEmpty(unmatchedSourceMails, $"Unmigrated emails found. Mailbox Id: {projectItemId} " +
					$"Source: {Consts.Office365SourceMigratingUserName} Destination: {Consts.Office365DestinationMigratingUserName}. " +
					$"List of unmatched mails {string.Join(", ", unmatchedSourceMails)}.");

                ClassicAssert.AreEqual(sourceEvents.Count, destinationEvents.Count, $"The no of events at the source account {Consts.Office365SourceMigratingUserName} " +
					$"is not equal to the no of events at the destination account {Consts.Office365DestinationMigratingUserName}");

                ClassicAssert.IsEmpty(unmatchedSourceEvents, $"Unmigrated events found. Mailbox Id: {projectItemId} " +
					$"Source: {Consts.Office365SourceMigratingUserName} Destination: {Consts.Office365DestinationMigratingUserName}. " +
					$"List of unmatched events {string.Join(", ", unmatchedSourceEvents)}.");

                ClassicAssert.AreEqual(sourceContacts.Count, destinationContacts.Count, $"The no of contacts at the source account {Consts.Office365SourceMigratingUserName} " +
					$"is not equal to the no of contacts at the destination account {Consts.Office365DestinationMigratingUserName}");

                ClassicAssert.IsEmpty(unmatchedSourceContacts, $"Unmigrated contatcs found. Mailbox Id: {projectItemId} " +
					$"Source: {Consts.Office365SourceMigratingUserName} Destination: {Consts.Office365DestinationMigratingUserName}. " +
					$"List of unmatched contacts {string.Join(", ", unmatchedSourceContacts)}.");

			});
		}

		[Test]
		[TestCaseSource("Office365VerifyCredentialsFailTestCases")]
		public void Office365MailVerifyCredentialsFailTest(ExchangeConfiguration sourceConfig,
			ExchangeConfiguration destinationConfig, 
			string expectedErrorMessage)
		{
			//Act - Perform Migration
			Guid projectItemId = this.migrationHelper.PerformEndToEndMigration(Consts.MigrationWizUser,
				Consts.MigrationWizUserPassword,
				sourceConfig, 
				destinationConfig,
				MailboxConnectorTypes.ExchangeOnline2,
				MailboxConnectorTypes.ExchangeOnline2,
				Consts.Office365SourceMigratingUserName,
				Consts.Office365DestinationMigratingUserName,
				MailboxQueueTypes.Verification,
				null,
				ProjectTypes.Mailbox);

            //Assert
            //Validate migration status
            ClassicAssert.AreEqual(MailboxQueueStatus.Failed, this.migrationHelper.GetMigrationStatus(projectItemId));

			//Validate migrations errors
			List<ProjectItemError> migrationErros = this.migrationHelper.GetProjectItemErros(projectItemId);
            ClassicAssert.AreEqual(1, migrationErros.Count);
            ClassicAssert.AreEqual(expectedErrorMessage, migrationErros[0].Message);
		}

		/// <summary>
		/// Test cases for <see cref="Office365ToOffice365MigrationTests.Office365MailVerifyCredentialsFailTest"/>
		/// </summary>
		public static IEnumerable<TestCaseData> Office365VerifyCredentialsFailTestCases
		{
			get
			{
				//Test Case: Invalid Source User Name
				yield return new TestCaseData
				(
					new ExchangeConfiguration()
					{
						UseAdministrativeCredentials = true,
						AdministrativePassword = Consts.Office365SourceAdminUserPassword,
						AdministrativeUsername = Consts.Office365SourceAdminUserName.Replace(Consts.Office365SourceAdminUserName.Split('@')[0], Guid.NewGuid().ToString())
					},
					new ExchangeConfiguration
					{
						UseAdministrativeCredentials = true,
						AdministrativePassword = Consts.Office365DestinationAdminUserPassword,
						AdministrativeUsername = Consts.Office365DestinationAdminUserName
					},
                    "Your migration failed while checking source credentials. Http POST request to 'autodiscover-s.outlook.com' failed - 401 "
                );

				//Test Case: Invalid Destination User Name
				yield return new TestCaseData
				(
					new ExchangeConfiguration()
					{
						UseAdministrativeCredentials = true,
						AdministrativePassword = Consts.Office365SourceAdminUserPassword,
						AdministrativeUsername = Consts.Office365SourceAdminUserName
					},
					new ExchangeConfiguration
					{
						UseAdministrativeCredentials = true,
						AdministrativePassword = Consts.Office365DestinationAdminUserPassword,
						AdministrativeUsername = Consts.Office365DestinationAdminUserName.Replace(Consts.Office365DestinationAdminUserName.Split('@')[0], Guid.NewGuid().ToString()),
					},
					"Your migration failed while checking source credentials. Http POST request to 'autodiscover-s.outlook.com' failed - 401 "
				);
			}
		}

		/// <summary>
		/// Test cases for <see cref="Office365ToOffice365MigrationTests.Office365ToOffice365FullMigrationTest"/>
		/// </summary>
		public static IEnumerable<TestCaseData> Office365ToOffice365FullMigrationTestCases
		{
			get
			{
				yield return new TestCaseData
				(
					JsonCommon.DeserializeJSonToObject<Office365MailboxModel>
					(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), Consts.Office365ReleaseRegressionSourceTestDataFile))
				);
			}
		}
	}

	


}
