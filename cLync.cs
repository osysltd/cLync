// .NET namespaces
using System;
using System.Configuration;
using System.Windows.Forms;

// UCMA namespaces
using Microsoft.Rtc.Collaboration.Presence;
using Microsoft.Rtc.Signaling;

// UCMA samples namespaces
using Microsoft.Rtc.Collaboration.Sample.Common;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Reflection;
using System.Collections.Generic;

namespace Microsoft.Rtc.Collaboration.cLync
{
    public class cLyncApp : Form
    {

        #region Locals
        // This is an example of Note category being specified via XML.
        // A category can either be specified via XML or directly through a 
        // property of the Presence object of an endpoint. The below is an 
        // example of specifying a category via XML. The value of the Note 
        // category is displayed in the Microsoft Lync user interface. This value 
        // will be displayed while the sample is running; and will be reset to 
        // its original value upon exit from this sample code.
        private static String _noteXml = "<note xmlns=\"http://schemas.microsoft.com/2006/09/sip/note\" >"
            + "<body type=\"personal\" uri=\"\" >{0}</body></note>";
        private static string _noteValue = ConfigurationManager.AppSettings["UserNote"];

        // Category Note published using raw xml.
        private CustomPresenceCategory _note;

        // This variable stores the helper class instance.
        private UCMASampleHelper _helper;

        #region UCMA 3.0 Core Classes
        // This variable stores the user endpoint created on behalf of the user 
        // that the sample logs-in as.
        private UserEndpoint _userEndpoint;

        // This class encapsulates the presence of the sample's user endpoint.
        private LocalOwnerPresence _localOwnerPresence;

        // This variable is used to publish the availability of the sample's 
        // logged-in user.
        private PresenceState _userState;

        // This variable is used to publish the state of the samples's logged-in
        // user's phone.
        private PresenceState _phoneState;

        // This variable is used to publish the state of the sample's logged-in 
        // user's computer.
        private PresenceState _machineState;

        // This variable is used to publish the ContactCard of the sample's 
        // logged-in user.
        private ContactCard _contactCard;

        #endregion
        #endregion
        // Presence published flag
        private bool _isPresencePublished = false;

        // TrayIcon and context menu
        private static int _balloonTipTimeout = 5;
        private static NotifyIcon _trayIcon;
        private ContextMenu _trayMenu;
        private MenuItem _menuStatus;
        private List<string> _statuses = new List<string>();

        [STAThread]
        public static void Main()
        {
            if (DateTime.Now < new DateTime(2017, 08, 23))
            { Application.Run(new cLyncApp()); }
            else { throw new UnauthorizedAccessException(); }
        }

        public cLyncApp()
        {
            try
            {
                // Set Console output to the notification area
                using (var consoleWriter = new ConsoleWriter())
                {
                    consoleWriter.WriteEvent += consoleWriter_WriteEvent;
                    consoleWriter.WriteLineEvent += consoleWriter_WriteLineEvent;
                    Console.SetOut(consoleWriter);
                }

                //Set the list of statuses
                _statuses.Add("Available");
                _statuses.Add("Busy");
                _statuses.Add("In a call");
                _statuses.Add("In a meeting");
                _statuses.Add("In a conference call");
                _statuses.Add("Presenting");
                _statuses.Add("Offline");

                _menuStatus = new MenuItem("Status");
                foreach (string status in _statuses)
                {
                    _menuStatus.MenuItems.Add(new MenuItem(status, OnPublishPresence));
                }

                // Create a simple tray menu
                _trayMenu = new ContextMenu();
                _trayMenu.MenuItems.Add("Connect", OnConnect);
                _trayMenu.MenuItems.Add(_menuStatus);
                _trayMenu.MenuItems.Add("Disconnect", OnPublishPresence);
                _trayMenu.MenuItems.Add("Test", OnTest);
                _trayMenu.MenuItems.Add("Reload", OnReload);
                _trayMenu.MenuItems.Add("Restart", OnRestart);
                _trayMenu.MenuItems.Add("Exit", OnExit);

                // Create a tray icon from application icon
                _trayIcon = new NotifyIcon();
                _trayIcon.Text = ((AssemblyTitleAttribute)Attribute.GetCustomAttribute(Assembly.GetExecutingAssembly(), typeof(AssemblyTitleAttribute), false)).Title;
                _trayIcon.Icon = new Icon(Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location), 40, 40);

                // Add menu to tray icon and show it
                _trayIcon.ContextMenu = _trayMenu;
                _trayIcon.Visible = true;
            }
            catch (Exception) { throw; }
        }

        protected override void OnLoad(EventArgs e)
        {
            Visible = false; // Hide form window
            ShowInTaskbar = false; // Remove from taskbar
            base.OnLoad(e);
        }

        private void OnTest(object sender, EventArgs e)
        {

            Console.WriteLine("Test message");
        }

        private void OnConnect(object sender, EventArgs e)
        {
            try
            {
                if (object.ReferenceEquals(null, _helper))
                {
                    // Prepare and instantiate the platform with friendly name of the sample's user endpoint
                    _helper = new UCMASampleHelper();
                    UserEndpointSettings userEndpointSettings = _helper.ReadUserSettings(((AssemblyTitleAttribute)Attribute.GetCustomAttribute(Assembly.GetExecutingAssembly(), typeof(AssemblyTitleAttribute), false)).Title);

                    // Set auto subscription to LocalOwnerPresence.
                    userEndpointSettings.AutomaticPresencePublicationEnabled = true;

                    // Set the capabilities
                    userEndpointSettings.Presence.PreferredServiceCapabilities.ApplicationSharingSupport = CapabilitySupport.Supported;
                    userEndpointSettings.Presence.PreferredServiceCapabilities.AudioSupport = CapabilitySupport.Supported;
                    userEndpointSettings.Presence.PreferredServiceCapabilities.InstantMessagingSupport = CapabilitySupport.Supported;
                    userEndpointSettings.Presence.PreferredServiceCapabilities.VideoSupport = CapabilitySupport.Supported;

                    // Set the status and create endpoint
                    userEndpointSettings.Presence.UserPresenceState = PresenceState.UserAway;
                    _userEndpoint = _helper.CreateUserEndpoint(userEndpointSettings);

                    // LocalOwnerPresence is the main class to manage the 
                    // sample user's presence data.
                    _localOwnerPresence = _userEndpoint.LocalOwnerPresence;

                    // Wire up handlers to receive presence notifications to self.
                    _localOwnerPresence.PresenceNotificationReceived
                        += LocalOwnerPresence_PresenceNotificationReceived;

                    // Establish the endpoint.
                    _helper.EstablishUserEndpoint(_userEndpoint);

                    // Set 'Connected' text for the first menu item
                    _trayMenu.MenuItems[0].Text = "Connected";
                    _trayMenu.MenuItems[0].Checked = true;
                }
            }
            catch (Exception) { throw; }
        }
        private void OnReload(object sender, EventArgs e)
        {
            ConfigurationManager.RefreshSection("appSettings");
            Console.WriteLine("Endpoint contact card configuration has been reloaded.");
        }

        private void OnRestart(object sender, EventArgs e)
        {
            if (!object.ReferenceEquals(null, _helper))
            {
                // Un-wire the presence notification event handler.
                _localOwnerPresence.PresenceNotificationReceived -= LocalOwnerPresence_PresenceNotificationReceived;

                // Shut down platform before exiting the sample.
                _helper.ShutdownPlatform();
            }
            Application.Restart();
        }

        private void OnExit(object sender, EventArgs e)
        {
            if (!object.ReferenceEquals(null, _helper))
            {
                // Un-wire the presence notification event handler.
                _localOwnerPresence.PresenceNotificationReceived -= LocalOwnerPresence_PresenceNotificationReceived;

                // Shut down platform before exiting the sample.
                _helper.ShutdownPlatform();
            }
            this.Dispose(true);
            Application.Exit();
        }

        private void OnPublishPresence(object sender, EventArgs e)
        {
            MenuItem status = (MenuItem)sender;
            if (status.Text == "Disconnect")
            {
                if (_isPresencePublished)
                { PublishPresenceCategories(false); }

                // Checked cleanup
                foreach (MenuItem menuItem in _menuStatus.MenuItems) { menuItem.Checked = false; }
            }
            else
            {
                if (object.ReferenceEquals(null, _helper)) { OnConnect(sender, e); }
                PublishPresenceCategories(true, status.Text);

                // Set appropriate flag for menu item
                foreach (MenuItem menuItem in _menuStatus.MenuItems)
                {
                    if (menuItem.Text == status.Text) { menuItem.Checked = true; }
                    else { menuItem.Checked = false; }
                }
            }
        }

        // Publish note, state, contact card presence categories
        private void PublishPresenceCategories(bool publishFlag, string status = null)
        {
            try
            {
                if (publishFlag == true)
                {
                    // The CustomPresenceCategory class enables creation of a
                    // category using XML. This allows precise crafting of a
                    // category, but it is also possible to create a category
                    // in other, more simple ways, shown below.
                    _note = new CustomPresenceCategory("note", String.Format(_noteXml, _noteValue));

                    switch (status)
                    {
                        default: //Available
                            _userState = PresenceState.UserAvailable;
                            Console.WriteLine(DateTime.Now.ToString() + " Setting available status.");
                            // It is possible to create and publish state with a custom availablity string, shown below
                            LocalizedString localizedAvailableString = new LocalizedString("Available");

                            // Create a PresenceActivity indicating the "In a call" state.
                            PresenceActivity Available = new PresenceActivity(localizedAvailableString);

                            // Set the Availability of the "In a call" state to Busy.
                            Available.SetAvailabilityRange((int)PresenceAvailability.Online,
                                (int)PresenceAvailability.Online);

                            // Microsoft Lync will also show the Busy presence icon.
                            _phoneState = new PresenceState(
                                (int)PresenceAvailability.Online,
                                Available,
                                PhoneCallType.Voip,
                                "phone uri");
                            break;

                        case "Busy": //Busy
                            Console.WriteLine(DateTime.Now.ToString() + " Setting busy status.");
                            _userState = PresenceState.UserBusy;
                            LocalizedString localizedBusyString = new LocalizedString("Busy");
                            PresenceActivity Busy = new PresenceActivity(localizedBusyString);
                            Busy.SetAvailabilityRange((int)PresenceAvailability.Busy,
                                (int)PresenceAvailability.IdleBusy);
                            _phoneState = new PresenceState(
                                (int)PresenceAvailability.Busy,
                                Busy,
                                PhoneCallType.Voip,
                                "phone uri");
                            break;


                        case "In a call": //In a call
                            Console.WriteLine(DateTime.Now.ToString() + " Setting in a call status.");
                            _userState = PresenceState.UserBusy;
                            LocalizedString localizedCallString = new LocalizedString("In a call");
                            PresenceActivity inACall = new PresenceActivity(localizedCallString);
                            inACall.SetAvailabilityRange((int)PresenceAvailability.Busy,
                                (int)PresenceAvailability.IdleBusy);
                            _phoneState = new PresenceState(
                                (int)PresenceAvailability.Busy,
                                inACall,
                                PhoneCallType.Voip,
                                "phone uri");
                            break;

                        case "In a meeting": //In a meeting
                            Console.WriteLine(DateTime.Now.ToString() + " Setting in a meeting status.");
                            _userState = PresenceState.UserBusy;
                            LocalizedString localizedMeetingString = new LocalizedString("In a meeting");
                            PresenceActivity inAMeeting = new PresenceActivity(localizedMeetingString);
                            inAMeeting.SetAvailabilityRange((int)PresenceAvailability.Busy,
                                (int)PresenceAvailability.IdleBusy);
                            _phoneState = new PresenceState(
                                (int)PresenceAvailability.Busy,
                                inAMeeting,
                                PhoneCallType.Voip,
                                "phone uri");
                            break;

                        case "In a conference call": //In a conference call
                            Console.WriteLine(DateTime.Now.ToString() + " Setting in a conference call status.");
                            _userState = PresenceState.UserBusy;
                            LocalizedString localizedConferenceString = new LocalizedString("In a conference call");
                            PresenceActivity inAConference = new PresenceActivity(localizedConferenceString);
                            inAConference.SetAvailabilityRange((int)PresenceAvailability.Busy,
                                (int)PresenceAvailability.IdleBusy);
                            _phoneState = new PresenceState(
                                (int)PresenceAvailability.Busy,
                                inAConference,
                                PhoneCallType.Voip,
                                "phone uri");
                            break;

                        case "Presenting": //Presenting
                            Console.WriteLine(DateTime.Now.ToString() + " Setting presenting status.");
                            _userState = PresenceState.UserBusy;
                            LocalizedString localizedPresentingString = new LocalizedString("Presenting");
                            PresenceActivity Presenting = new PresenceActivity(localizedPresentingString);
                            Presenting.SetAvailabilityRange((int)PresenceAvailability.Busy,
                                (int)PresenceAvailability.DoNotDisturb);
                            _phoneState = new PresenceState(
                                (int)PresenceAvailability.DoNotDisturb,
                                Presenting,
                                PhoneCallType.Voip,
                                "phone uri");
                            break;

                        case "Offline": //Offline
                            Console.WriteLine(DateTime.Now.ToString() + " Setting offline status.");
                            _userState = PresenceState.UserOffline;
                            LocalizedString localizedOfflineString = new LocalizedString("Offline");
                            PresenceActivity Offline = new PresenceActivity(localizedOfflineString);
                            Offline.SetAvailabilityRange((int)PresenceAvailability.Offline,
                                (int)PresenceAvailability.Offline);
                            _phoneState = new PresenceState(
                                (int)PresenceAvailability.Offline,
                                Offline,
                                PhoneCallType.Voip,
                                "phone uri");
                            break;
                    }

                    // Machine or Endpoint states must always be published to
                    // indicate the endpoint is actually online, otherwise it is
                    // assumed the endpoint is offline, and no presence
                    // published from that endpoint will be displayed.
                    _machineState = PresenceState.EndpointOnline;

                    // It is also possible to create presence categories such
                    // as ContactCard, Note, PresenceState, and Services with
                    // their constructors.
                    // Here we create a ContactCard and change the title.
                    _contactCard = new ContactCard(0);

                    /* The title string to be displayed. */
                    if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["UserTitle"]))
                    {
                        LocalizedString localizedTitleString = new LocalizedString(ConfigurationManager.AppSettings["UserTitle"]);
                        _contactCard.JobTitle = localizedTitleString.Value;
                    }

                    if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["UserOffice"]))
                    {
                        LocalizedString localizedTitleString = new LocalizedString(ConfigurationManager.AppSettings["UserOffice"]);
                        _contactCard.OfficeLocation = localizedTitleString.Value;
                    }

                    if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["JobTitle"]))
                    {
                        LocalizedString localizedTitleString = new LocalizedString(ConfigurationManager.AppSettings["JobTitle"]);
                        _contactCard.JobTitle = localizedTitleString.Value;
                    }
                    // Publish a photo
                    // If the supplied value for photo is null or empty, then set value of IsAllowedToShowPhoto to false
                    _contactCard.IsAllowedToShowPhoto = false;

                    string photoUri = ConfigurationManager.AppSettings["PhotoUri"];
                    if (!String.IsNullOrEmpty(photoUri))
                    {
                        _contactCard.IsAllowedToShowPhoto = true;
                        _contactCard.PhotoUri = photoUri;
                    }

                    // Publish all presence categories with new values.
                    _localOwnerPresence.BeginPublishPresence(
                        new PresenceCategory[]
                        {
                            _userState,
                            _phoneState,
                            _machineState,
                            _note,
                            _contactCard
                        },
                        PublishPresenceCompleted, /* async callback when publishing operation completes. */
                        publishFlag /* value TRUE indicates that presence to be published with new values. */);
                }
                else
                {
                    // Delete all presence categories.
                    Console.WriteLine(DateTime.Now.ToString() + " Revert status back.");
                    _localOwnerPresence.BeginDeletePresence(
                        new PresenceCategory[]
                        {
                            _userState,
                            _phoneState,
                            _machineState,
                            _note,
                            _contactCard
                        },
                        PublishPresenceCompleted,
                        publishFlag /* value FALSE indicates that presence reverted to original values. */);
                }
            }
            catch (PublishSubscribeException pse)
            {
                // PublishSubscribeException is thrown when there were
                // exceptions during this presence operation such as badly
                // formed sip request, duplicate publications in the same
                // request etc.
                // TODO (Left to the reader): Include exception handling code
                // here.
                Console.WriteLine(pse.ToString());
            }
            catch (RealTimeException rte)
            {
                // RealTimeException is thrown when SIP Transport, SIP
                // Authentication, and credential-related errors are 
                // encountered.
                // TODO (Left to the reader): Include exception handling code
                // here.
                Console.WriteLine(rte.ToString());
            }
        }

        // AsyncCallback to publish presence state.
        private void PublishPresenceCompleted(IAsyncResult result)
        {
            try
            {
                // Since the same call back function is used to publish
                // presence categories and to delete presence categories,
                // retrieve the flag indicating which operation is desired.
                bool isPublishOperation;
                bool.TryParse(result.AsyncState.ToString(), out isPublishOperation);

                if (isPublishOperation)
                {
                    // Complete the publishing of presence categories.
                    _localOwnerPresence.EndPublishPresence(result);
                    _isPresencePublished = true;
                    Console.WriteLine("Presence state has been published.");
                }
                else
                {
                    // Complete the deleting of presence categories.
                    _localOwnerPresence.EndDeletePresence(result);
                    _isPresencePublished = false;
                    Console.WriteLine("Presence state has been deleted.");
                }
            }
            catch (PublishSubscribeException pse)
            {
                // PublishSubscribeException is thrown when there were
                // exceptions during the publication of this category such as
                // badly formed sip request, duplicate publications in the same
                // request etc
                // TODO (Left to the reader): Include exception handling code
                // here
                Console.WriteLine(pse.ToString());
            }
            catch (RealTimeException rte)
            {
                // RealTimeException is thrown when SIP Transport, SIP
                // Authentication, and credential-related errors are
                // encountered.
                // TODO (Left to the reader): Include exception handling code
                // here.
                Console.WriteLine(rte.ToString());
            }
        }

        /// <summary>
        /// The event handler for the Category Notification. Notifications come
        /// in the form of a list of items. We are only interested in state
        /// publications here, so we will only process those.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LocalOwnerPresence_PresenceNotificationReceived(object sender,
            LocalPresentityNotificationEventArgs e)
        {
            Console.WriteLine("Presence notifications received for target {0}.",
                                this._userEndpoint.OwnerUri);
            // Notifications contain all the notifications for one user.
            // Each user will send a list of updated categories. We will choose
            // the ones we are interested in and process them.

            // Display to console the value of the Aggregate Presence category.
            if (e.AggregatedPresenceState != null)
            {
                Console.WriteLine("Aggregate State = " + e.AggregatedPresenceState.Availability);
            }

            // Display to console the value of the capabilities of the 
            // sample's user endpoint.
            if (e.ServiceCapabilities != null)
            {
                Console.WriteLine("IM: {0}", e.ServiceCapabilities.InstantMessagingEnabled);
                Console.WriteLine("Audio: {0}", e.ServiceCapabilities.AudioEnabled);
                Console.WriteLine("Video: {0}", e.ServiceCapabilities.VideoEnabled);
                Console.WriteLine("AppSharing: {0}", e.ServiceCapabilities.ApplicationSharingEnabled);
            }

            // Display to console the value of the ContactCard category.
            if (e.ContactCard != null)
            {
                Console.WriteLine("Title = {0}", e.ContactCard.JobTitle);
            }

            // Display to console the value of the Note category.
            if (e.PersonalNote != null && !String.IsNullOrEmpty(e.PersonalNote.Message))
            {
                Console.WriteLine("{0}: {1}", NoteType.Personal, e.PersonalNote.Message);
            }

        }

        static void consoleWriter_WriteLineEvent(object sender, cLync.ConsoleWriterEventArgs e)
        {
            //MessageBox.Show(e.Value, "WriteLine");
            if (!String.IsNullOrEmpty(e.Value))
            {
                _trayIcon.Visible = true;
                _trayIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
                _trayIcon.BalloonTipText = e.Value;
                _trayIcon.ShowBalloonTip(_balloonTipTimeout);
            }
        }

        static void consoleWriter_WriteEvent(object sender, cLync.ConsoleWriterEventArgs e)
        {
            //MessageBox.Show(e.Value, "Write");
            if (!String.IsNullOrEmpty(e.Value))
            {
                _trayIcon.Visible = true;
                _trayIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
                _trayIcon.BalloonTipText = e.Value;
                _trayIcon.ShowBalloonTip(_balloonTipTimeout);
            }
        }

        protected override void Dispose(bool isDisposing)
        {
            if (isDisposing)
            {
                // Release the icon resource.
                _trayIcon.Dispose();
            }
            base.Dispose(isDisposing);
        }
    }

    public class ConsoleWriterEventArgs : EventArgs
    {
        public string Value { get; private set; }
        public ConsoleWriterEventArgs(string value) { Value = value; }
    }

    public class ConsoleWriter : TextWriter
    {
        public override Encoding Encoding { get { return Encoding.UTF8; } }

        public override void Write(string value)
        {
            if (WriteEvent != null) WriteEvent(this, new ConsoleWriterEventArgs(value));
            base.Write(value);
        }

        public override void WriteLine(string value)
        {
            if (WriteLineEvent != null) WriteLineEvent(this, new ConsoleWriterEventArgs(value));
            base.WriteLine(value);
        }

        public event EventHandler<ConsoleWriterEventArgs> WriteEvent;
        public event EventHandler<ConsoleWriterEventArgs> WriteLineEvent;
    }
}
