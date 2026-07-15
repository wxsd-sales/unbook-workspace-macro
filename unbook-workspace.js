/********************************************************
 *
 * Macro Author:      	William Mills
 *                    	Solutions Engineer
 *                    	wimills@cisco.com
 *                    	Cisco Systems
 *
 * Version: 3-0-1
 * Released: 07/15/26
 *
 * This example macro releases empty workspace bookings based
 * off configurable policies. Additionally this macro can log
 * the actions it unbooking actions to a remote logging server
 * in order to monitor and review.
 *
 * MTR devices should work but they must be registered to Webex
 * and Hybrid Calendar enabled.
 *
 * New Configuration Wizard Available:
 * https://wxsd-sales.github.io/unbook-workspace-macro/webapp/
 *
 * Full Readme, source code and license details for this macro are available
 * on GitHub: https://github.com/wxsd-sales/unbook-workspace-macro
 *
 ********************************************************/

import xapi from "xapi";

/*********************************************************
 * Configuration Start
 **********************************************************/

const config = {
  profiles: [
    // An array of profiles for the macro to match and apply difference behaviours
    {
      type: "duration", // Profile type: duration | keywords | organizer
      name: "Short Meetings", // Name of profile for logging
      duration: [0, 60], // Duration of booking in minutes: From zero minutes to 60 minute meetings
      monitor: true, // Enable monitoring for this matched profile
      startMonitoringDelay: 0, // Number of minutes after the booking starts in which to begin monitoring
      stopMonitoringAfter: 10, // Number of minutes after the booking starts in which to stop monitoring
      requiredUnoccupiedDuration: 2, // Number of minutes the workspace is unoccupied before unbooking
      alertBeforeUnbookingDuration: 1, // Number of minutes before unbooking in which to alert user
    },
    {
      type: "duration",
      name: "Between 1 and 2 hour Meetings",
      duration: [61, 180], // Duration of booking in minutes: From 61 minutes to 180 minute meetings
      monitor: true,
      startMonitoringDelay: 0,
      stopMonitoringAfter: 20,
      requiredUnoccupiedDuration: 5,
      alertBeforeUnbookingDuration: 1,
    },
    {
      type: "duration",
      name: "All day meetings - Don't monitor during lunch hours",
      duration: [180, 480], // This profile is a work in progress
      monitor: true,
      startMonitoringDelay: 0,
      stopMonitoringAfter: 10,
      requiredUnoccupiedDuration: 5,
      alertBeforeUnbookingDuration: 1,
    },
    {
      type: "keywords", // Profile type: Keywords
      name: "Meeting Title Keyword", // Name of profile for logging
      keywords: ["Training", "Test"], // Array of keywords in which to look for in the booking title
      monitor: true, // Enable monitoring for this matched profile
      startMonitoringDelay: 0, // Number of minutes after the booking starts in which to begin monitoring
      stopMonitoringAfter: 5, // Number of minutes after the booking starts in which to stop monitoring
      requiredUnoccupiedDuration: 1, // Number of minutes the workspace is unoccupied before unbooking
      alertBeforeUnbookingDuration: 1, // Number of minutes before unbooking in which to alert user
    },
    {
      type: "organizers", // Profile type: Keywords
      name: "Organizers Name", // Name of profile for logging
      organizers: ["John Smith"], // Array of organizer names to match with bookings
      monitor: false, // Disable monitoring for these matched booking
    },
    {
      type: "default",
      name: "Default Booking Handling Profile",
      monitor: true,
      startMonitoringDelay: 0,
      stopMonitoringAfter: 10,
      requiredUnoccupiedDuration: 5,
    },
  ],
  presenceDetection: {
    activeCalls: true, // Consider active calls as presence detected
    presentation: true, // Consider presentations as presence detected
    peopleCount: true, // Consider peopele count as presence detected
    peoplePresence: true, // Consider peopele presence as presence detected
    guiInteractions: true, // Consider GUI inputs as presence detected
  },
  externalLogging: {
    enabled: false, // Enable or Disable External Logging of macro events: true | false
    type: "bot", // Type of logging service: 'webhook' | 'bot'
    contact: "user@example.com", // If logging service is 'bot', give the contact or roomId to notify
    url: "https://example.com/webhook", // If logging service is 'webhook', give the URL of your webhook service or 'https://webexapis.com/v1/messages' for Webex
    token: "<Access Token>", // Either the webhook Access Token or the Webex Bot token
  },
  debugging: false,
};

/*********************************************************
 * Configuration End
 **********************************************************/

var workspaceBookingIds = [];
let bookingStartListener;

xapi.Config.Bookings.CheckIn.Enabled.on(initialize);
xapi.Config.Bookings.CheckIn.Enabled.get().then(initialize);

async function initialize(checkInEnabled) {
  if (checkInEnabled == "True") {
    if (bookingStartListener) {
      bookingStartListener();
      bookingStartListener = () => void 0;
    }
    console.warn(
      "Native RoomOS CheckIn is enabled - Unbooking Workspace Macro Suspended",
    );

    return;
  }

  if (config.externalLogging.enabled) {
    const httpMode = await xapi.Config.HttpClient.Mode.get();
    if (httpMode == "Off") {
      console.warn("HTTP Client Mode is Off - Enabling");
      await xapi.Config.HttpClient.Mode.set("On");
    }
  }

  bookingStartListener = xapi.Event.Bookings.Start.on((event) =>
    processBookingStart(event),
  );

  const bookings = await xapi.Command.Bookings.List();

  if (bookings.Booking) {
    for (let booking of bookings.Booking) {
      console.log("MeetingId:", booking.MeetingId);
      const now = new Date();
      const startTime = new Date(booking.Time.StartTime);
      const endTime = new Date(booking.Time.EndTime);
      console.debug("-Start-", startTime);
      console.debug("--End--", endTime);
      console.debug(
        "Duration:",
        Math.abs(endTime.getTime() - startTime.getTime()) / (60 * 1000),
      );
      if (now >= startTime && now < endTime) {
        console.debug("Current Booking!");
        await startMonitor(booking);
      }
    }
  } else {
    console.debug("No bookings found.");
  }
}

async function processBookingStart({ Id, MeetingId }) {
  console.log("Booking Start Event - Id:", Id, "- MeetingId:", MeetingId);

  const booking = await xapi.Command.Bookings.Get({ Id })
    .then((result) => result.Booking)
    .catch(() => console.log("Cound not get find meeting: ", Id));
  if (!booking) return;
  await startMonitor(booking);
}

async function startMonitor(booking) {
  if (workspaceBookingIds.indexOf(booking.Id) >= 0) {
    console.debug(`Booking ID ${booking.Id} already monitored.`);
    return;
  }

  console.log("Booking Details: ", JSON.stringify(booking));
  const profile = mapToProfile(booking);
  if (!profile) return;

  const mtr = await detectMTR();
  const monitor = new workspaceMonitor(
    booking,
    profile,
    mtr,
    config.externalLogging,
    config.debugging,
  );
  if (profile.monitor) monitor.initialize();
}

function mapToProfile(booking) {
  const profiles = config.profiles;
  const startTime = new Date(booking.Time.StartTime);
  const endTime = new Date(booking.Time.EndTime);
  const duration =
    Math.abs(endTime.getTime() - startTime.getTime()) / (60 * 1000);
  const title = booking.Title;
  const organizer =
    `${booking.Organizer.FirstName ?? ""} ${booking.Organizer.LastName ?? ""}`.trim();
  const profile = profiles.find((profile) =>
    compareProfile(profile, duration, title, organizer),
  );
  console.log("Matched Profile: ", profile);
  return profile;
}

function compareProfile(profile, duration, title, organizer) {
  console.debug(
    `Comparing Profile [${profile.name}] using type [${profile.type}] with - Duration [${duration}] - Title [${title}] - Organizer [${organizer}]`,
  );
  switch (profile.type) {
    case "duration":
      if (duration > profile.duration[0] && duration <= profile.duration[1]) {
        console.debug(
          `Duration [${duration}] is within range [ ${profile.duration[0]} - ${profile.duration[1]} ]`,
        );
        return true;
      } else {
        console.debug(
          `Duration [${duration}] is not with range [ ${profile.duration[0]} - ${profile.duration[1]} ]`,
        );
        return false;
      }
    case "keywords":
      for (let i = 0; i < profile.keywords.length; i++) {
        if (title.includes(profile.keywords[i])) {
          console.debug(
            `Title [${title}] includes keyword [ ${profile.keywords[i]} ]`,
          );
          return true;
        }
      }
      console.debug(
        `Title [${title}] didn't include the keywords - ${profile.keywords}`,
      );
      return false;
    case "organizers":
      for (let i = 0; i < profile.organizers.length; i++) {
        if (organizer == profile.organizers[i]) {
          console.log(`Organizer [${organizer}] matched`);
          return true;
        }
      }
      console.debug(`Organizer [${organizer}] wasn't matched`);
      return false;
    case "default":
      return true;
  }

  console.debug(`Profile [${profile.name}] was not a match`);
  return false;
}

function detectMTR() {
  return xapi.Command.MicrosoftTeams.List({ Show: "Installed" })
    .then(() => true)
    .catch(() => false);
}

class workspaceMonitor {
  constructor(booking, profile, mtr, externalLogging, debug = false) {
    workspaceBookingIds.push(booking.Id);
    console.log(
      `'New Workspace Monitor created for Booking Id [${booking.Id}] using Profile [${profile.name}]`,
    );
    if (!profile.monitor) return;

    this._mtr = mtr;
    this._externalLogging = externalLogging;
    this._debugging = debug;
    this._booking = booking;
    this._profile = profile;
    this._unbookTimer = null;
    this._unbookAlertTimer = null;
    this._monitors = [];
    // this._startMonitoringTimer = setTimeout(this._startMonitoring.bind(this), profile.startMonitoringAfter * 60 * 1000)
    // this._stopMonitoringTimer = setTimeout(this._stopMonitoring.bind(this), profile.stopMonitoringAfter * 60 * 1000)
    // this._monitors.push(xapi.Event.Bookings.End.on(bookingEvt => this._processBookingEnd(bookingEvt.id)));
  }

  initialize() {
    console.log(
      `Workspace Monitor Initialized for Booking Id [${this._booking.Id}]`,
    );
    this._startMonitoringTimer = setTimeout(
      () => this._startMonitoring(),
      this._profile.startMonitoringDelay * 60 * 1000,
    );
    this._stopMonitoringTimer = setTimeout(
      () => this._stopMonitoring(),
      this._profile.stopMonitoringAfter * 60 * 1000,
    );
    this._monitors.push(
      xapi.Event.Bookings.End.on((event) => this._processBookingEnd(event.id)),
    );
  }

  _processBookingEnd(bookingId) {
    if (bookingId != this._booking.Id) return;
    console.log(`Booking Id [${bookingId}] had ended`);
    this._stopMonitoring();
  }

  _startUnbookingCountdown() {
    console.log(
      `Starting Unbooking Countdown for Booking Id [${this._booking.Id}] - [${this._profile.requiredUnoccupiedDuration}] minutes`,
    );
    this._clearUnbookTimers();
    this._unbookTimer = setTimeout(
      this._unbook.bind(this),
      this._profile.requiredUnoccupiedDuration * 60 * 1000,
    );

    // Only schedule the pre-unbook alert when the profile defines a valid
    // alert window. Profiles without alertBeforeUnbookingDuration (e.g. the
    // default profile) would otherwise compute NaN and fire immediately.
    const alertBefore = this._profile.alertBeforeUnbookingDuration;
    if (
      alertBefore > 0 &&
      alertBefore <= this._profile.requiredUnoccupiedDuration
    ) {
      const alertDelayMinutes =
        this._profile.requiredUnoccupiedDuration - alertBefore;
      this._unbookAlertTimer = setTimeout(
        this._unbookAlert.bind(this),
        alertDelayMinutes * 60 * 1000,
      );
    }
  }

  _stopUnbookingCountdown() {
    console.log("Stopping countdown");
    const targets = ["Controller", "RoomScheduler", "OSD"];
    for (var t of targets) {
      xapi.Command.UserInterface.Message.Prompt.Clear({
        FeedbackId: `unbookingprompt-${t}`,
      });
    }
    this._clearUnbookTimers();
    // The booking is assumed accepted by default, so no action is needed to
    // keep it. We only cancel the pending unbooking countdown.
  }

  _unbookSummary() {
    const now = new Date();
    const start = new Date(this._booking.Time.StartTime);
    const end = new Date(this._booking.Time.EndTime);
    const toMinutes = (ms) => Math.max(0, Math.round(ms / (60 * 1000)));

    const bookingDuration = toMinutes(end.getTime() - start.getTime());
    const unbookedTime = toMinutes(now.getTime() - start.getTime());
    // The unbook only fires after the full unoccupied countdown, so the time
    // the room was actually used is the elapsed time minus that countdown.
    const usedTime = Math.max(
      0,
      unbookedTime - this._profile.requiredUnoccupiedDuration,
    );
    const savedTime = toMinutes(end.getTime() - now.getTime());

    return (
      `No Presence in Room detected - Matched Profile [${this._profile.name}]` +
      ` - Booking Duration: [${bookingDuration} minutes]` +
      ` - Used Time: [${usedTime} minutes]` +
      ` - Unbooked Time [${unbookedTime} minutes]` +
      ` - Saved Time [${savedTime} minutes]`
    );
  }

  async _unbook() {
   

    const summary = this._unbookSummary();


    let resultText = this._debugging
    ? " - Debug Mode ( No Action Taken )"
    : "";

    if (!this._debugging) {
      console.log(
        `Declining Booking - Meeting Id [${this._booking.MeetingId}]...`,
      );
      try {
        const result = await xapi.Command.Bookings.Respond({
          Comment: summary,
          MeetingId: this._booking.MeetingId,
          Type: "Decline",
        });

        resultText = `Booking Declined - Meeting Id [${this._booking.MeetingId}] - Result: ${JSON.stringify(result)}`;
      
      } catch (declineError) {
        console.debug(`Decline Booking failed - ${JSON.stringify(declineError)}`);
        console.log(`Trying Delete Booking - Meeting Id [${this._booking.MeetingId}]...`);
        try {
          const result = await xapi.Command.Bookings.Delete({
            MeetingId: this._booking.MeetingId,
          });
          resultText = `Booking Deleted - Meeting Id [${this._booking.MeetingId}] - Result: ${JSON.stringify(result)}`;
        } catch (deleteError) {
          resultText = `Booking Delete failed - Meeting Id [${this._booking.MeetingId}] - Result: ${JSON.stringify(deleteError)}`;
        }
      }
    }

    this._reportMacroAction(
      `Unbooking Booking Id [${this._booking.Id}]] - Meeting Id [${this._booking.MeetingId}] - ${summary} - Result: ${resultText}`,
    );


    this._stopMonitoring();
  }

  _unbookAlert() {
    console.log(`Displaying unbook alert for Booking Id [${this._booking.Id}]`);
    var plural = "s";
    if (this._profile.alertBeforeUnbookingDuration === 1) {
      plural = "";
    }

    xapi.Command.Bookings.Get({ Id: this._booking.Id })
      .then((result) => {
        const targets = ["Controller", "RoomScheduler", "OSD"];
        for (var t of targets) {
          xapi.Command.UserInterface.Message.Prompt.Display({
            Duration: 30,
            FeedbackId: `unbookingprompt-${t}`,
            Title: "No Presence Detected",
            Text: `Booking [${this._booking.Title}] will be Unbooked in [${this._profile.alertBeforeUnbookingDuration}] minute${plural}`,
            "Option.1": "Don't Unbook",
            Target: t,
          });
        }
      })
      .catch((e) => {
        console.warn("Prompt not displayed. Booking no longer exists.");
      });
  }

  async _checkPresence() {
    console.log("Checking Presence");
    let detections = [];
    const monitor = config.presenceDetection;
    // if precence detected - stop countdown
    // if no precence detected - start countdown

    if (monitor.activeCalls) {
      const call = await xapi.Status.Call.get();
      const activeCall = call?.[0]?.Status == "Connected";
      detections.push({ type: "activeCall", result: activeCall });

      if (this._mtr) {
        const activeMTRCall =
          await xapi.Status.MicrosoftTeams.Calling.InCall.get()
            .then((result) => result == "True")
            .catch(() => false);
        detections.push({ type: "activeMTRCall", result: activeMTRCall });
      }
    }

    if (monitor.presentation) {
      const presentation = await xapi.Status.Conference.Presentation.get();
      const mode = presentation?.Mode ? presentation?.Mode != "Off" : false;
      const localInstance = presentation?.LocalInstance
        ? presentation?.LocalInstance.length > 0
        : false;
      const presenting = mode || localInstance;
      detections.push({ type: "presentation", result: presenting });
    }

    if (monitor.peopleCount) {
      const peopleCount =
        await xapi.Status.RoomAnalytics.PeopleCount.Current.get();
      detections.push({ type: "peopleCount", result: peopleCount > 0 });
    }

    if (monitor.peoplePresence) {
      const peoplePresence =
        await xapi.Status.RoomAnalytics.PeoplePresence.get();
      detections.push({
        type: "peoplePresence",
        result: peoplePresence == "Yes",
      });
    }

    const presenceDetected =
      detections.filter((detection) => detection.result).length > 0;

    if (presenceDetected) {
      console.log("Presence Detected");
      this._stopUnbookingCountdown();
    } else {
      console.log("No Presence Detected");
      this._startUnbookingCountdown();
    }
  }

  // This function subscribes to status changes and events based off macro config
  _startMonitoring() {
    console.log(
      `Starting Workspace Monitor for Booking Id [${this._booking.Id}]`,
    );
    const monitor = config.presenceDetection;
    if (monitor.activeCalls) {
      console.debug(
        `Monitoring Active Calls for Booking Id [${this._booking.Id}]`,
      );
      this._monitors.push(
        xapi.Status.SystemUnit.State.NumberOfActiveCalls.on(() =>
          this._checkPresence(),
        ),
      );
      if (this._mtr) {
        console.debug(
          `Monitoring MTR Calls for Booking Id [${this._booking.Id}]`,
        );
        this._monitors.push(
          xapi.Status.MicrosoftTeams.Calling.InCall.on(() =>
            this._checkPresence(),
          ),
        );
      }
    }

    if (monitor.presentation) {
      console.debug(
        `Monitoring Presentations for Booking Id [${this._booking.Id}]`,
      );
      this._monitors.push(
        xapi.Event.PresentationStarted.on(() => this._checkPresence()),
      );
    }

    if (monitor.peopleCount) {
      console.debug(
        `Monitoring PeopleCount for Booking Id [${this._booking.Id}]`,
      );
      this._monitors.push(
        xapi.Status.RoomAnalytics.PeopleCount.Current.on(() =>
          this._checkPresence(),
        ),
      );
    }

    if (monitor.peoplePresence) {
      console.debug(
        `Monitoring PeoplePresence for Booking Id [${this._booking.Id}]`,
      );
      this._monitors.push(
        xapi.Status.RoomAnalytics.PeoplePresence.on(() =>
          this._checkPresence(),
        ),
      );
    }

    this._monitors.push(
      xapi.Event.UserInterface.Message.Prompt.Response.on((event) => {
        console.log("Message.Prompt.Response:");
        console.log(event);
        if (
          event.FeedbackId.indexOf("unbookingprompt") >= 0 &&
          event.OptionId === "1"
        ) {
          this._stopUnbookingCountdown();
        }
      }),
    );
    if (monitor.guiInteractions) {
      console.debug(
        `Monitoring GUI Interactions for Booking Id [${this._booking.Id}]`,
      );
      this._monitors.push(
        xapi.Event.UserInterface.Extensions.Panel.Clicked.on(() =>
          this._stopUnbookingCountdown(),
        ),
      );
      this._monitors.push(
        xapi.Event.UserInterface.Extensions.Widget.Action.on(() =>
          this._stopUnbookingCountdown(),
        ),
      );
    }

    this._checkPresence();
  }

  // This function unsubscribes from all status changes and events
  _stopMonitoring() {
    console.log(
      `Stopping Workspace Monitor for Booking Id [${this._booking.Id}]`,
    );
    this._clearMonitors();
    this._clearUnbookTimers();
    clearTimeout(this._startMonitoringTimer);
    clearTimeout(this._stopMonitoringTimer);

    const index = workspaceBookingIds.indexOf(this._booking.Id);
    if (index > -1) {
      workspaceBookingIds.splice(index, 1);
    }
  }

  _clearMonitors() {
    for (let i = 0; i < this._monitors.length; i++) {
      this._monitors[i]();
      this._monitors[i] = () => void 0;
    }
    this._monitors = [];
  }

  _clearUnbookTimers() {
    clearTimeout(this._unbookTimer);
    clearTimeout(this._unbookAlertTimer);
  }

  async _reportMacroAction(action) {
    console.log(action);
    if (!this._externalLogging.enabled) return;

    console.log("Preparing to send external logging report");

    const workspaceName =
      await xapi.Status.UserInterface.ContactInfo.Name.get();

    let payload = {};

    if (this._externalLogging.type == "webhook") {
      payload = {
        workspaceName,
        bookingTile: this._booking.Title,
        profile: this._profile.name,
        action: action,
      };
    } else {
      const message =
        "Unbooking Macro Event:\n" +
        "Workspace Name: " +
        workspaceName +
        "\n" +
        "Booking Title: " +
        this._booking.Title +
        "\n" +
        "Monitoring Profile: " +
        this._profile.name +
        "\n" +
        "Final Action: " +
        action;

      payload = { text: message };

      const emailRegex = /^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$/g;
      if (emailRegex.test(this._externalLogging.contact)) {
        payload.toPersonEmail = this._externalLogging.contact;
      } else {
        payload.roomId = this._externalLogging.contact;
      }
    }

    console.log(`Sending Event ${payload} to [${this._externalLogging.url}]`);

    const Header = [
      "Content-Type: application/json",
      "Authorization: Bearer " + this._externalLogging.token,
    ];

    if (this._externalLogging.type == "webhook") {
      return xapi.Command.HttpClient.Post(
        { Header, Url: this._externalLogging.url },
        JSON.stringify(payload),
      )
        .then((result) => result?.Body)
        .catch((error) =>
          console.log(
            `Error Sending Macros Action Data to URL: ${this._externalLogging.url} -`,
            error,
          ),
        );
    } else {
      return xapi.Command.HttpClient.Post(
        { Header, Url: this._externalLogging.url },
        JSON.stringify(payload),
      )
        .then((result) => result?.Body)
        .then((result) => result?.id)
        .catch((error) =>
          console.log(
            `Error Sending Macros Action Data to Webex Contact: ${this._externalLogging.contact} -`,
            error,
          ),
        );
    }
  }
}

// Config export for testing only
export { config };
