/********************************************************
 * 
 * Macro Author:      	William Mills
 *                    	Technical Solutions Specialist 
 *                    	wimills@cisco.com
 *                    	Cisco Systems
 * 
 * Version: 1-0-0
 * Released: 10/20/23
 * 
 * This example macro release empty workspace booking based
 * off configurable policies. Additional for auditing, this
 * macro can log the actions take so admin can review which 
 * booking were released or not and get a better insight into
 * how their workspaces are being used.
 * 
 * Full Readme, source code and license details for this macro are available 
 * on Github: https://github.com/wxsd-sales/unbook-workspace-macro
 * 
 ********************************************************/

import xapi from 'xapi';

/*********************************************************
 * Configure the settings below
**********************************************************/

const config = {
  profiles: [
    {
      type: 'duration',         // Profile type: duration |
      name: 'Short Meetings',   // Name of profile for logging
      duration: [0, 60],        // Duration in minutes
      monitor: true,
      startMonitoringDelay: 0,
      stopMonitoringAfter: 10,
      requiredUnoccpiedDuration: 5,
      alertBeforeUnbookingDuration: 1

    },
    {
      type: 'duration',
      name: 'Between 1 and 2 hour Meetings',
      duration: [61, 180],
      policy: 'long meetings',
      monitor: true,
      startMonitoringDelay: 0,
      stopMonitoringAfter: 10,
      requiredUnoccpiedDuration: 5,
    },
    {
      type: 'duration',
      name: 'All day meetings - Don\'t monitor during lunch hours',
      duration: [180, 480],
      monitor: true,
      startMonitoringDelay: 0,
      stopMonitoringAfter: 10,
      requiredUnoccpiedDuration: 5
    },
    {
      type: 'keywords',
      name: 'Meeting Title Keyword',
      keywords: ['Training', 'Test'],
      monitor: true
    },
    {
      type: 'organizers',
      name: 'Organizers Name',
      organizers: ['John Smyth'],
      monitor: true
    },
    {
      type: 'default',
      name: 'Default Booking Handling Profile',
      organizers: ['John Smyth'],
      monitor: true
    }],
  presenceDetection: {
    activeCalls: true,
    presentation: true,
    peopleCount: true,
    peoplePresence: true,
    presenceAndPeopleCount: false,
    guiInteractions: true
  },
  externalLogging: {
    enabled: false,
    url: 'https://<your logging server>',
    token: '<logging server access token>'
  },
  debugging: true
}


/*********************************************************
 * Do not change below
**********************************************************/

xapi.Event.Bookings.Start.on(event => processBookingStart(event));

processBookingStart({ Id: 'webex-1' })
async function processBookingStart(bookingEvt) {
  const bookingId = bookingEvt.Id;
  console.log('Booking Start Event:', bookingId);
  const booking = await xapi.Command.Bookings.Get({ Id: bookingId })
    .then(result => result.Booking)
    .catch(() => console.log('Cound not get find meeting: ', bookingId))

  if (!booking) return
  console.log('Booking Details: ', JSON.stringify(booking))
  const profile = mapToProfile(booking);
  if (!profile) return

  new workspaceMonitor(booking, profile)
  
}

function mapToProfile(booking) {
  const profiles = config.profiles;
  const startTime = new Date(booking.Time.StartTime)
  const endTime = new Date(booking.Time.EndTime)
  const duration = endTime.getMinutes() - startTime.getMinutes();
  const title = booking.Title;
  const organizer = booking.Organizer.FirstName;
  const profile = profiles.find(profile => compareProfile(profile, duration, title, organizer))
  console.log('Matched Profile: ', profile)
  return profile
}

function compareProfile(profile, duration, title, organizer) {
  console.debug(`Comparing Profile [${profile.name}] using type [${profile.type}] with - Duration [${duration}] - Title [${title}] - Organizer [${organizer}]`)
  switch (profile.type) {
    case 'duration':
      if (duration > profile.duration[0] && duration <= profile.duration[1]) {
        console.debug(`Duration [${duration}] is within range [ ${profile.duration[0]} - ${profile.duration[1]} ]`)
        return true;
      } else {
        console.debug(`Duration [${duration}] is not with range [ ${profile.duration[0]} - ${profile.duration[1]} ]`)
        return false
      }
    case 'keywords':
      for (let i = 0; i < profile.keywords.length; i++) {
        if (title.includes(profile.keywords[i])) {
          console.debug(`Title [${title}] includes keyword [ ${profile.keywords[i]} ]`)
          return true
        }
      }
      console.debug(`Title [${title}] didn't include the keywords - ${profile.keywords}`)
      return false
    case 'organizers':
      for (let i = 0; i < profile.organizers.length; i++) {
        if (organizer == profile.organizers[i]) {
          console.log(`Organiser [${organizer}] matched`)
          return true
        }
      }
      console.debug(`Organiser [${organizer}] wasn't matched`)
      return false
  }

  console.debug(`Profile [${profile.name}] was not a match`)
  return false;
}


class workspaceMonitor {
  constructor(booking, profile) {
    console.log(`'New Workspace Monitor started for Booking Id [${booking.Id}] using Profile [${profile.name}]`)
    if (!profile.monitor) return
    this._booking = booking;
    this._profile = profile;
    this._startMonitoringTimer = null;
    this._stopMonitoringTimer = null;
    this._unbookTimer = null;
    this._unbookAlertTimer = null;
    this._monitors = [];
    this._startMonitoringTimer = setTimeout(this._startMonitoring.bind(this), profile.startMonitoringAfter * 60 * 1000)
    this._stopMonitoringTimer = setTimeout(this._stopMonitoring.bind(this), profile.stopMonitoringAfter * 60 * 1000)
    this._monitors.push(xapi.Event.Bookings.End.on(bookingEvt => this._processBookingEnd(bookingEvt.id)));
  }

  _processBookingEnd(bookingId) {
    if (bookingId != this.booking.Id) return;
    console.log(`Booking Id [${bookingId}] had ended`);
    this._stopMonitoring();
  }

  _processEvent(type) {
    if (this._unbookTimer == null) {
      console.log(`Event [${type}] occuried - not active timer`)
    } else {
      console.log(`Event [${type}] occuried - active timer - resetting`)
      this._startCountdown();
    }
  }

  _startCountdown() {

    console.log(`Starting Unbooking Countdown for Booking Id [${this._booking.Id}] - [${this._profile.requiredUnoccpiedDuration}] minutes`)
    this._clearUnbookTimers();
    this._unbookTimer = setTimeout(this._unbook.bind(this), this._profile.requiredUnoccpiedDuration * 60 * 1000);
    this._unbookAlertTimer = setTimeout(this._unbookAlert.bind(this), (this._profile.requiredUnoccpiedDuration - this._profile.alertBeforeUnbookingDuration)* 60 * 1000);
  }

  _stopCountdown() {
    console.log('Stopping countdown')
    xapi.Command.UserInterface.Message.Prompt.Clear({ FeedbackId: 'unbookingprompt' });
    this._clearUnbookTimers();
  }

  _unbook() {
    console.log(`Unbooking Booking Id [${this._booking.Id}]] - Meeting Id [${this._booking.MeetingId}]`);
    this._reportMacroAction(`Unbooking Booking Id [${this._booking.Id}]] - Meeting Id [${this._booking.MeetingId}]`)

    return;
    xapi.Command.Bookings.Respond({ MeetingId: this._booking.MeetingId, Type: 'Decline' })
      .then(value => {
        console.log(value)
      })
      .catch(error => {
        console.warn(error)
        this._reportMacroAction(`Unable to Unbook Booking Id [${this._booking.Id}]] - Meeting Id [${this._booking.MeetingId}] - Message:`, error.message)
      })

    this._stopMonitoring();
  }

  _unbookAlert() {
    console.log(`Displaying soon to unbook alert for Booking Id [${this._booking.Id}]`);

    xapi.Command.UserInterface.Message.Prompt.Display({
      Duration: 30,
      FeedbackId: 'unbookingprompt',
       Title: 'No Presence Detected' ,
       Text: `Booking [${this._booking.Title}] will be Unbooked in [${this._profile.alertBeforeUnbookingDuration}] minutes`,
      "Option.1": 'Don\'t Unbook' 
      });
    
  }

  async _checkPresence() {
    console.log('Checking Presence');
    let detections = [];
    const monitor = config.presenceDetection;
    // if precence detected - stop countdown
    // if no precence detected - start countdown

    if (monitor.activeCalls) {
      const numOfCalls = await xapi.Status.SystemUnit.State.NumberOfActiveCalls.get();
      detections.push({ type: 'numOfCalls', result: (numOfCalls > 0) });
    }

    if (monitor.presentation) {
      // Query Presentations

      // dections.push({presentations: (numOfCalls > 0) });

    }

    if (monitor.peopleCount) {
      const peopleCount = await xapi.Status.RoomAnalytics.PeopleCount.get();
      detections.push({ type: 'peopleCount', result: peopleCount > 0 });
    }

    if (monitor.peoplePresence) {
      const peoplePresence = await xapi.Status.RoomAnalytics.PeoplePresence.get();
      detections.push({ type: 'peoplePresence', result: peoplePresence == 'Yes' });
    }

    console.log(detections)

    const presenceDetected = detections.filter((detection) => detection.result).length > 0;
    
    if(presenceDetected){
      console.log('Presence Detected')
      this._stopCountdown();
    } else { 
      console.log('No Presence Detected')
      this._startCountdown();
    }
  }

  // This function subscribes to status changes and events based off macro config
  _startMonitoring() {
    console.log(`Starting Workspace Monitor for Booking Id [${this._booking.Id}]`)
    const monitor = config.presenceDetection;
    if (monitor.activeCalls) {
      console.debug(`Monitoring Active Calls for Booking Id [${this._booking.Id}]`)
      this._monitors.push(xapi.Status.SystemUnit.State.NumberOfActiveCalls.on(value => this._processPresence('ActiveCalls', value)));
    }

    if (monitor.presentation) {
      console.debug(`Monitoring Presentations for Booking Id [${this._booking.Id}]`)
      this._monitors.push(xapi.Event.PresentationStarted.on(value => {
        console.log('Presentation started', value);
        }));
    }

    if (monitor.peopleCount) {
      console.debug(`Monitoring PeopleCount for Booking Id [${this._booking.Id}]`)
      this._monitors.push(xapi.Status.RoomAnalytics.PeopleCount.on(value => {
        console.log('People PeopleCount', value);
        this._checkPresence();
      }));
    }

    if (monitor.peoplePresence) {
      console.debug(`Monitoring PeoplePresence for Booking Id [${this._booking.Id}]`)
      this._monitors.push(xapi.Status.RoomAnalytics.PeoplePresence.on(value => {
        console.log('People Presence', value)
        this._checkPresence();
      }));
    }

    if (monitor.guiInteractions) {
      console.debug(`Monitoring GUI Interactions for Booking Id [${this._booking.Id}]`)
      this._monitors.push(xapi.Event.UserInterface.Extensions.Panel.Clicked.on(value => console.log('Panel Clicked')));
      this._monitors.push(xapi.Event.UserInterface.Extensions.Widget.Action.on(value => console.log('Widget Action')));
    }

    this._checkPresence();

  }

  // This function unsubscribe from all status changes and events
  _stopMonitoring() {
    console.log(`Stopping Workspace Monitor for Booking Id [${this._booking.Id}]`)
    this._clearMonitors();
    this._clearUnbookTimers();
    this._reportMacroAction('Macro stopped monitoring workspace without unbooking')
  }

  _clearMonitors() {
    for (let i = 0; i < this._monitors.length; i++) {
      this._monitors[i]();
      this._monitors[i] = () => void 0;
    }
    this._monitors = [];
  }

  _clearUnbookTimers(){
     if (this._unbookTimer != null) {
      clearTimeout(this._unbookTimer)
    }

    if (this._unbookAlertTimer != null) {
      clearTimeout(this._unbookAlertTimer)
    }
  }


  _reportMacroAction(action) {
    if (!config.externalLogging.enabled) return

    const server = config.externalLogging;
    const payload = {
      bookingTile: this._booking.Title,
      profile: this._profile.name,
      action: action
    }

    console.log(`Sending Event ${payload} to [${server.url}]`)

    const Header = [
      'Content-Type: application/json',
      'Authorization: Bearer ' + server.token]

    xapi.Command.HttpClient.Post({
      Header,
      Url: server.url,
    }, JSON.stringify(payload))
  }

}
