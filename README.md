# Unbook Workspace Macro

This is an example enhanced unbooking macro for empty Webex workspaces with added support of handling bookings of different durations, meeting title and organizer names along with a simple telemetry logging feature to track the macros behaviour.

## Overview

This macro monitors booking start event and applies a room monitoring unbooking policy based on booking duration, meeting title keywords or organizer name. Below is a breakdown of the how it can handle each type of matched meeting.

### Duration

In this example duration profile, if the booking has a duration from 0 to 60 minutes, then the macro will begin to monitor the workspace for people presence immedicably and will only monitor the workspace the first 10 minutes of the meeting. If no one detected for 5 minutes during that time, then the macro will unbook the meeting.

```js
 {
  type: 'duration',         // Profile type: duration | keywords | organizer
  name: 'Short Meetings',   // Name of profile for logging
  duration: [0, 60],        // Duration of booking in minutes: From zero minutes to 60 minute meetings
  monitor: true,            // Enable monitoring for this matched profile
  startMonitoringDelay: 0,  // Number of minutes after the booking starts in which to begin monitoring
  stopMonitoringAfter: 10,  // Number of minutes after the booking starts in which to stop monitoring
  requiredUnoccupiedDuration: 5,    // Number of minutes the workspace is unoccupied before unbooking
  alertBeforeUnbookingDuration: 1   // Number of minutes before unbooking in which to alert user
}
```

### Meeting Title Keywords

We may want to leave high priority bookings unmonitored so we can use the Meeting Title keyboard profile to disable monitoring in these cases. It is also possible to monitor these booking with different timer delays if required

```js
{
  type: 'keywords',               // Profile type: Keywords
  name: 'Meeting Title Keyword',  // Name of profile for logging
  keywords: ['Training', 'Test'], // Array of keywords in which to look for in the booking title
  monitor: false                  // Disable monitoring for these matched bookings
}
```
 
### Organizer

We may want to disable monitoring for bookings made by specific organisers. This example profile will check the Organiser name associated with the booking.

```js
{
  type: 'organizers',             // Profile type: Keywords
  name: 'Organizers Name',        // Name of profile for logging
  organizers: ['William Mills'],  // Array of organizer names to match with bookings
  monitor: false                  // Disable monitoring for these matched booking
}
```

### Flow Diagram

![Unbook Workspace Macro](https://github.com/wxsd-sales/unbook-workspace-macro/assets/21026209/b99a69ac-9e65-481f-af86-48dee2598eee)



## Setup

### Prerequisites & Dependencies: 

- Webex Device with RoomOS 11.x or above
- Web admin access to the device to upload the macro


### Installation Steps:

1. Download the ``unbook-workspace.js`` file and upload it to your Webex Devices Macro editor via the web interface.
2. Configure the macros monitoring policies, presence detection and external logging (optional) as required, there are comments for each field to help with the configuration.
3. Enable the Macro on the editor.

## Demo

*For more demos & PoCs like this, check out our [Webex Labs site](https://collabtoolbox.cisco.com/webex-labs).

## License

All contents are licensed under the MIT license. Please see [license](LICENSE) for details.


## Disclaimer

Everything included is for demo and Proof of Concept purposes only. Use of the site is solely at your own risk. This site may contain links to third party content, which we do not warrant, endorse, or assume liability for. These demos are for Cisco Webex use cases, but are not Official Cisco Webex Branded demos.


## Questions
Please contact the WXSD team at [wxsd@external.cisco.com](mailto:wxsd@external.cisco.com?subject=RepoName) for questions. Or, if you're a Cisco internal employee, reach out to us on the Webex App via our bot (globalexpert@webex.bot). In the "Engagement Type" field, choose the "API/SDK Proof of Concept Integration Development" option to make sure you reach our team. 
