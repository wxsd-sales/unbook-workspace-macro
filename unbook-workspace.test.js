import {
  afterEach,
  beforeEach,
  describe,
  expect,
  it,
  jest,
} from "@jest/globals";

const macroName = "./unbook-workspace.js";


function mockCoreStatus(xapi) {
  xapi.Status.Network[1].IPv4.Address.set("1.1.1.1");
  xapi.Status.Network[1].IPv6.Address.set("::1");
  xapi.Status.Webex.DeveloperId.set("12345678890abcdef");
  xapi.Status.UserInterface.ContactInfo.Name.set("Workspace Name");
  xapi.Status.SystemUnit.State.NumberOfActiveCalls.set(0);
  xapi.Status.MicrosoftTeams.Calling.InCall.set("False");
  xapi.Config.Bookings.CheckIn.Enabled.set("False");
  mockCurrentBooking(xapi, undefined);

  mockMTRMode(xapi, false);
  mockRoomAnalytics(xapi, 0);
}

function mockRoomAnalytics(xapi, peopleCount = 0) {
  if (peopleCount > 0) {
    xapi.Status.RoomAnalytics.PeopleCount.Current.set(peopleCount);
    xapi.Status.RoomAnalytics.PeoplePresence.set("Yes");
    xapi.Status.RoomAnalytics.RoomInUse.set("True");
  } else {
    xapi.Status.RoomAnalytics.PeopleCount.Current.set(0);
    xapi.Status.RoomAnalytics.PeoplePresence.set("No");
    xapi.Status.RoomAnalytics.RoomInUse.set("False");
  }
}

function mockMTRMode(xapi, isMTR = false) {
  if (isMTR) {
    xapi.Command.MicrosoftTeams.List.mockResolvedValue({
      Entry: [{ Name: "MicrosoftTeamsRooms", id: "1", multiple: "true" }],
      status: "OK",
    });
  } else {
    xapi.Command.MicrosoftTeams.List.mockRejectedValue({
      code: 1,
      message: "No platforms found",
    });
  }
}

function createMockBooking(xapi, options = {}) {

  const { duration = 60, title = "This current meeting" , organizerName = "Person Name" } = options;
  const organizerEmail = `${organizerName.toLowerCase().replace(/ /g, '.')}@example.com`;
  const spaceIndex = organizerName.indexOf(" ");
  const firstName = spaceIndex === -1 ? organizerName : organizerName.slice(0, spaceIndex);
  const lastName = spaceIndex === -1 ? "" : organizerName.slice(spaceIndex + 1);

  const start = new Date();
  const end = new Date(start.getTime() + (duration * 60 * 1000));

  console.log('Creating mock booking:', start, end);
  const booking = {
    Agenda: "",
    BookedFromDevice: "False",
    Cancellable: "False",
    Id: "webex-1",
    MaximumMeetingExtension: "30",
    MeetingId: crypto.randomUUID(),
    Organizer: {
      Email: organizerEmail,
      FirstName: firstName,
      Id: crypto.randomUUID(),
      LastName: lastName,
    },
    Time: {
      EndTime:  end.toISOString(),
      EndTimeBuffer: "0",
      IsAllDay: "False",
      StartTime: start.toISOString(),
      StartTimeBuffer: "300",
    },
    Title: title,
    id: "1",
    multiple: "true",
  };

  console.log('Creating mock booking:', booking);

  mockCurrentBooking(xapi, booking);

  return booking;
}

function mockCurrentBooking(xapi, booking) {

  const now = new Date();
  const oneHourAgo = new Date( now - 60 * 60 * 1000).toISOString();

  if (typeof booking === "undefined") {
    
    xapi.Status.Bookings.Current.Id.set("");
    xapi.Status.Bookings.Availability.Status.set("Free");
    xapi.Status.Bookings.Availability.TimeStamp.set("");
    xapi.Command.Bookings.List.mockResolvedValue({
      LastUpdated: oneHourAgo,
      ResultInfo: { TotalRows: "0" },
      status: "OK",
    });

    xapi.Command.Bookings.Get.mockRejectedValue({"code":1,"message":"Not found"});

  } else {
    const {Id, MeetingId, Time} = booking;
    xapi.Status.Bookings.Current.Id.set(Id);
    xapi.Status.Bookings.Availability.Status.set("BookedUntil");
    xapi.Status.Bookings.Availability.TimeStamp.set(Time.EndTime);

    xapi.Command.Bookings.Get.mockResolvedValue({
      Booking: booking,
      status: "OK",
    });

    xapi.Command.Bookings.List.mockResolvedValue({
      Booking: [booking],
      LastUpdated: oneHourAgo,
      ResultInfo: {
        TotalRows: "1",
      },
      status: "OK",
    });

    xapi.Event.Bookings.Start.emit({ Id, MeetingId });
  }
}

async function loadMacro(xapi) {
  await import(macroName);
  await flushPromises();
}

async function flushPromises() {

  for (let i = 0; i < 20; i += 1) await Promise.resolve();
}

// Deterministically drains the macro's timer + microtask work. Each round
// flushes pending promises (so async callbacks can schedule their timers),
// then advances the fake clock by one step. It stops once no timers remain,
// so it doesn't depend on any specific profile's timing values.
async function settleTimers({ maxRounds = 60, stepMs = 60 * 1000 } = {}) {
  for (let i = 0; i < maxRounds; i += 1) {
    await flushPromises();
    if (jest.getTimerCount() === 0) break;
    jest.advanceTimersByTime(stepMs);
  }
  await flushPromises();
}

function matchedProfiles(logSpy) {
  return logSpy.mock.calls
    .filter((c) => c[0] === "Matched Profile: " && c[1])
    .map((c) => c[1]);
}


describe("Unbooking Workspace Macro Tests", () => {
  let logSpy;
  let debugSpy;
  let errorSpy;
  let warnSpy;

  beforeEach(() => {
    jest.resetModules();
    jest.useFakeTimers();
    logSpy = jest.spyOn(console, "log").mockImplementation(() => {});
    debugSpy = jest.spyOn(console, "debug").mockImplementation(() => {});
    errorSpy = jest.spyOn(console, "error").mockImplementation(() => {});
    warnSpy = jest.spyOn(console, "warn").mockImplementation(() => {});
  });

  afterEach(async () => {
    // Run pending timers, then let any async monitoring work (e.g. the
    // async _checkPresence chain) settle while fake timers are still
    // installed, so trailing clearTimeout calls don't run after teardown.
    jest.runOnlyPendingTimers();
    await flushPromises();
    jest.clearAllTimers();
    jest.restoreAllMocks();
    jest.useRealTimers();
  });

  it("Macro should load successfully", async () => {
    const { default: xapi } = await import("xapi");
    xapi.reset();
    mockCoreStatus(xapi);
    await loadMacro(xapi);

    expect(xapi.Command.Bookings.List).toHaveBeenCalled();
  });


  it("Macro should not load if CheckIn is enabled", async () => {
    const { default: xapi } = await import("xapi");
    xapi.reset();
    mockCoreStatus(xapi);
    xapi.Config.Bookings.CheckIn.Enabled.set("True");
    await loadMacro(xapi);

    expect(xapi.Command.Bookings.List).not.toHaveBeenCalled();
  });

  it("Macro should disable itself if CheckIn is enabled", async () => {
    const { default: xapi } = await import("xapi");
    xapi.reset();
    mockCoreStatus(xapi);
   
    await loadMacro(xapi);
    expect(xapi.Command.Bookings.List).toHaveBeenCalled();
    jest.clearAllMocks()
    expect(xapi.Command.Bookings.List).not.toHaveBeenCalled();
    xapi.Config.Bookings.CheckIn.Enabled.set("True");
    await flushPromises();

    expect(xapi.Command.Bookings.List).not.toHaveBeenCalled();

    xapi.Config.Bookings.CheckIn.Enabled.set("False");
    await flushPromises();

    expect(xapi.Command.Bookings.List).toHaveBeenCalled();
  
  });

  it("Macro current booking if present", async () => {
    const { default: xapi } = await import("xapi");
    xapi.reset();
    mockCoreStatus(xapi);
    const booking = createMockBooking(xapi, { duration: 60 });

    await loadMacro(xapi);

    expect(await xapi.Status.Bookings.Current.Id.get()).toBe(booking.Id);

    expect(xapi.Command.Bookings.List).toHaveBeenCalled();

    return;
  });

  describe("Profile matching", () => {
    it.each([
      [30, "Short Meetings"],
      [120, "Between 1 and 2 hour Meetings"],
      [300, "All day meetings - Don't monitor during lunch hours"],
    ])("duration %i minutes matches %s", async (duration, expectedName) => {
      const { default: xapi } = await import("xapi");
      xapi.reset();
      mockCoreStatus(xapi);
      createMockBooking(xapi, { duration });

      await loadMacro(xapi);

      expect(matchedProfiles(logSpy).at(-1)?.name).toBe(expectedName);
    });

    it.each([
      [60, "Short Meetings"],
      [61, "Default Booking Handling Profile"],
      [180, "Between 1 and 2 hour Meetings"],
      [181, "All day meetings - Don't monitor during lunch hours"],
      [480, "All day meetings - Don't monitor during lunch hours"],
      [481, "Default Booking Handling Profile"],
    ])("boundary duration %i minutes matches %s", async (duration, expectedName) => {
      const { default: xapi } = await import("xapi");
      xapi.reset();
      mockCoreStatus(xapi);
      createMockBooking(xapi, { duration });

      await loadMacro(xapi);

      expect(matchedProfiles(logSpy).at(-1)?.name).toBe(expectedName);
    });

    it.each([
      ["Weekly Training Session"],
      ["System Test Run"],
    ])("keyword title %s matches the keyword profile", async (title) => {
      const { default: xapi } = await import("xapi");
      xapi.reset();
      mockCoreStatus(xapi);
      createMockBooking(xapi, { duration: 600, title });

      await loadMacro(xapi);

      const matched = matchedProfiles(logSpy).at(-1);
      expect(matched?.type).toBe("keywords");
      expect(matched?.name).toBe("Meeting Title Keyword");
    });

    it("organizer full name matches the organizers profile and is not monitored", async () => {
      const { default: xapi } = await import("xapi");
      xapi.reset();
      mockCoreStatus(xapi);
      createMockBooking(xapi, {
        duration: 600,
        title: "Project Sync",
        organizerName: "John Smith",
      });

      await expect(loadMacro(xapi)).resolves.not.toThrow();

      const matched = matchedProfiles(logSpy).at(-1);
      expect(matched?.type).toBe("organizers");
      expect(matched?.name).toBe("Organizers Name");
      expect(matched?.monitor).toBe(false);
    });

    it("non-matching booking falls back to the default profile", async () => {
      const { default: xapi } = await import("xapi");
      xapi.reset();
      mockCoreStatus(xapi);
      createMockBooking(xapi, {
        duration: 600,
        title: "Project Sync",
        organizerName: "Jane Doe",
      });

      await loadMacro(xapi);

      const matched = matchedProfiles(logSpy).at(-1);
      expect(matched?.type).toBe("default");
      expect(matched?.name).toBe("Default Booking Handling Profile");
    });

    it("duration takes precedence over a matching keyword", async () => {
      const { default: xapi } = await import("xapi");
      xapi.reset();
      mockCoreStatus(xapi);
      createMockBooking(xapi, { duration: 30, title: "Training" });

      await loadMacro(xapi);

      expect(matchedProfiles(logSpy).at(-1)?.name).toBe("Short Meetings");
    });

    it("duration takes precedence over a matching organizer", async () => {
      const { default: xapi } = await import("xapi");
      xapi.reset();
      mockCoreStatus(xapi);
      createMockBooking(xapi, { duration: 30, organizerName: "John Smith" });

      await loadMacro(xapi);

      expect(matchedProfiles(logSpy).at(-1)?.name).toBe("Short Meetings");
    });
  });

  describe("Presence and unbooking", () => {
    async function enableExternalLogging() {
      const { config } = await import(macroName);
      config.externalLogging.enabled = true;
      config.externalLogging.type = "bot";
      config.externalLogging.contact = "logs@example.com";
      config.externalLogging.url = "https://logging.example.test/hook";
      config.externalLogging.token = "test-token";
      config.debugging = false;
      return config;
    }

    it("releases the booking and calls the logging API when no presence is detected", async () => {
      const { default: xapi } = await import("xapi");
      xapi.reset();
      mockCoreStatus(xapi);
      mockRoomAnalytics(xapi, 0);
      await xapi.Config.HttpClient.Mode.set("On");
      xapi.Command.HttpClient.Post.mockResolvedValue({
        Body: '{"id":"message-1"}',
        StatusCode: "200",
        status: "OK",
      });

      // Load with no active booking so the monitor picks up the enabled logging config.
      await loadMacro(xapi);
      await enableExternalLogging();

      const booking = createMockBooking(xapi, { duration: 30 });
      await settleTimers();

      expect(xapi.Command.Bookings.Respond).toHaveBeenCalledWith(
        expect.objectContaining({
          Type: "Decline",
          MeetingId: booking.MeetingId,
        }),
      );

      // The decline comment should carry the unbooking detail.
      const declineCall = xapi.Command.Bookings.Respond.mock.calls.find(
        ([args]) => args?.Type === "Decline",
      );
      expect(declineCall).toBeDefined();
      expect(declineCall[0].Comment).toContain(
        "Matched Profile [Short Meetings]",
      );
      expect(declineCall[0].Comment).toContain("Booking Duration: [30 minutes]");
      expect(declineCall[0].Comment).toContain("Used Time:");
      expect(declineCall[0].Comment).toContain("Unbooked Time");
      expect(declineCall[0].Comment).toContain("Saved Time");

      // When Respond succeeds the booking is declined, so no Delete fallback
      // is needed.
      expect(xapi.Command.Bookings.Delete).not.toHaveBeenCalled();
      expect(xapi.Command.HttpClient.Post).toHaveBeenCalledWith(
        expect.objectContaining({
          Url: "https://logging.example.test/hook",
        }),
        expect.any(String),
      );
    });

    it("does not release or log when no external logging is configured", async () => {
      const { default: xapi } = await import("xapi");
      xapi.reset();
      mockCoreStatus(xapi);
      mockRoomAnalytics(xapi, 0);

      const booking = createMockBooking(xapi, { duration: 30 });
      await loadMacro(xapi);

      await settleTimers();

      // Booking is still released (declined) with default (disabled) logging...
      expect(xapi.Command.Bookings.Respond).toHaveBeenCalledWith(
        expect.objectContaining({
          Type: "Decline",
          MeetingId: booking.MeetingId,
        }),
      );
      // ...but the logging API must not be called.
      expect(xapi.Command.HttpClient.Post).not.toHaveBeenCalled();
    });

    it("does not release the booking or log when presence is detected", async () => {
      const { default: xapi } = await import("xapi");
      xapi.reset();
      mockCoreStatus(xapi);
      mockRoomAnalytics(xapi, 4);
  
      xapi.Command.HttpClient.Post.mockResolvedValue({
        Body: '{"id":"message-1"}',
        StatusCode: "200",
        status: "OK",
      });

      await loadMacro(xapi);
      await enableExternalLogging();

      createMockBooking(xapi, { duration: 30 });
      await settleTimers();

      expect(xapi.Command.Bookings.Delete).not.toHaveBeenCalled();
      expect(xapi.Command.HttpClient.Post).not.toHaveBeenCalled();
      // The booking is assumed accepted by default, so the macro takes no
      // Bookings.Respond action when presence is detected.
      expect(xapi.Command.Bookings.Respond).not.toHaveBeenCalled();
    });

    it("releases a default-profile booking without firing a malformed alert", async () => {
      // The default profile has no alertBeforeUnbookingDuration. Previously the
      // countdown computed NaN and fired an "undefined minutes" alert prompt
      // immediately. The booking should still be released, but no alert prompt
      // should be shown.
      const { default: xapi } = await import("xapi");
      xapi.reset();
      mockCoreStatus(xapi);
      mockRoomAnalytics(xapi, 0);

      const booking = createMockBooking(xapi, {
        duration: 600,
        title: "Project Sync",
        organizerName: "Jane Doe",
      });
      await loadMacro(xapi);

      // Confirm the default profile is the one being monitored.
      expect(matchedProfiles(logSpy).at(-1)?.name).toBe(
        "Default Booking Handling Profile",
      );

      await settleTimers();

      expect(xapi.Command.Bookings.Respond).toHaveBeenCalledWith(
        expect.objectContaining({
          Type: "Decline",
          MeetingId: booking.MeetingId,
        }),
      );
      expect(
        xapi.Command.UserInterface.Message.Prompt.Display,
      ).not.toHaveBeenCalled();
    });

    it("deletes by MeetingId when Respond throws because the device is the organizer", async () => {
      const { default: xapi } = await import("xapi");
      xapi.reset();
      mockCoreStatus(xapi);
      mockRoomAnalytics(xapi, 0);

      const booking = createMockBooking(xapi, { duration: 30 });

      // A device that owns the hybrid calendar mailbox can be the meeting
      // organizer, in which case Bookings.Respond is rejected.
      xapi.Command.Bookings.Respond.mockRejectedValue({
        code: 1,
        message:
          "You can't respond to this meeting because you're the meeting organizer.",
      });

      await loadMacro(xapi);
      await settleTimers();

      // Respond is always attempted first.
      expect(xapi.Command.Bookings.Respond).toHaveBeenCalledWith(
        expect.objectContaining({
          Type: "Decline",
          MeetingId: booking.MeetingId,
        }),
      );

      // Only when Respond fails does it fall back to Delete, and it must delete
      // by MeetingId so the hybrid calendar backend event is removed (deleting
      // by the local Id would only remove the local booking).
      expect(xapi.Command.Bookings.Delete).toHaveBeenCalledWith({
        MeetingId: booking.MeetingId,
      });
      expect(xapi.Command.Bookings.Delete).not.toHaveBeenCalledWith(
        expect.objectContaining({ Id: expect.anything() }),
      );
    });
  });


});
