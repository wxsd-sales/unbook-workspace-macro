import defaultConfig from "./example.json" with { type: "json" };

/**
 * Working configuration state. Seeded from example.json and mutated in place as
 * the user edits the form. buildConfig() derives the clean, exportable object.
 */
const config = {
  profiles: (defaultConfig.profiles || []).map(normalizeProfile),
  presenceDetection: {
    activeCalls: true,
    presentation: true,
    peopleCount: true,
    peoplePresence: true,
    guiInteractions: true,
    ...(defaultConfig.presenceDetection || {}),
  },
  externalLogging: {
    enabled: false,
    type: "bot",
    contact: "",
    url: "",
    token: "",
    ...(defaultConfig.externalLogging || {}),
  },
  debugging: Boolean(defaultConfig.debugging),
};

const PROFILE_TYPES = [
  { value: "duration", label: "Duration" },
  { value: "keywords", label: "Keywords" },
  { value: "organizers", label: "Organizers" },
  { value: "default", label: "Default" },
];

/** Fills in every possible field so each profile card renders consistently. */
function normalizeProfile(profile = {}) {
  return {
    type: profile.type || "duration",
    name: profile.name || "",
    duration: Array.isArray(profile.duration) ? profile.duration : [0, 60],
    keywords: Array.isArray(profile.keywords) ? profile.keywords : [],
    organizers: Array.isArray(profile.organizers) ? profile.organizers : [],
    monitor: profile.monitor !== false,
    startMonitoringDelay: numberOr(profile.startMonitoringDelay, 0),
    stopMonitoringAfter: numberOr(profile.stopMonitoringAfter, 10),
    requiredUnoccupiedDuration: numberOr(profile.requiredUnoccupiedDuration, 5),
    alertBeforeUnbookingDuration: numberOr(
      profile.alertBeforeUnbookingDuration,
      1,
    ),
  };
}

function numberOr(value, fallback) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function splitList(value) {
  return value
    .split(",")
    .map((item) => item.trim())
    .filter((item) => item.length > 0);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

/**
 * Builds the clean configuration object for export, dropping fields that do not
 * apply to a profile's type or that are irrelevant when monitoring is off.
 */
function buildConfig() {
  const profiles = config.profiles.map((profile) => {
    const output = { type: profile.type, name: profile.name };

    if (profile.type === "duration") {
      output.duration = [
        numberOr(profile.duration[0], 0),
        numberOr(profile.duration[1], 0),
      ];
    } else if (profile.type === "keywords") {
      output.keywords = [...profile.keywords];
    } else if (profile.type === "organizers") {
      output.organizers = [...profile.organizers];
    }

    output.monitor = profile.monitor;

    if (profile.monitor) {
      output.startMonitoringDelay = numberOr(profile.startMonitoringDelay, 0);
      output.stopMonitoringAfter = numberOr(profile.stopMonitoringAfter, 0);
      output.requiredUnoccupiedDuration = numberOr(
        profile.requiredUnoccupiedDuration,
        0,
      );
      output.alertBeforeUnbookingDuration = numberOr(
        profile.alertBeforeUnbookingDuration,
        0,
      );
    }

    return output;
  });

  return {
    profiles,
    presenceDetection: { ...config.presenceDetection },
    externalLogging: { ...config.externalLogging },
    debugging: config.debugging,
  };
}

/** Location of the macro relative to the webapp, and its output filename. */
const MACRO_URL = "../unbook-workspace.js";
const MACRO_FILENAME = "unbook-workspace.js";

/** Block-comment markers used to locate the config region inside the macro. */
const CONFIG_START_MARKER = "Configuration Start";
const CONFIG_END_MARKER = "Configuration End";

/**
 * Triggers a browser download of arbitrary text content.
 * @param {string} content - The file contents.
 * @param {string} filename - The download filename.
 * @param {string} mime - The MIME type for the blob.
 */
function downloadText(content, filename, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
}

/**
 * Replaces the config in the macro source with a new config block. The region
 * is located via the "Configuration Start" / "Configuration End" block comments
 * (rather than a regex) so it keeps working as the config schema changes: the
 * comment blocks are preserved and only the code between them is swapped.
 * @param {string} source - The full macro source.
 * @param {string} configBlock - The replacement `const config = { ... }` block.
 * @returns {string} The macro source with the new config applied.
 */
function injectConfig(source, configBlock) {
  const startMarker = source.indexOf(CONFIG_START_MARKER);
  const endMarker = source.indexOf(CONFIG_END_MARKER);
  if (startMarker === -1 || endMarker === -1) {
    throw new Error(
      "Could not find the Configuration Start / Configuration End markers in the macro.",
    );
  }

  // Keep everything up to and including the end of the "Configuration Start"
  // comment block, and everything from the start of the "Configuration End"
  // comment block onwards.
  const startBlockEnd = source.indexOf("*/", startMarker);
  const endBlockStart = source.lastIndexOf("/*", endMarker);
  if (
    startBlockEnd === -1 ||
    endBlockStart === -1 ||
    startBlockEnd >= endBlockStart
  ) {
    throw new Error("The configuration markers in the macro are malformed.");
  }

  const before = source.slice(0, startBlockEnd + 2);
  const after = source.slice(endBlockStart);
  return `${before}\n\n${configBlock}\n\n${after}`;
}

/**
 * Profiles tab: renders the dynamic list of profile cards and keeps the
 * `config.profiles` state in sync as the user edits, adds or removes profiles.
 */
(function initProfiles() {
  const profileList = document.getElementById("profileList");
  const addProfileBtn = document.getElementById("addProfileBtn");

  if (!profileList || !addProfileBtn) {
    return;
  }

  function typeOptions(selected, index) {
    // Only one profile may be the "Default". Disable the option for every other
    // profile once a default already exists elsewhere.
    const defaultTakenElsewhere = config.profiles.some(
      (profile, i) => i !== index && profile.type === "default",
    );

    return PROFILE_TYPES.map((type) => {
      const isSelected = type.value === selected;
      const disabled =
        type.value === "default" && defaultTakenElsewhere && !isSelected;
      return `<option value="${type.value}" ${isSelected ? "selected" : ""} ${
        disabled ? "disabled" : ""
      }>${type.label}</option>`;
    }).join("");
  }

  function profileCardHtml(profile, index) {
    return `
      <div class="profile-card" data-index="${index}">
        <div class="profile-card__header">
          <span class="profile-card__title">Profile ${index + 1}</span>
          <button
            type="button"
            class="icon-button icon-button--with-label secondary-button"
            data-action="remove-profile"
          >
            <span class="icon icon-delete-regular" aria-hidden="true"></span>
            <span class="icon-button__label">Remove</span>
          </button>
        </div>
        <div class="config-form">
          <div class="field">
            <label class="field__label">Profile Type</label>
            <div class="select-wrap">
              <select class="field__input field__select" data-field="type">
                ${typeOptions(profile.type, index)}
              </select>
            </div>
          </div>
          <div class="field">
            <label class="field__label">Profile Name</label>
            <input
              class="field__input"
              type="text"
              data-field="name"
              value="${escapeHtml(profile.name)}"
              placeholder="Descriptive name for logging"
            />
          </div>

          <div class="field" data-group="duration">
            <label class="field__label">Booking Duration (minutes)</label>
            <div class="field-row">
              <input
                class="field__input"
                type="number"
                min="0"
                data-field="durationMin"
                value="${escapeHtml(profile.duration[0])}"
                placeholder="Min"
                aria-label="Minimum duration in minutes"
              />
              <input
                class="field__input"
                type="number"
                min="0"
                data-field="durationMax"
                value="${escapeHtml(profile.duration[1])}"
                placeholder="Max"
                aria-label="Maximum duration in minutes"
              />
            </div>
            <span class="field__hint"
              >Match bookings whose length falls within this range.</span
            >
          </div>

          <div class="field" data-group="keywords">
            <label class="field__label">Title Keywords</label>
            <input
              class="field__input"
              type="text"
              data-field="keywords"
              value="${escapeHtml(profile.keywords.join(", "))}"
              placeholder="Training, Test"
            />
            <span class="field__hint"
              >Comma separated keywords to look for in the booking title.</span
            >
          </div>

          <div class="field" data-group="organizers">
            <label class="field__label">Organizers</label>
            <input
              class="field__input"
              type="text"
              data-field="organizers"
              value="${escapeHtml(profile.organizers.join(", "))}"
              placeholder="John Smith, Jane Doe"
            />
            <span class="field__hint"
              >Comma separated organizer names to match against bookings.</span
            >
          </div>

          <div class="toggle-row">
            <div class="toggle-row__text">
              <span class="toggle-row__label">Monitor matching bookings</span>
              <span class="toggle-row__hint"
                >When off, matching bookings are ignored and never
                unbooked.</span
              >
            </div>
            <label class="switch">
              <input
                type="checkbox"
                class="switch__input"
                data-field="monitor"
                ${profile.monitor ? "checked" : ""}
              />
              <span class="switch__track"
                ><span class="switch__thumb"></span
              ></span>
            </label>
          </div>

          <div data-group="monitoring">
            <div class="field-row">
              <div class="field">
                <label class="field__label">Start Monitoring Delay (min)</label>
                <input
                  class="field__input"
                  type="number"
                  min="0"
                  data-field="startMonitoringDelay"
                  value="${escapeHtml(profile.startMonitoringDelay)}"
                />
              </div>
              <div class="field">
                <label class="field__label">Stop Monitoring After (min)</label>
                <input
                  class="field__input"
                  type="number"
                  min="0"
                  data-field="stopMonitoringAfter"
                  value="${escapeHtml(profile.stopMonitoringAfter)}"
                />
              </div>
            </div>
            <div class="field-row">
              <div class="field">
                <label class="field__label"
                  >Required Unoccupied Duration (min)</label
                >
                <input
                  class="field__input"
                  type="number"
                  min="0"
                  data-field="requiredUnoccupiedDuration"
                  value="${escapeHtml(profile.requiredUnoccupiedDuration)}"
                />
              </div>
              <div class="field">
                <label class="field__label">Alert Before Unbooking (min)</label>
                <input
                  class="field__input"
                  type="number"
                  min="0"
                  data-field="alertBeforeUnbookingDuration"
                  value="${escapeHtml(profile.alertBeforeUnbookingDuration)}"
                />
              </div>
            </div>
          </div>
        </div>
      </div>`;
  }

  function syncCardVisibility(card, profile) {
    card.querySelectorAll("[data-group]").forEach((group) => {
      const name = group.dataset.group;
      let visible = true;
      if (name === "duration") {
        visible = profile.type === "duration";
      } else if (name === "keywords") {
        visible = profile.type === "keywords";
      } else if (name === "organizers") {
        visible = profile.type === "organizers";
      } else if (name === "monitoring") {
        visible = profile.monitor;
      }
      group.hidden = !visible;
    });
  }

  function render() {
    profileList.innerHTML = config.profiles
      .map((profile, index) => profileCardHtml(profile, index))
      .join("");

    profileList.querySelectorAll(".profile-card").forEach((card) => {
      const index = Number(card.dataset.index);
      syncCardVisibility(card, config.profiles[index]);
    });
  }

  function handleFieldChange(event) {
    const input = event.target.closest("[data-field]");
    if (!input) {
      return;
    }
    const card = input.closest(".profile-card");
    const index = Number(card.dataset.index);
    const profile = config.profiles[index];
    if (!profile) {
      return;
    }

    switch (input.dataset.field) {
      case "type":
        profile.type = input.value;
        // Re-render every card so the "Default" option enables/disables
        // consistently across all profiles.
        render();
        updatePreview();
        return;
      case "name":
        profile.name = input.value;
        break;
      case "durationMin":
        profile.duration[0] = numberOr(input.value, 0);
        break;
      case "durationMax":
        profile.duration[1] = numberOr(input.value, 0);
        break;
      case "keywords":
        profile.keywords = splitList(input.value);
        break;
      case "organizers":
        profile.organizers = splitList(input.value);
        break;
      case "monitor":
        profile.monitor = input.checked;
        syncCardVisibility(card, profile);
        break;
      default:
        profile[input.dataset.field] = numberOr(input.value, 0);
        break;
    }

    updatePreview();
  }

  profileList.addEventListener("input", handleFieldChange);
  profileList.addEventListener("change", handleFieldChange);

  profileList.addEventListener("click", (event) => {
    const removeBtn = event.target.closest("[data-action='remove-profile']");
    if (!removeBtn) {
      return;
    }
    const index = Number(removeBtn.closest(".profile-card").dataset.index);
    config.profiles.splice(index, 1);
    render();
    updatePreview();
  });

  addProfileBtn.addEventListener("click", () => {
    config.profiles.push(normalizeProfile({ name: "New Profile" }));
    render();
    updatePreview();
    const cards = profileList.querySelectorAll(".profile-card");
    cards[cards.length - 1]?.scrollIntoView({
      behavior: "smooth",
      block: "nearest",
    });
  });

  render();
})();

/**
 * Sensors tab: binds each presence-detection toggle to config.presenceDetection.
 */
(function initPresenceDetection() {
  const toggles = document.querySelectorAll("[data-presence]");
  if (!toggles.length) {
    return;
  }

  toggles.forEach((toggle) => {
    const key = toggle.dataset.presence;
    toggle.checked = Boolean(config.presenceDetection[key]);
    toggle.addEventListener("change", () => {
      config.presenceDetection[key] = toggle.checked;
      updatePreview();
    });
  });
})();

/**
 * Logging tab: binds the external logging fields and the global debugging flag.
 */
(function initLogging() {
  const enabled = document.getElementById("loggingEnabled");
  const fields = document.getElementById("loggingFields");
  const typeSelect = document.getElementById("loggingType");
  const contact = document.getElementById("loggingContact");
  const contactGroup = document.querySelector("[data-logging-group='contact']");
  const url = document.getElementById("loggingUrl");
  const urlGroup = document.querySelector("[data-logging-group='url']");
  const token = document.getElementById("loggingToken");
  const debugging = document.getElementById("debugging");

  // Webex Bot messages always post to this endpoint, so the URL is fixed and
  // hidden for the bot service.
  const WEBEX_MESSAGES_URL = "https://webexapis.com/v1/messages";

  if (
    !enabled ||
    !fields ||
    !typeSelect ||
    !contact ||
    !url ||
    !token ||
    !debugging
  ) {
    return;
  }

  const logging = config.externalLogging;

  // Remember the user's own webhook URL so switching bot -> webhook restores it.
  let lastWebhookUrl = logging.type === "webhook" ? logging.url : "";

  const syncEnabledState = () => {
    fields.hidden = !logging.enabled;
  };

  const syncTypeState = () => {
    const isBot = logging.type === "bot";

    // The contact / room ID is only used by the Webex Bot service.
    if (contactGroup) {
      contactGroup.hidden = !isBot;
    }

    // The Webex Bot service uses a fixed endpoint, so hide the URL field and
    // force the messages URL. Webhook mode shows the field and restores the
    // last custom URL.
    if (urlGroup) {
      urlGroup.hidden = isBot;
    }

    if (isBot) {
      logging.url = WEBEX_MESSAGES_URL;
      url.value = WEBEX_MESSAGES_URL;
    } else {
      logging.url = lastWebhookUrl;
      url.value = lastWebhookUrl;
    }
  };

  enabled.checked = logging.enabled;
  typeSelect.value = logging.type;
  contact.value = logging.contact;
  url.value = logging.url;
  token.value = logging.token;
  debugging.checked = config.debugging;

  syncEnabledState();
  syncTypeState();

  enabled.addEventListener("change", () => {
    logging.enabled = enabled.checked;
    syncEnabledState();
    updatePreview();
  });

  typeSelect.addEventListener("change", () => {
    logging.type = typeSelect.value;
    syncTypeState();
    updatePreview();
  });

  contact.addEventListener("input", () => {
    logging.contact = contact.value;
    updatePreview();
  });

  url.addEventListener("input", () => {
    logging.url = url.value;
    // Only track manual edits as the webhook URL; bot mode hides this field.
    if (logging.type !== "bot") {
      lastWebhookUrl = url.value;
    }
    updatePreview();
  });

  token.addEventListener("input", () => {
    logging.token = token.value;
    updatePreview();
  });

  debugging.addEventListener("change", () => {
    config.debugging = debugging.checked;
    updatePreview();
  });
})();

/**
 * Export tab: keeps a live preview of the generated config and wires the copy
 * and download actions.
 */
const updatePreview = (function initExport() {
  const preview = document.getElementById("configPreview");
  const copyBtn = document.getElementById("copyConfigBtn");
  const downloadBtn = document.getElementById("downloadMacroBtn");
  const status = document.getElementById("exportStatus");

  const snippet = () => `const config = ${JSON.stringify(buildConfig(), null, 2)}`;

  const update = () => {
    if (preview) {
      preview.textContent = snippet();
    }
  };

  const setStatus = (message, kind) => {
    if (!status) {
      return;
    }
    status.textContent = message || "";
    status.hidden = !message;
    status.dataset.kind = kind || "";
  };

  if (copyBtn && preview) {
    copyBtn.addEventListener("click", async () => {
      try {
        await navigator.clipboard.writeText(preview.textContent);
      } catch (error) {
        return;
      }

      const label = copyBtn.querySelector(".icon-button__label");
      const icon = copyBtn.querySelector(".icon");
      const previousLabel = label.textContent;

      label.textContent = "Copied";
      icon.classList.remove("icon-copy-bold");
      icon.classList.add("icon-check-circle-bold");

      window.setTimeout(() => {
        label.textContent = previousLabel;
        icon.classList.remove("icon-check-circle-bold");
        icon.classList.add("icon-copy-bold");
      }, 1600);
    });
  }

  if (downloadBtn) {
    downloadBtn.addEventListener("click", async () => {
      setStatus("Building macro\u2026", "pending");
      downloadBtn.disabled = true;
      try {
        const response = await fetch(MACRO_URL, { cache: "no-store" });
        if (!response.ok) {
          throw new Error(`Failed to fetch the macro (HTTP ${response.status}).`);
        }
        const source = await response.text();
        const macro = injectConfig(source, snippet());
        downloadText(macro, MACRO_FILENAME, "text/javascript");
        setStatus(
          `Downloaded ${MACRO_FILENAME} with your configuration applied.`,
          "success",
        );
      } catch (error) {
        setStatus(
          `${error.message} You can still copy the config block above and paste it into the macro.`,
          "error",
        );
      } finally {
        downloadBtn.disabled = false;
      }
    });
  }

  update();
  return update;
})();

/**
 * Theme selector: toggles the menu and applies System / Light / Dark themes.
 * Light/Dark are persisted via the URL hash (read by the inline boot script),
 * while System clears the hash and follows the OS preference.
 */
(function initThemeSelect() {
  const root = document.documentElement;
  const select = document.getElementById("theme-select");
  const button = document.getElementById("theme-select-button");
  const menu = document.getElementById("theme-select-menu");
  const label = document.getElementById("theme-select-label");
  const currentIcon = document.getElementById("theme-select-current-icon");

  if (!select || !button || !menu || !label || !currentIcon) {
    return;
  }

  const options = Array.from(menu.querySelectorAll(".theme-select-option"));

  const META = {
    system: { label: "System", icon: "icon-laptop-regular" },
    light: { label: "Light", icon: "icon-brightness-high-filled" },
    dark: { label: "Dark", icon: "icon-quiet-hours-presence-filled" },
  };
  const ICON_CLASSES = Object.values(META).map((meta) => meta.icon);

  const readChoice = () => {
    const raw = window.location.hash.startsWith("#")
      ? window.location.hash.slice(1)
      : window.location.hash;
    const theme = raw ? new URLSearchParams(raw).get("theme") : null;
    return theme === "light" || theme === "dark" ? theme : "system";
  };

  const applyTheme = (choice) => {
    const dark =
      choice === "dark" ||
      (choice === "system" &&
        window.matchMedia("(prefers-color-scheme: dark)").matches);

    root.classList.remove(
      "mds-theme-stable-lightWebex",
      "mds-theme-stable-darkWebex",
    );
    root.classList.add(
      dark ? "mds-theme-stable-darkWebex" : "mds-theme-stable-lightWebex",
    );
    root.style.colorScheme = dark ? "dark" : "light";
  };

  const syncButton = (choice) => {
    const meta = META[choice] || META.system;
    label.textContent = meta.label;
    currentIcon.classList.remove(...ICON_CLASSES);
    currentIcon.classList.add(meta.icon);
    options.forEach((option) => {
      option.setAttribute(
        "aria-selected",
        String(option.dataset.themeChoice === choice),
      );
    });
  };

  const setChoice = (choice) => {
    if (choice === "system") {
      history.replaceState(
        null,
        "",
        window.location.pathname + window.location.search,
      );
    } else {
      window.location.hash = "theme=" + choice;
    }
    applyTheme(choice);
    syncButton(choice);
  };

  const openMenu = () => {
    menu.hidden = false;
    select.dataset.open = "true";
    button.setAttribute("aria-expanded", "true");
  };

  const closeMenu = () => {
    menu.hidden = true;
    select.dataset.open = "false";
    button.setAttribute("aria-expanded", "false");
  };

  button.addEventListener("click", (event) => {
    event.stopPropagation();
    if (menu.hidden) {
      openMenu();
    } else {
      closeMenu();
    }
  });

  options.forEach((option) => {
    option.addEventListener("click", () => {
      setChoice(option.dataset.themeChoice);
      closeMenu();
      button.focus();
    });
  });

  document.addEventListener("click", (event) => {
    if (!select.contains(event.target)) {
      closeMenu();
    }
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !menu.hidden) {
      closeMenu();
      button.focus();
    }
  });

  syncButton(readChoice());
})();

/**
 * Tab list: toggles which configuration panel is visible.
 */
(function initTabs() {
  const tabs = Array.from(document.querySelectorAll(".tab"));
  if (!tabs.length) {
    return;
  }

  const activate = (tab) => {
    tabs.forEach((current) => {
      const selected = current === tab;
      current.setAttribute("aria-selected", String(selected));
      current.tabIndex = selected ? 0 : -1;
      const panel = document.getElementById(current.dataset.tabTarget);
      if (panel) {
        panel.hidden = !selected;
      }
    });
  };

  tabs.forEach((tab, index) => {
    tab.addEventListener("click", () => activate(tab));
    tab.addEventListener("keydown", (event) => {
      if (event.key !== "ArrowRight" && event.key !== "ArrowLeft") {
        return;
      }
      event.preventDefault();
      const direction = event.key === "ArrowRight" ? 1 : -1;
      const next = tabs[(index + direction + tabs.length) % tabs.length];
      next.focus();
      activate(next);
    });
  });
})();
