/*
 * From Guard – warns when a message is about to go out from a blocked domain
 * while you are in a given location.
 *
 * Location is read from the Mac's time zone (Intl API). With macOS set to
 * "Set time zone automatically using your current location", the time zone is
 * Europe/Copenhagen whenever you are in Denmark.
 */

const CONFIG = {
  locationName: "Denmark",
  locationTimeZones: ["Europe/Copenhagen"],
  blockedDomains: ["aris-labs.com"],
};

const NOTIFICATION_KEY = "fromGuardWarning";

function currentTimeZone() {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || "";
  } catch (e) {
    return "";
  }
}

function domainOf(address) {
  const at = (address || "").lastIndexOf("@");
  return at < 0 ? "" : address.slice(at + 1).toLowerCase();
}

function isBlockedAddress(address) {
  const domain = domainOf(address);
  return CONFIG.blockedDomains.some((b) => domain === b || domain.endsWith("." + b));
}

function shouldWarn(address, timeZone) {
  return CONFIG.locationTimeZones.includes(timeZone) && isBlockedAddress(address);
}

// How long to wait for any single Outlook call, and for a whole check, before
// giving up. Outlook shows "taking longer than expected" after about 5 seconds.
const CALL_TIMEOUT_MS = 2000;
const SEND_TIMEOUT_MS = 3000;
const COMPOSE_TIMEOUT_MS = 4000;

const TRACE_KEY = "fromGuardTrace";
const STATUS_KEY = "fromGuardStatus";

// Keeps the last few steps so "Check From" can show what the last check did.
function trace(step) {
  try {
    const log = JSON.parse(localStorage.getItem(TRACE_KEY) || "[]");
    log.push(new Date().toTimeString().slice(0, 8) + " " + step);
    localStorage.setItem(TRACE_KEY, JSON.stringify(log.slice(-12)));
  } catch (e) {
    // Storage unavailable – tracing is optional.
  }
}

function readTrace() {
  try {
    return JSON.parse(localStorage.getItem(TRACE_KEY) || "[]");
  } catch (e) {
    return [];
  }
}

// Resolves to `fallback` if `promise` hasn't settled within `ms`.
function withTimeout(promise, ms, fallback) {
  return Promise.race([promise, new Promise((resolve) => setTimeout(() => resolve(fallback), ms))]);
}

// event.completed may only be called once; the real answer and the fallback race.
function completeOnce(event) {
  let done = false;
  return (options) => {
    if (done) return;
    done = true;
    if (options) {
      event.completed(options);
    } else {
      event.completed();
    }
  };
}

// Resolves to the From address, "" if it can't be read, or null if Outlook
// never answered (never rejects).
function getFromAddress() {
  const request = new Promise((resolve) => {
    try {
      Office.context.mailbox.item.from.getAsync((result) => {
        const ok = result.status === Office.AsyncResultStatus.Succeeded && result.value;
        resolve(ok ? result.value.emailAddress || "" : "");
      });
    } catch (e) {
      resolve("");
    }
  });
  return withTimeout(request, CALL_TIMEOUT_MS, null);
}

function describeFrom(address) {
  if (address === null) return "from=timeout";
  return address ? "from=" + domainOf(address) : "from=unreadable";
}

function showNotification(key, message) {
  const request = new Promise((resolve) => {
    try {
      Office.context.mailbox.item.notificationMessages.replaceAsync(
        key,
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: message.slice(0, 150),
          icon: "Icon.16x16",
          persistent: false,
        },
        () => resolve()
      );
    } catch (e) {
      resolve();
    }
  });
  return withTimeout(request, CALL_TIMEOUT_MS);
}

function clearNotification(key) {
  const request = new Promise((resolve) => {
    try {
      Office.context.mailbox.item.notificationMessages.removeAsync(key, () => resolve());
    } catch (e) {
      resolve();
    }
  });
  return withTimeout(request, CALL_TIMEOUT_MS);
}

function warningText(address) {
  return `You're in ${CONFIG.locationName} – switch From away from ${domainOf(address)} unless you mean it.`;
}

// OnNewMessageCompose (new mail, reply, reply all, forward) and OnMessageFromChanged.
function onComposeHandler(event) {
  const done = completeOnce(event);
  trace("compose: start");
  setTimeout(() => {
    trace("compose: gave up waiting");
    done();
  }, COMPOSE_TIMEOUT_MS);

  getFromAddress()
    .then((address) => {
      trace("compose: " + describeFrom(address));
      return shouldWarn(address, currentTimeZone())
        ? showNotification(NOTIFICATION_KEY, warningText(address))
        : clearNotification(NOTIFICATION_KEY);
    })
    .then(
      () => {
        trace("compose: done");
        done();
      },
      () => {
        trace("compose: error");
        done();
      }
    );
}

function uncheckedFromMessage() {
  return (
    `You're in ${CONFIG.locationName} and From Guard couldn't confirm which address this message is sent from. ` +
    `Check the From field: choose "Don't Send" to change it, or "Send Anyway" if it's right.`
  );
}

// OnMessageSend – Smart Alert. SendMode="PromptUser" in the manifest gives the
// dialog a "Send Anyway" button, which is the manual override. Outside the
// location the message goes straight through; inside it, if the From address
// can't be confirmed in time, the pop-up is shown as a backstop.
function onMessageSendHandler(event) {
  const done = completeOnce(event);
  trace("send: start");

  if (!CONFIG.locationTimeZones.includes(currentTimeZone())) {
    trace("send: not in " + CONFIG.locationName + ", allowed");
    done({ allowEvent: true });
    return;
  }

  setTimeout(() => {
    trace("send: gave up waiting, pop-up");
    done({ allowEvent: false, errorMessage: uncheckedFromMessage() });
  }, SEND_TIMEOUT_MS);

  getFromAddress().then(
    (address) => {
      trace("send: " + describeFrom(address));
      if (!address) {
        trace("send: unconfirmed, pop-up");
        done({ allowEvent: false, errorMessage: uncheckedFromMessage() });
      } else if (isBlockedAddress(address)) {
        trace("send: warned");
        done({
          allowEvent: false,
          errorMessage:
            `You're in ${CONFIG.locationName} and this message is being sent from ${address}. ` +
            `Choose "Don't Send" and change the From address, or "Send Anyway" if you meant to use it.`,
        });
      } else {
        trace("send: allowed");
        done({ allowEvent: true });
      }
    },
    () => {
      trace("send: error, pop-up");
      done({ allowEvent: false, errorMessage: uncheckedFromMessage() });
    }
  );
}

// Ribbon button: shows what the add-in currently sees, plus the last steps it
// recorded, handy for checking the setup.
function showStatus(event) {
  const done = completeOnce(event);
  setTimeout(() => done(), COMPOSE_TIMEOUT_MS);

  getFromAddress()
    .then((address) => {
      const tz = currentTimeZone() || "unknown";
      const inLocation = CONFIG.locationTimeZones.includes(tz);
      const verdict = shouldWarn(address, tz) ? "would warn" : "OK to send";
      const from = address === null ? "not answered" : address || "unknown";
      const lastSteps = readTrace().slice(-4).join(" | ") || "nothing recorded yet";
      return Promise.all([
        showNotification(
          STATUS_KEY,
          `From Guard: ${tz} (${inLocation ? "in" : "not in"} ${CONFIG.locationName}), From ${from} – ${verdict}.`
        ),
        showNotification(TRACE_KEY, "Last steps: " + lastSteps),
      ]);
    })
    .then(
      () => done(),
      () => done()
    );
}

function registerHandlers() {
  if (typeof Office === "undefined" || !Office.actions) return;
  Office.actions.associate("onComposeHandler", onComposeHandler);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("showStatus", showStatus);
}

// Register straight away, and again once Office.js reports it's ready, in case
// the first attempt ran before Outlook was listening.
registerHandlers();
if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(registerHandlers);
}
