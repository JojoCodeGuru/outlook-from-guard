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

// Resolves to the From address, or "" if it can't be read (never rejects).
function getFromAddress() {
  return new Promise((resolve) => {
    try {
      Office.context.mailbox.item.from.getAsync((result) => {
        const ok = result.status === Office.AsyncResultStatus.Succeeded && result.value;
        resolve(ok ? result.value.emailAddress || "" : "");
      });
    } catch (e) {
      resolve("");
    }
  });
}

function showNotification(message) {
  return new Promise((resolve) => {
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      NOTIFICATION_KEY,
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: message.slice(0, 150),
        icon: "Icon.16x16",
        persistent: false,
      },
      () => resolve()
    );
  });
}

function clearNotification() {
  return new Promise((resolve) => {
    Office.context.mailbox.item.notificationMessages.removeAsync(NOTIFICATION_KEY, () => resolve());
  });
}

function warningText(address) {
  return `You're in ${CONFIG.locationName} – switch From away from ${domainOf(address)} unless you mean it.`;
}

async function refreshNotification() {
  const address = await getFromAddress();
  if (shouldWarn(address, currentTimeZone())) {
    await showNotification(warningText(address));
  } else {
    await clearNotification();
  }
}

// OnNewMessageCompose (new mail, reply, reply all, forward) and OnMessageFromChanged.
function onComposeHandler(event) {
  refreshNotification().then(
    () => event.completed(),
    () => event.completed()
  );
}

// OnMessageSend – Smart Alert. SendMode="PromptUser" in the manifest gives the
// dialog a "Send Anyway" button, which is the manual override.
function onMessageSendHandler(event) {
  getFromAddress().then((address) => {
    if (shouldWarn(address, currentTimeZone())) {
      event.completed({
        allowEvent: false,
        errorMessage:
          `You're in ${CONFIG.locationName} and this message is being sent from ${address}. ` +
          `Choose "Don't Send" and change the From address, or "Send Anyway" if you meant to use it.`,
      });
    } else {
      event.completed({ allowEvent: true });
    }
  });
}

// Ribbon button: shows what the add-in currently sees, handy for checking the setup.
function showStatus(event) {
  getFromAddress().then((address) => {
    const tz = currentTimeZone() || "unknown";
    const inLocation = CONFIG.locationTimeZones.includes(tz);
    const verdict = shouldWarn(address, tz) ? "would warn" : "OK to send";
    return showNotification(
      `From Guard: time zone ${tz} (${inLocation ? "in" : "not in"} ${CONFIG.locationName}), From ${address || "unknown"} – ${verdict}.`
    );
  }).then(
    () => event.completed(),
    () => event.completed()
  );
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onComposeHandler", onComposeHandler);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("showStatus", showStatus);
}
