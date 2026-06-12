/*
 * DefaultFrom - Event Handler
 * Sideloading-Variante: alle 4 Events aktiv
 * - OnNewMessageCompose
 * - OnMessageReply
 * - OnMessageReplyAll
 * - OnMessageForward
 *
 * Terminal-Server-sicher: Settings in Exchange roamingSettings (pro User, session-unabhaengig)
 */

Office.onReady(() => {});

const STORAGE_KEY = "defaultfrom_settings";

function getSettings() {
  try {
    const raw = Office.context.roamingSettings.get(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return null;
}

function isUserAllowed(settings) {
  if (!settings.allowedUsers || settings.allowedUsers.length === 0) return true;
  try {
    const upn = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
    return settings.allowedUsers.map(u => u.toLowerCase()).includes(upn);
  } catch (e) {
    return false;
  }
}

function applyFrom(event) {
  try {
    const settings = getSettings();

    if (!settings || !settings.sharedEmail) {
      if (event) event.completed();
      return;
    }

    if (!isUserAllowed(settings)) {
      if (event) event.completed();
      return;
    }

    Office.context.mailbox.item.from.setAsync({
      emailAddress: {
        address: settings.sharedEmail,
        name: settings.sharedDisplay || settings.sharedEmail
      }
    }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("DefaultFrom setAsync failed:", result.error.message);
      }
      if (event) event.completed();
    });

  } catch (err) {
    console.error("DefaultFrom error:", err);
    if (event) event.completed();
  }
}

function onNewMessageCompose(event) { applyFrom(event); }
function onMessageReply(event)      { applyFrom(event); }
function onMessageReplyAll(event)   { applyFrom(event); }
function onMessageForward(event)    { applyFrom(event); }
