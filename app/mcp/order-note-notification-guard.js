/** Interim fail-closed guard for any MCP edit routed through the legacy note synchronizer. */
/** @param {string[]} currentIds @param {number} unreadCount @param {string[]} nextIds @param {string[]|null} recipientIds null means ordinary non-messaging note edit. */
export function orderNoteNotificationError(currentIds, unreadCount, nextIds, recipientIds) {
  if (currentIds.length) return "This note already contains staff mentions. Use an existing thread reply or wait for dedicated new-thread support; no note was changed.";
  if (unreadCount) return "This note has existing unread messages. A note edit could replace them; no note was changed.";
  const actual = new Set(nextIds);
  if (recipientIds === null) return actual.size ? "This note would notify staff. Use preview_send_staff_message to select and approve recipients explicitly." : null;
  if (!recipientIds.length || actual.size !== recipientIds.length || recipientIds.some((id) => !actual.has(id))) return "The notification recipients do not match the preview. Preview again.";
  return null;
}
