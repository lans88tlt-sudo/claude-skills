async function unmuteAll() {
  const tabs = await chrome.tabs.query({});
  for (const t of tabs) {
    if (t.mutedInfo && t.mutedInfo.muted) await chrome.tabs.update(t.id, { muted: false });
  }
}
chrome.runtime.onInstalled.addListener(unmuteAll);
chrome.runtime.onStartup.addListener(unmuteAll);
chrome.action.onClicked.addListener(unmuteAll);
chrome.commands.onCommand.addListener(c => { if (c === 'unmute-all') unmuteAll(); });
