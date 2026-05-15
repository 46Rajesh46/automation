// FIX 2: Edge/Chrome Manifest V3 extension — auto-accepts flagged downloads.
// Uses the native chrome.downloads API which is more reliable than Selenium
// shadow-DOM clicking or pywinauto UI automation.
//
// HOW TO INSTALL (one-time):
//   1. Open Edge and go to: edge://extensions/
//   2. Turn on "Developer mode" (toggle top-right)
//   3. Click "Load unpacked"
//   4. Select this folder (fix2_edge_extension)
//   5. Extension is now active — leave it installed permanently.
//
// WHAT IT DOES:
//   Whenever a download changes to a "dangerous" state (Edge blocks it with
//   "Keep" / "Keep anyway" prompt), this script calls acceptDanger() which
//   is the same action as the user clicking "Keep anyway" — no UI needed.

chrome.downloads.onChanged.addListener((delta) => {
    // Only act when the danger field changes to a non-safe, non-accepted value
    if (!delta.danger) return;

    const dangerState = delta.danger.current;
    const safeStates = ["safe", "accepted", "allowlistedByPolicy"];
    if (safeStates.includes(dangerState)) return;

    // Accept the dangerous download
    chrome.downloads.acceptDanger(delta.id, () => {
        if (chrome.runtime.lastError) {
            console.warn("[AutoAccept] acceptDanger failed:", chrome.runtime.lastError.message);
        } else {
            console.log(`[AutoAccept] Accepted download id=${delta.id} danger=${dangerState}`);
        }
    });
});
