/*
 * SecondBrain Add-in POC
 *
 * Purpose: Test whether body.setAsync() can inject draft text into
 *          Outlook's compose window when the user clicks Reply.
 *
 * This is a proof-of-concept — no backend calls, no AI.
 * Just proves the injection mechanism works.
 */

// Initialize Office.js runtime
Office.onReady(function (info) {
  console.log("[SecondBrain] Office.js ready. Host:", info.host, "Platform:", info.platform);
});

/**
 * Event handler: fires when ANY compose window opens.
 * We check if it's a Reply (not a new email) before injecting.
 */
function onComposeHandler(event) {
  var item = Office.context.mailbox.item;

  console.log("[SecondBrain] Compose window opened");
  console.log("[SecondBrain] Subject:", item.subject);
  console.log("[SecondBrain] Item type:", item.itemType);

  // Get the subject to determine if this is a reply
  item.subject.getAsync(function (subjectResult) {
    if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.log("[SecondBrain] Could not read subject:", subjectResult.error.message);
      event.completed();
      return;
    }

    var subject = subjectResult.value || "";
    console.log("[SecondBrain] Subject value:", subject);

    // Check if this is a reply (subject starts with RE: or FW:)
    var isReply = /^(RE|FW|Fwd):/i.test(subject.trim());

    if (!isReply) {
      console.log("[SecondBrain] New compose (not a reply) — skipping injection");
      event.completed();
      return;
    }

    console.log("[SecondBrain] Reply detected! Attempting draft injection...");

    // Get the conversation ID (useful for future backend lookups)
    var conversationId = item.conversationId;
    console.log("[SecondBrain] Conversation ID:", conversationId);

    // Build the test draft content
    var draftHtml =
      '<div style="font-family: Calibri, sans-serif; font-size: 11pt;">' +
        '<p>Hi [Client],</p>' +
        '<p>Thank you for your email regarding <strong>' + escapeHtml(subject.replace(/^RE:\s*/i, '')) + '</strong>.</p>' +
        '<p>This is a <span style="background-color: #FFFF00; padding: 2px 6px;">' +
        'SecondBrain POC test draft</span> — if you can see this text in your ' +
        'compose window, the injection mechanism works.</p>' +
        '<p><em>Test timestamp: ' + new Date().toISOString() + '</em></p>' +
        '<p>Kind regards,<br>[Broker Name]</p>' +
      '</div>';

    // THE CRITICAL TEST: body.setAsync()
    item.body.setAsync(
      draftHtml,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("[SecondBrain] Draft injected successfully!");
          console.log("[SecondBrain] Coercion type: Html");
          console.log("[SecondBrain] Conversation ID:", conversationId);
        } else {
          console.error("[SecondBrain] Injection FAILED:", result.error.message);
          console.error("[SecondBrain] Error code:", result.error.code);

          // Fallback: try plain text if HTML fails
          console.log("[SecondBrain] Trying plain text fallback...");
          var draftText =
            "Hi [Client],\n\n" +
            "Thank you for your email regarding " + subject.replace(/^RE:\s*/i, '') + ".\n\n" +
            "[SecondBrain POC test draft] — if you can see this text, the injection works.\n\n" +
            "Test timestamp: " + new Date().toISOString() + "\n\n" +
            "Kind regards,\n[Broker Name]";

          item.body.setAsync(
            draftText,
            { coercionType: Office.CoercionType.Text },
            function (textResult) {
              if (textResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("[SecondBrain] Plain text fallback succeeded!");
              } else {
                console.error("[SecondBrain] Plain text fallback also FAILED:", textResult.error.message);
              }
              event.completed();
            }
          );
          return;
        }

        event.completed();
      }
    );
  });
}

/**
 * Escape HTML special characters to prevent XSS in injected content.
 */
function escapeHtml(text) {
  var div = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return text.replace(/[&<>"']/g, function (char) {
    return div[char];
  });
}

// REQUIRED: Map the manifest event handler name to the JS function.
// Classic Outlook on Windows requires Office.actions.associate().
Office.actions.associate("onComposeHandler", onComposeHandler);
