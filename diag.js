/* Rein lesende Kontextdiagnose fuer das Outlook-Add-in "Betreff" (Schritt 0).
   Liest ausschliesslich Office.context.mailbox.item / userProfile. Schreibt nichts:
   keine item.subject.setAsync, keine Notification-Messages, kein Graph, keine Netzwerkzugriffe. */
(function (root) {
  "use strict";

  // Reine Funktion; Office wird injiziert, damit sie ohne Outlook mit einem Mock testbar ist.
  function collectDiagnosis(Office) {
    var out = { zeit: new Date().toISOString(), fehler: [] };
    var mb = Office && Office.context && Office.context.mailbox;
    if (!mb) { out.fehler.push("Office.context.mailbox fehlt"); return Promise.resolve(out); }

    try { out.host = Office.context.diagnostics ? Office.context.diagnostics.host + " / " + Office.context.diagnostics.platform + " / " + Office.context.diagnostics.version : "n/a"; } catch (e) { out.host = "n/a"; }
    try { out.postfach = mb.userProfile && mb.userProfile.emailAddress; } catch (e) { out.postfach = null; }

    var item = mb.item;
    if (!item) { out.fehler.push("mailbox.item ist leer (keine Mail ausgewaehlt?)"); return Promise.resolve(out); }

    out.itemType = item.itemType;
    out.itemClass = item.itemClass || null;
    out.conversationId = item.conversationId || null;

    // Compose-Modus: subject ist ein Objekt (Subject-Klasse) statt String -> nicht der Zielfall.
    if (item.subject && typeof item.subject === "object") {
      out.modus = "COMPOSE (Entwurf) - nicht Zielfall dieses Add-ins";
      out.subject = null;
    } else {
      out.modus = "READ";
      out.subject = item.subject;
    }

    out.itemId = item.itemId || null; // undefined bei ungespeichertem Entwurf
    out.restId = null;
    if (out.itemId && typeof mb.convertToRestId === "function") {
      try {
        out.restId = mb.convertToRestId(out.itemId, Office.MailboxEnums.RestVersion.v2_0);
      } catch (e) { out.fehler.push("convertToRestId: " + e.message); }
    } else if (!out.itemId) {
      out.fehler.push("itemId fehlt");
    }
    out.internetMessageId = item.internetMessageId || null; // zweiter, versionsunabhaengiger Anker
    out.dateTimeCreated = item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : null;
    out.ok = !!(out.modus === "READ" && out.subject != null && out.itemId && out.restId);
    return Promise.resolve(out);
  }

  root.BetreffDiag = { collectDiagnosis: collectDiagnosis };
})(typeof window !== "undefined" ? window : this);
