const CONFIG = {
  SEED_BILLING_FROM: "seed",
  BKBL_PATTERN: /\b(BK|BL)[\s\-]?([A-Z0-9\-]{4,20})\b/i,
  SEARCH_DAYS_BACK: 365,
  DEFAULT_MESSAGE: `Estimado/a cliente,\n\nAdjunto encontrará la factura electrónica y el archivo XML correspondiente a su embarque.\n\nPor favor, revise los documentos adjuntos. Si tiene alguna consulta o requiere información adicional, no dude en contactarnos.\n\nQuedamos a sus órdenes.\n\nAtentamente,\nDpto. de Facturación\nTransoceanica C.A.\ncustomerser@transoceanica.com.ec`
};

let state = { bkbl:null, clientEmail:null, clientName:null, clientSubject:null, foundItemId:null, foundSubject:null, foundDate:null, accessToken:null };

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) { showAlert("Solo funciona en Outlook.", "error"); return; }
  await initAddin();
});

async function initAddin() {
  try {
    setStep(1, "active");
    const item = Office.context.mailbox.item;
    const subject = item.subject;
    const from = item.from;
    state.clientSubject = subject;
    state.clientEmail = from?.emailAddress || null;
    state.clientName = from?.displayName || "Cliente";
    document.getElementById("subjectDisplay").textContent = subject || "(sin asunto)";
    const match = subject ? subject.match(CONFIG.BKBL_PATTERN) : null;
    if (match) {
      state.bkbl = (match[1] + match[2]).toUpperCase();
      displayBkbl(state.bkbl);
      setStep(1, "done", "Detectado");
      await searchBillingEmail();
    } else {
      displayBkbl(null);
      setStep(1, "error", "No encontrado");
      showAlert("No se detectó BK/BL en el asunto. Ingrésalo manualmente y presiona <strong>Usar</strong>.", "info");
      document.getElementById("bkblManual").focus();
    }
  } catch (err) {
    showAlert("Error al leer el correo: " + err.message, "error");
  }
}

function displayBkbl(value) {
  const el = document.getElementById("bkblDisplay");
  el.innerHTML = value ? `<strong>${escHtml(value)}</strong>` : `<span class="empty">No detectado automáticamente</span>`;
}

async function applyManualBkbl() {
  const val = document.getElementById("bkblManual").value.trim().toUpperCase();
  if (!val) { showAlert("Por favor ingresa un BK/BL válido.", "error"); return; }
  state.bkbl = val;
  displayBkbl(val);
  setStep(1, "done", "Manual");
  hideAlert();
  await searchBillingEmail();
}

async function searchBillingEmail() {
  setStep(2, "active");
  try {
    const token = await getAccessToken();
    if (!token) { await searchWithOfficeJs(); return; }
    state.accessToken = token;
    await searchWithGraph(token);
  } catch (err) {
    setStep(2, "error", "Error");
    showAlert("Error al buscar en el buzón: " + err.message, "error");
  }
}

async function searchWithGraph(token) {
  const since = new Date();
  since.setDate(since.getDate() - CONFIG.SEARCH_DAYS_BACK);
  let filter = `contains(subject, '${state.bkbl}') and receivedDateTime ge ${since.toISOString()}`;
  if (CONFIG.SEED_BILLING_FROM) filter += ` and contains(from/emailAddress/address, '${CONFIG.SEED_BILLING_FROM}')`;
  const url = `https://graph.microsoft.com/v1.0/me/messages?$filter=${encodeURIComponent(filter)}&$select=id,subject,from,receivedDateTime,hasAttachments&$top=5&$orderby=receivedDateTime desc`;
  const resp = await fetch(url, { headers: { Authorization: "Bearer " + token } });
  if (!resp.ok) { await searchWithGraphFallback(token); return; }
  const data = await resp.json();
  const messages = data.value || [];
  if (messages.length === 0) { await searchWithGraphFallback(token); return; }
  const msg = messages[0];
  state.foundItemId = msg.id; state.foundSubject = msg.subject; state.foundDate = msg.receivedDateTime;
  onBillingEmailFound(msg.subject, msg.from?.emailAddress?.address || "", msg.receivedDateTime, msg.hasAttachments);
}

async function searchWithGraphFallback(token) {
  const since = new Date();
  since.setDate(since.getDate() - CONFIG.SEARCH_DAYS_BACK);
  const filter = `contains(subject, '${state.bkbl}') and receivedDateTime ge ${since.toISOString()}`;
  const url = `https://graph.microsoft.com/v1.0/me/messages?$filter=${encodeURIComponent(filter)}&$select=id,subject,from,receivedDateTime,hasAttachments&$top=10&$orderby=receivedDateTime desc`;
  const resp = await fetch(url, { headers: { Authorization: "Bearer " + token } });
  const data = await resp.json();
  const messages = (data.value || []).filter(m => m.hasAttachments);
  if (messages.length === 0) {
    setStep(2, "error", "No encontrado");
    showAlert(`No se encontró correo con BK/BL <strong>${escHtml(state.bkbl)}</strong>. Verifica que el correo de Seed Billing esté en tu buzón.`, "error");
    return;
  }
  const msg = messages[0];
  state.foundItemId = msg.id; state.foundSubject = msg.subject; state.foundDate = msg.receivedDateTime;
  onBillingEmailFound(msg.subject, msg.from?.emailAddress?.address || "", msg.receivedDateTime, msg.hasAttachments);
}

async function searchWithOfficeJs() {
  const ewsRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Body><m:FindItem Traversal="Shallow"><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties><t:FieldURI FieldURI="item:Subject"/><t:FieldURI FieldURI="item:HasAttachments"/><t:FieldURI FieldURI="message:From"/></t:AdditionalProperties></m:ItemShape>
  <m:Restriction><t:Contains ContainmentMode="Substring" ContainmentComparison="IgnoreCase"><t:FieldURI FieldURI="item:Subject"/><t:Constant Value="${escHtml(state.bkbl)}"/></t:Contains></m:Restriction>
  <m:ParentFolderIds><t:DistinguishedFolderId Id="inbox"/></m:ParentFolderIds></m:FindItem></soap:Body></soap:Envelope>`;
  Office.context.mailbox.makeEwsRequestAsync(ewsRequest, (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      setStep(2, "error", "Error EWS");
      showAlert("No se pudo buscar automáticamente. Intenta ingresar el BK/BL manualmente.", "error");
      return;
    }
    const xml = result.value;
    const idMatch = xml.match(/Id="([^"]+)"/);
    if (!idMatch) { setStep(2, "error", "No encontrado"); showAlert(`No se encontró correo con BK/BL <strong>${escHtml(state.bkbl)}</strong>.`, "error"); return; }
    state.foundItemId = idMatch[1];
    const subjMatch = xml.match(/<t:Subject>([^<]+)<\/t:Subject>/);
    state.foundSubject = subjMatch ? subjMatch[1] : "Correo de Seed Billing";
    onBillingEmailFound(state.foundSubject, "(via EWS)", null, true);
  });
}

function onBillingEmailFound(subject, from, date, hasAttachments) {
  setStep(2, "done", "Encontrado");
  document.getElementById("foundSubject").textContent = subject;
  const dateStr = date ? new Date(date).toLocaleDateString("es-EC", {day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"}) : "";
  document.getElementById("foundMeta").textContent = `De: ${from}` + (dateStr ? ` · ${dateStr}` : "") + (hasAttachments ? " · 📎 Con adjuntos" : "");
  document.getElementById("foundEmail").classList.add("visible");
  document.getElementById("msgBody").value = CONFIG.DEFAULT_MESSAGE;
  document.getElementById("msgSection").classList.add("visible");
  setStep(3, "active");
  document.getElementById("btnSend").disabled = false;
  hideAlert();
}

function resetMessage() { document.getElementById("msgBody").value = CONFIG.DEFAULT_MESSAGE; }

async function runFlow() {
  if (!state.foundItemId || !state.clientEmail) { showAlert("Faltan datos para enviar.", "error"); return; }
  setBtnLoading(true); setStep(3, "active");
  try {
    if (state.accessToken) { await sendViaGraph(); } else { await sendViaEws(); }
    setStep(3, "done", "Enviado ✓");
    showAlert(`✅ Factura enviada a <strong>${escHtml(state.clientEmail)}</strong>.`, "success");
    document.getElementById("btnSend").disabled = true;
    document.getElementById("btnLabel").textContent = "✓ Enviado";
  } catch (err) {
    setStep(3, "error", "Error");
    showAlert("Error al enviar: " + err.message, "error");
  } finally { setBtnLoading(false); }
}

async function sendViaGraph() {
  const token = state.accessToken;
  const bodyMsg = document.getElementById("msgBody").value;
  const fwdResp = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${state.foundItemId}/createForward`, {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ message: { subject: `Factura / BK-BL: ${state.bkbl}`, toRecipients: [{ emailAddress: { address: state.clientEmail, name: state.clientName } }] }, comment: bodyMsg })
  });
  if (!fwdResp.ok) throw new Error("createForward: " + await fwdResp.text());
  const draftId = (await fwdResp.json()).id;
  const sendResp = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${draftId}/send`, { method: "POST", headers: { Authorization: "Bearer " + token } });
  if (!sendResp.ok) throw new Error("send: " + await sendResp.text());
}

async function sendViaEws() {
  const bodyMsg = document.getElementById("msgBody").value.replace(/\n/g, "<br/>");
  const ews = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Body><m:CreateItem MessageDisposition="SendAndSaveCopy"><m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems"/></m:SavedItemFolderId>
  <m:Items><t:Message><t:Subject>Factura / BK-BL: ${escHtml(state.bkbl)}</t:Subject><t:Body BodyType="HTML">${escHtml(bodyMsg)}</t:Body>
  <t:ToRecipients><t:Mailbox><t:Name>${escHtml(state.clientName)}</t:Name><t:EmailAddress>${escHtml(state.clientEmail)}</t:EmailAddress></t:Mailbox></t:ToRecipients>
  <t:Attachments><t:ItemAttachment><t:Name>Factura_${escHtml(state.bkbl)}.eml</t:Name><t:Message><t:ItemId Id="${escHtml(state.foundItemId)}"/></t:Message></t:ItemAttachment></t:Attachments>
  </t:Message></m:Items></m:CreateItem></soap:Body></soap:Envelope>`;
  await new Promise((resolve, reject) => {
    Office.context.mailbox.makeEwsRequestAsync(ews, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) reject(new Error(result.error.message));
      else resolve(result.value);
    });
  });
}

async function getAccessToken() {
  try {
    return await new Promise((resolve, reject) => {
      Office.context.auth.getAccessTokenAsync({ allowSignInPrompt: true, allowConsentPrompt: true }, (result) => {
        if (result.status === "succeeded") resolve(result.value); else reject(new Error(result.error.message));
      });
    });
  } catch (e) { console.warn("Sin token Graph, usando EWS:", e.message); return null; }
}

function setStep(num, status, badgeText) {
  const step = document.getElementById("step"+num);
  const badge = document.getElementById("badge"+num);
  const icon = step.querySelector(".step-icon");
  step.classList.remove("active","done","error");
  badge.classList.remove("badge-wait","badge-active","badge-done","badge-error");
  if (status==="active") { step.classList.add("active"); badge.classList.add("badge-active"); badge.textContent=badgeText||"En proceso..."; icon.innerHTML=`<div style="width:12px;height:12px;border:2px solid rgba(255,255,255,.3);border-top-color:#fff;border-radius:50%;animation:spin .7s linear infinite"></div>`; }
  else if (status==="done") { step.classList.add("done"); badge.classList.add("badge-done"); badge.textContent=badgeText||"Listo"; icon.textContent="✓"; }
  else if (status==="error") { step.classList.add("error"); badge.classList.add("badge-error"); badge.textContent=badgeText||"Error"; icon.textContent="✕"; }
  else { badge.classList.add("badge-wait"); badge.textContent=badgeText||"Espera"; icon.textContent=num; }
}

function showAlert(html, type="info") { const el=document.getElementById("mainAlert"); el.innerHTML=html; el.className="alert "+type+" visible"; }
function hideAlert() { document.getElementById("mainAlert").classList.remove("visible"); }
function setBtnLoading(loading) {
  const spinner=document.getElementById("btnSpinner"); const label=document.getElementById("btnLabel"); const btn=document.getElementById("btnSend");
  if (loading) { spinner.classList.add("visible"); label.textContent="Enviando..."; btn.disabled=true; }
  else { spinner.classList.remove("visible"); label.textContent="Enviar Factura al Cliente"; btn.disabled=false; }
}
function escHtml(str) { if(!str)return""; return String(str).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
