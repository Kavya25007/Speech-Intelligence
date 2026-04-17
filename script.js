const STORAGE_KEY = "azure_speech_settings";
const HISTORY_KEY = "azure_speech_history";
const SESSION_KEY = "azure_speech_sessions";

const pages = {
  dashboard: document.getElementById("dashboardPage"),
  speech: document.getElementById("speechPage"),
  tts: document.getElementById("ttsPage"),
  translate: document.getElementById("translatePage"),
  pronounce: document.getElementById("pronouncePage"),
  voiceLab: document.getElementById("voiceLabPage"),
  upload: document.getElementById("uploadPage"),
  history: document.getElementById("historyPage"),
  settings: document.getElementById("settingsPage")
};

const titleMap = {
  dashboard: "Dashboard",
  speech: "Speech to Text",
  tts: "Text to Speech",
  translate: "Translation",
  pronounce: "Pronunciation Assessment",
  voiceLab: "Voice Lab",
  upload: "File Upload",
  history: "History",
  settings: "Settings"
};

let currentRecognizer = null;
let currentTranslator = null;
let currentTts = null;
let currentPronounce = null;
let currentSpeechStart = 0;
let finalTranscript = "";

function showToast(message) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.classList.add("show");
  setTimeout(() => toast.classList.remove("show"), 2400);
}

function getSettings() { return JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}"); }
function saveSettings(data) { localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); }
function getHistory() { return JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]"); }
function saveHistory(item) { const arr = getHistory(); arr.unshift(item); localStorage.setItem(HISTORY_KEY, JSON.stringify(arr.slice(0, 20))); }
function getSessions() { return JSON.parse(localStorage.getItem(SESSION_KEY) || "[]"); }
function saveSession(item) { const arr = getSessions(); arr.unshift(item); localStorage.setItem(SESSION_KEY, JSON.stringify(arr.slice(0, 50))); }
function formatText(text) { return (text || "").replace(/\s+/g, " ").trim(); }

function setPage(name) {
  Object.values(pages).forEach(p => p.classList.remove("active"));
  pages[name].classList.add("active");
  document.getElementById("pageTitle").textContent = titleMap[name];
  document.querySelectorAll(".nav-btn").forEach(btn => btn.classList.toggle("active", btn.dataset.page === name));
  if (name === "history") renderHistory();
  if (name === "dashboard") loadDashboard();
}

function loadDashboard() {
  const history = getHistory();
  const sessions = getSessions();
  const wordCount = history.reduce((sum, item) => sum + formatText(item.text).split(" ").filter(Boolean).length, 0);
  document.getElementById("statSessions").textContent = sessions.length;
  document.getElementById("statWords").textContent = wordCount;
  document.getElementById("statWpm").textContent = sessions.length ? Math.round(wordCount / Math.max(sessions.length, 1)) * 5 : 0;
  document.getElementById("statAccuracy").textContent = sessions.length ? "95%" : "—";
  const list = document.getElementById("recentTranscripts");
  list.innerHTML = history.length ? history.map(item => `<div class="list-item"><div class="muted">${new Date(item.time).toLocaleString()}</div><div>${item.text}</div></div>`).join("") : `<div class="list-item">No transcripts yet.</div>`;
}

function loadSettingsToForm() {
  const s = getSettings();
  document.getElementById("apiKey").value = s.apiKey || "";
  document.getElementById("region").value = s.region || "";
  document.getElementById("endpoint").value = s.endpoint || "";
  document.getElementById("lang").value = s.lang || "en-US";
  document.getElementById("theme").value = s.theme || "dark";
  document.getElementById("connStatus").textContent = s.apiKey && s.region ? "Configured" : "Not configured";
  document.getElementById("settingsStatus").textContent = s.apiKey && s.region ? "Settings loaded from localStorage." : "Add your Azure Speech credentials.";
}

function updateSpeechStats(text) {
  const words = formatText(text).split(" ").filter(Boolean).length;
  const mins = Math.max((Date.now() - currentSpeechStart) / 60000, 0.01);
  const wpm = Math.round(words / mins);
  document.getElementById("wordCountChip").textContent = `Words: ${words}`;
  document.getElementById("wpmChip").textContent = `WPM: ${wpm}`;
}

function resetSpeechUI() {
  document.getElementById("liveDot").classList.remove("live");
  document.getElementById("liveState").textContent = "Idle";
  document.getElementById("startMicBtn").disabled = false;
  document.getElementById("stopMicBtn").disabled = true;
  document.getElementById("waveform").style.display = "none";
}

function safeClose(obj) { try { if (obj) obj.close(); } catch (_) {} }

async function startLiveSpeech() {
  const s = getSettings();
  if (!s.apiKey || !s.region) { showToast("Please save Azure key and region in Settings."); setPage("settings"); return; }
  const speechConfig = SpeechSDK.SpeechConfig.fromSubscription(s.apiKey.trim(), s.region.trim());
  speechConfig.speechRecognitionLanguage = (s.lang || "en-US").trim();
  const audioConfig = SpeechSDK.AudioConfig.fromDefaultMicrophoneInput();
  currentRecognizer = new SpeechSDK.SpeechRecognizer(speechConfig, audioConfig);
  currentSpeechStart = Date.now();
  finalTranscript = "";
  document.getElementById("liveDot").classList.add("live");
  document.getElementById("liveState").textContent = "Listening";
  document.getElementById("waveform").style.display = "flex";
  document.getElementById("startMicBtn").disabled = true;
  document.getElementById("stopMicBtn").disabled = false;
  document.getElementById("partialTranscript").textContent = "Waiting...";
  document.getElementById("liveTranscript").textContent = "";
  document.getElementById("langChip").textContent = `Lang: ${(s.lang || "en-US").trim()}`;
  currentRecognizer.recognizing = (_, e) => {
    const partial = formatText(e.result.text);
    document.getElementById("partialTranscript").textContent = partial || "Listening...";
    document.getElementById("liveTranscript").innerHTML = `${finalTranscript}<span style="color:#93a9c4">${partial}</span>`;
    updateSpeechStats(finalTranscript + " " + partial);
  };
  currentRecognizer.recognized = (_, e) => {
    if (e.result.reason === SpeechSDK.ResultReason.RecognizedSpeech) {
      const text = formatText(e.result.text);
      if (text) {
        finalTranscript += (finalTranscript ? " " : "") + text;
        document.getElementById("liveTranscript").textContent = finalTranscript;
        document.getElementById("partialTranscript").textContent = "";
        saveHistory({ time: Date.now(), text, type: "stt" });
        updateSpeechStats(finalTranscript);
      }
    }
  };
  currentRecognizer.canceled = (_, e) => {
    document.getElementById("liveTranscript").textContent = `Canceled: ${e.reason || "unknown"} | ${e.errorDetails || "no details"}`;
    if ((e.errorDetails || "").includes("1006") || (e.errorDetails || "").includes("Unable to contact server")) showToast("WebSocket 1006: check key, region, network, or endpoint.");
    stopLiveSpeech();
  };
  currentRecognizer.sessionStopped = () => stopLiveSpeech();
  try { await currentRecognizer.startContinuousRecognitionAsync(); showToast("Mic started."); } catch (err) { document.getElementById("liveTranscript").textContent = err.message || String(err); showToast("Failed to start recognition."); stopLiveSpeech(); }
}

async function stopLiveSpeech() {
  if (currentRecognizer) {
    try { await currentRecognizer.stopContinuousRecognitionAsync(); } catch (_) {}
    safeClose(currentRecognizer);
    currentRecognizer = null;
  }
  if (finalTranscript.trim()) saveSession({ time: Date.now(), text: finalTranscript.trim(), type: "stt" });
  resetSpeechUI();
}

function makeSsml(text, voiceName, lang, preset, rate, pitch) {
  const styleMap = {
    neutral: "",
    bollywood: `<mstts:express-as style="cheerful">`,
    dramatic: `<mstts:express-as style="serious">`,
    comic: `<mstts:express-as style="chat">`,
    news: `<mstts:express-as style="newscast">`,
    calm: `<mstts:express-as style="calm">`
  };
  const closeStyle = styleMap[preset] ? `</mstts:express-as>` : "";
  const openStyle = styleMap[preset] || "";
  return `
<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xmlns:mstts="http://www.w3.org/2001/mstts" xml:lang="${lang}">
  <voice name="${voiceName}">
    ${openStyle}
      <prosody rate="${rate}" pitch="${pitch}%">
        ${text}
      </prosody>
    ${closeStyle}
  </voice>
</speak>`.trim();
}

function speakSsml(ssml) {
  const s = getSettings();
  if (!s.apiKey || !s.region) { showToast("Please configure Azure settings first."); setPage("settings"); return; }
  safeClose(currentTts);
  const cfg = SpeechSDK.SpeechConfig.fromSubscription(s.apiKey.trim(), s.region.trim());
  currentTts = new SpeechSDK.SpeechSynthesizer(cfg);
  document.getElementById("ttsStatus").textContent = "Speaking...";
  currentTts.speakSsmlAsync(ssml, () => { document.getElementById("ttsStatus").textContent = "Playback finished."; safeClose(currentTts); currentTts = null; }, err => { document.getElementById("ttsStatus").textContent = err; safeClose(currentTts); currentTts = null; });
}

async function startTranslation() {
  const mode = document.getElementById("translateMode").value;
  const s = getSettings();
  if (!s.apiKey || !s.region) { showToast("Please configure Azure settings first."); setPage("settings"); return; }
  const targetLang = document.getElementById("targetLang").value;
  document.getElementById("origText").textContent = "";
  document.getElementById("translatedText").textContent = "";
  if (mode === "text") {
    const txt = document.getElementById("translateInput").value.trim();
    document.getElementById("origText").textContent = txt;
    document.getElementById("translatedText").textContent = `[mock] translated to ${targetLang}: ${txt}`;
    return;
  }
  const cfg = SpeechSDK.SpeechTranslationConfig.fromSubscription(s.apiKey.trim(), s.region.trim());
  cfg.speechRecognitionLanguage = (document.getElementById("sourceLang").value || "en-US").trim();
  cfg.addTargetLanguage(targetLang);
  currentTranslator = new SpeechSDK.TranslationRecognizer(cfg, SpeechSDK.AudioConfig.fromDefaultMicrophoneInput());
  document.getElementById("startTranslateBtn").disabled = true;
  document.getElementById("stopTranslateBtn").disabled = false;
  currentTranslator.recognizing = (_, e) => { document.getElementById("origText").textContent = formatText(e.result.text); };
  currentTranslator.recognized = (_, e) => {
    if (e.result.reason === SpeechSDK.ResultReason.TranslatedSpeech) {
      document.getElementById("origText").textContent = formatText(e.result.text);
      document.getElementById("translatedText").textContent = formatText(e.result.translations.get(targetLang) || "");
    }
  };
  currentTranslator.canceled = (_, e) => { document.getElementById("origText").textContent = `Canceled: ${e.reason || "unknown"} | ${e.errorDetails || "no details"}`; if ((e.errorDetails || "").includes("1006")) showToast("Translation WebSocket 1006."); stopTranslation(); };
  try { await currentTranslator.startContinuousRecognitionAsync(); showToast("Translation started."); } catch (err) { document.getElementById("origText").textContent = err.message || String(err); stopTranslation(); }
}

async function stopTranslation() {
  if (currentTranslator) { try { await currentTranslator.stopContinuousRecognitionAsync(); } catch (_) {} safeClose(currentTranslator); currentTranslator = null; }
  document.getElementById("startTranslateBtn").disabled = false;
  document.getElementById("stopTranslateBtn").disabled = true;
}

async function startPronunciation() {
  const s = getSettings();
  if (!s.apiKey || !s.region) { showToast("Please configure Azure settings first."); setPage("settings"); return; }
  const ref = document.getElementById("referenceText").value.trim();
  const lang = document.getElementById("pronounceLang").value.trim() || "en-US";
  const speechConfig = SpeechSDK.SpeechConfig.fromSubscription(s.apiKey.trim(), s.region.trim());
  speechConfig.speechRecognitionLanguage = lang;
  currentPronounce = new SpeechSDK.SpeechRecognizer(speechConfig, SpeechSDK.AudioConfig.fromDefaultMicrophoneInput());
  const pronunciationConfig = new SpeechSDK.PronunciationAssessmentConfig(ref, SpeechSDK.PronunciationAssessmentGradingSystem.HundredMark, SpeechSDK.PronunciationAssessmentGranularity.Phoneme, true);
  pronunciationConfig.applyTo(currentPronounce);
  document.getElementById("startPronounceBtn").disabled = true;
  document.getElementById("stopPronounceBtn").disabled = false;
  currentPronounce.recognized = (_, e) => {
    if (e.result.reason === SpeechSDK.ResultReason.RecognizedSpeech) {
      const res = SpeechSDK.PronunciationAssessmentResult.fromResult(e.result);
      document.getElementById("pronounceResult").textContent = formatText(e.result.text);
      document.getElementById("scoreAccuracy").textContent = res.accuracyScore.toFixed(0);
      document.getElementById("scoreFluency").textContent = res.fluencyScore.toFixed(0);
      document.getElementById("scoreCompleteness").textContent = res.completenessScore.toFixed(0);
      document.getElementById("scoreProsody").textContent = res.prosodyScore ? res.prosodyScore.toFixed(0) : "—";
    }
  };
  currentPronounce.canceled = (_, e) => { document.getElementById("pronounceResult").textContent = `Canceled: ${e.reason || "unknown"} | ${e.errorDetails || "no details"}`; stopPronunciation(); };
  try { await currentPronounce.startContinuousRecognitionAsync(); showToast("Pronunciation assessment started."); } catch (err) { document.getElementById("pronounceResult").textContent = err.message || String(err); stopPronunciation(); }
}

async function stopPronunciation() {
  if (currentPronounce) { try { await currentPronounce.stopContinuousRecognitionAsync(); } catch (_) {} safeClose(currentPronounce); currentPronounce = null; }
  document.getElementById("startPronounceBtn").disabled = false;
  document.getElementById("stopPronounceBtn").disabled = true;
}

function renderHistory() {
  const q = document.getElementById("historySearch").value.toLowerCase().trim();
  const list = getHistory().filter(x => !q || (x.text || "").toLowerCase().includes(q));
  const box = document.getElementById("historyList");
  box.innerHTML = list.length ? list.map(item => `<div class="list-item"><div class="muted">${new Date(item.time).toLocaleString()} • ${item.type || "stt"}</div><div>${item.text}</div></div>`).join("") : `<div class="list-item">No matching history.</div>`;
}

function exportHistory() {
  const data = JSON.stringify(getHistory(), null, 2);
  const blob = new Blob([data], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "speech-history.json";
  a.click();
}

function bindNavigation() {
  document.querySelectorAll(".nav-btn, .action-btn").forEach(btn => btn.addEventListener("click", () => setPage(btn.dataset.page)));
  document.getElementById("quickStartBtn").onclick = () => setPage("speech");
}

function bindControls() {
  document.getElementById("startMicBtn").onclick = startLiveSpeech;
  document.getElementById("stopMicBtn").onclick = stopLiveSpeech;
  document.getElementById("copyTranscriptBtn").onclick = async () => navigator.clipboard.writeText(document.getElementById("liveTranscript").innerText || "");
  document.getElementById("downloadTxtBtn").onclick = () => {
    const text = document.getElementById("liveTranscript").innerText || "";
    const blob = new Blob([text], { type: "text/plain" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "transcript.txt";
    a.click();
  };
  document.getElementById("clearTranscriptBtn").onclick = () => { finalTranscript = ""; document.getElementById("liveTranscript").textContent = "Cleared."; document.getElementById("partialTranscript").textContent = ""; };
  document.getElementById("playTtsBtn").onclick = () => {
    const text = document.getElementById("ttsText").value.trim();
    const voice = document.getElementById("voiceSelect").value;
    const rate = document.getElementById("ttsRate").value;
    const s = getSettings();
    const lang = (s.lang || "en-US").trim();
    const ssml = `<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xmlns:mstts="http://www.w3.org/2001/mstts" xml:lang="${lang}"><voice name="${voice}"><prosody rate="${rate}" pitch="0%">${text}</prosody></voice></speak>`;
    speakSsml(ssml);
  };
  document.getElementById("stopTtsBtn").onclick = () => { safeClose(currentTts); currentTts = null; document.getElementById("ttsStatus").textContent = "Stopped."; };
  document.getElementById("startTranslateBtn").onclick = startTranslation;
  document.getElementById("stopTranslateBtn").onclick = stopTranslation;
  document.getElementById("swapLangBtn").onclick = () => {
    const src = document.getElementById("sourceLang").value;
    const tgt = document.getElementById("targetLang").value;
    document.getElementById("sourceLang").value = tgt;
    document.getElementById("targetLang").value = src.startsWith("en") ? "en" : "hi";
  };
  document.getElementById("speakTranslatedBtn").onclick = () => {
    const text = document.getElementById("translatedText").innerText.trim();
    const voice = document.getElementById("voiceSelect").value;
    const rate = document.getElementById("ttsRate").value;
    const s = getSettings();
    const lang = (s.lang || "en-US").trim();
    speakSsml(`<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xmlns:mstts="http://www.w3.org/2001/mstts" xml:lang="${lang}"><voice name="${voice}"><prosody rate="${rate}" pitch="0%">${text}</prosody></voice></speak>`);
  };
  document.getElementById("startPronounceBtn").onclick = startPronunciation;
  document.getElementById("stopPronounceBtn").onclick = stopPronunciation;
  document.getElementById("copyScoreBtn").onclick = async () => navigator.clipboard.writeText(`Accuracy: ${document.getElementById("scoreAccuracy").textContent}, Fluency: ${document.getElementById("scoreFluency").textContent}, Completeness: ${document.getElementById("scoreCompleteness").textContent}, Prosody: ${document.getElementById("scoreProsody").textContent}`);
  document.getElementById("audioFile").onchange = e => { const file = e.target.files[0]; document.getElementById("fileInfo").textContent = file ? `Selected: ${file.name} (${Math.round(file.size / 1024)} KB)` : "No file selected."; };
  document.getElementById("simulateUploadBtn").onclick = () => document.getElementById("uploadStatus").textContent = "Simulated processing complete. Ready to send to Azure batch transcription.";
  document.getElementById("clearUploadBtn").onclick = () => { document.getElementById("audioFile").value = ""; document.getElementById("fileInfo").textContent = "No file selected."; document.getElementById("uploadStatus").textContent = "Ready for Azure batch transcription or simulation."; };
  document.getElementById("saveSettingsBtn").onclick = () => {
    saveSettings({
      apiKey: document.getElementById("apiKey").value.trim(),
      region: document.getElementById("region").value.trim(),
      endpoint: document.getElementById("endpoint").value.trim(),
      lang: document.getElementById("lang").value.trim() || "en-US",
      theme: document.getElementById("theme").value
    });
    loadSettingsToForm();
    showToast("Settings saved.");
  };
  document.getElementById("testSettingsBtn").onclick = () => {
    const s = getSettings();
    if (!s.apiKey || !s.region) { document.getElementById("settingsStatus").textContent = "Missing API key or region."; showToast("Please fill API key and region."); return; }
    document.getElementById("settingsStatus").textContent = "Saved settings found. If 1006 continues, check network, firewall, region, or endpoint.";
    showToast("Settings look ready.");
  };
  document.getElementById("fillDemoBtn").onclick = () => {
    document.getElementById("apiKey").value = "PASTE_YOUR_KEY";
    document.getElementById("region").value = "centralindia";
    document.getElementById("lang").value = "en-US";
    document.getElementById("theme").value = "dark";
    showToast("Demo fields filled.");
  };
  document.getElementById("historySearch").oninput = renderHistory;
  document.getElementById("exportHistoryBtn").onclick = exportHistory;
  document.getElementById("clearHistoryBtn").onclick = () => { localStorage.removeItem(HISTORY_KEY); renderHistory(); loadDashboard(); };
  document.getElementById("clearRecentBtn").onclick = () => { localStorage.removeItem(HISTORY_KEY); localStorage.removeItem(SESSION_KEY); loadDashboard(); renderHistory(); };
  document.getElementById("applyPresetBtn").onclick = () => {
    const preset = document.getElementById("voicePreset").value;
    const map = {
      neutral: { voice: "en-US-JennyNeural", rate: 1, pitch: 0, text: "This is a neutral preview." },
      bollywood: { voice: "en-IN-PrabhatNeural", rate: 1.05, pitch: 8, text: "Arre waah! Aaj ka cinematic voice lab shuru hota hai." },
      dramatic: { voice: "en-US-GuyNeural", rate: 0.9, pitch: -4, text: "Tonight... the story begins." },
      comic: { voice: "en-US-JennyNeural", rate: 1.15, pitch: 6, text: "Haha! That was a fun line!" },
      news: { voice: "en-US-GuyNeural", rate: 1, pitch: -1, text: "Breaking news: voice presets are now live." },
      calm: { voice: "en-US-JennyNeural", rate: 0.92, pitch: -2, text: "Take a breath and listen calmly." }
    };
    const p = map[preset];
    document.getElementById("voiceLabSelect").value = p.voice;
    document.getElementById("voiceSpeed").value = p.rate;
    document.getElementById("voicePitch").value = p.pitch;
    document.getElementById("voiceLabText").value = p.text;
    document.getElementById("voiceLabOutput").textContent = `Preset applied: ${preset}`;
  };
  document.getElementById("previewVoiceBtn").onclick = () => {
    const voice = document.getElementById("voiceLabSelect").value;
    const text = document.getElementById("voiceLabText").value.trim();
    const lang = document.getElementById("voiceLang").value;
    const preset = document.getElementById("voicePreset").value;
    const rateVal = parseFloat(document.getElementById("voiceSpeed").value);
    const pitchVal = parseInt(document.getElementById("voicePitch").value, 10);
    const rate = rateVal === 1 ? "0%" : (rateVal > 1 ? `+${Math.round((rateVal - 1) * 100)}%` : `-${Math.round((1 - rateVal) * 100)}%`);
    const ssml = makeSsml(text, voice, lang, preset, rate, pitchVal);
    document.getElementById("voiceLabOutput").textContent = `Playing ${preset} style with ${voice}.`;
    speakSsml(ssml);
  };
  document.getElementById("randomVoiceBtn").onclick = () => {
    const voices = ["en-US-GuyNeural","en-US-JennyNeural","en-IN-PrabhatNeural","hi-IN-MadhurNeural","hi-IN-SwaraNeural"];
    const presets = ["neutral","bollywood","dramatic","comic","news","calm"];
    document.getElementById("voiceLabSelect").value = voices[Math.floor(Math.random() * voices.length)];
    document.getElementById("voicePreset").value = presets[Math.floor(Math.random() * presets.length)];
    showToast("Random voice selected.");
  };
  document.getElementById("copySsmlBtn").onclick = async () => {
    const voice = document.getElementById("voiceLabSelect").value;
    const text = document.getElementById("voiceLabText").value.trim();
    const lang = document.getElementById("voiceLang").value;
    const preset = document.getElementById("voicePreset").value;
    const rateVal = parseFloat(document.getElementById("voiceSpeed").value);
    const pitchVal = parseInt(document.getElementById("voicePitch").value, 10);
    const rate = rateVal === 1 ? "0%" : (rateVal > 1 ? `+${Math.round((rateVal - 1) * 100)}%` : `-${Math.round((1 - rateVal) * 100)}%`);
    const ssml = makeSsml(text, voice, lang, preset, rate, pitchVal);
    await navigator.clipboard.writeText(ssml);
    showToast("SSML copied.");
  };
}

function init() {
  bindNavigation();
  bindControls();
  loadSettingsToForm();
  loadDashboard();
  renderHistory();
  setPage("dashboard");
}

window.addEventListener("load", init);
