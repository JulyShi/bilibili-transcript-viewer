// ==UserScript==
// @name         Bilibili Transcript Viewer
// @namespace    https://github.com/github-copilot
// @version      0.1.0
// @description  Show Bilibili subtitles in a synced transcript panel while watching videos.
// @author       GitHub Copilot
// @match        https://www.bilibili.com/video/*
// @match        https://www.bilibili.com/bangumi/play/*
// @require      https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js
// @grant        GM_xmlhttpRequest
// @connect      cdn.jsdelivr.net
// @connect      api.bilibili.com
// @connect      *.bilibili.com
// @connect      *.hdslb.com
// ==/UserScript==

(function () {
  'use strict';

  const PANEL_ID = 'bili-transcript-viewer-panel';
  const STYLE_ID = 'bili-transcript-viewer-style';
  const STORAGE_KEY = 'bili-transcript-viewer-panel-state';
  const ACTIVE_CLASS = 'bili-transcript-line-active';
  const API_ORIGIN = 'https://api.bilibili.com';
  const CAPTURE_FALLBACK_WIDTH = 1280;
  const CAPTURE_MIN_GAP_SECONDS = 1.2;
  const CAPTURE_AUDIO_THRESHOLD = 0.028;
  const CAPTURE_SILENCE_FRAMES = 10;
  const CAPTURE_SEGMENT_MIN_SECONDS = 0.2;
  const CAPTURE_MAX_SEGMENT_SECONDS = 1.35;
  const CAPTURE_SILENCE_HOLD_SECONDS = 0.55;
  const CAPTURE_PRE_ROLL_MAX_SECONDS = 2.0;
  const CAPTURE_LOOP_INTERVAL_MS = 80;
  const PPTX_LIB_URL = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
  const CAPTURE_AUDIO_BINDING_KEY = '__biliTranscriptCaptureAudioBinding';

  const state = {
    video: null,
    cues: [],
    activeCueIndex: -1,
    selectedTrackUrl: '',
    playbackHandler: null,
    retryTimer: null,
    lastUrl: location.href,
    routeWatcherStarted: false,
    captureItems: [],
    captureStatus: '准备好后点击“开始采集”，脚本会按语音片段自动截取画面。',
    captureRunning: false,
    captureLoopHandle: 0,
    capturePendingShotTimer: null,
    captureAudioContext: null,
    captureAnalyser: null,
    captureSourceNode: null,
    captureBoundVideo: null,
    captureSampleBuffer: null,
    captureVoiceActive: false,
    captureSilenceFrames: 0,
    captureSilenceStartedAt: -1,
    captureLastShotAt: -Infinity,
    captureMode: 'audio',
    captureSeenCueIndices: new Set(),
    capturePendingCueIndex: -1,
    captureSegmentStartTime: -1,
    captureSegmentPeakTime: -1,
    captureSegmentPeakLevel: 0,
    captureSegmentSnapshot: null,
    captureSegmentTailSnapshot: null,
    captureCapturedPreRoll: false,
    capturePreRollSamples: 0,
    captureFirstVoiceSegmentStarted: false,
    pptxLibraryPromise: null,
  };

  function ensureStyles() {
    if (document.getElementById(STYLE_ID)) {
      return;
    }

    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = `
      #${PANEL_ID} {
        position: fixed;
        top: 18px;
        right: 18px;
        z-index: 999999;
        width: min(390px, calc(100vw - 36px));
        max-height: calc(100vh - 36px);
        display: flex;
        flex-direction: column;
        border-radius: 18px;
        overflow: hidden;
        border: 1px solid rgba(124, 93, 59, 0.18);
        box-shadow: 0 18px 54px rgba(30, 24, 18, 0.22);
        background: rgba(247, 243, 235, 0.96);
        color: #1e1812;
        backdrop-filter: blur(14px);
        font-family: "Source Han Serif SC", "Noto Serif SC", "PingFang SC", serif;
      }

      #${PANEL_ID}.is-collapsed {
        width: 228px;
        max-height: 84px;
      }

      #${PANEL_ID} * {
        box-sizing: border-box;
      }

      #${PANEL_ID} .bili-transcript-header {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 14px 16px;
        background: linear-gradient(135deg, #e5cfab 0%, #f8efe1 100%);
        border-bottom: 1px solid rgba(124, 93, 59, 0.14);
      }

      #${PANEL_ID} .bili-transcript-title {
        flex: 1;
        min-width: 0;
        margin: 0;
        font-size: 15px;
        line-height: 1.35;
        font-weight: 700;
      }

      #${PANEL_ID} .bili-transcript-body {
        display: flex;
        flex-direction: column;
        gap: 12px;
        padding: 14px 16px 16px;
      }

      #${PANEL_ID}.is-collapsed .bili-transcript-body {
        display: none;
      }

      #${PANEL_ID} .bili-transcript-current {
        min-height: 76px;
        padding: 14px;
        border-radius: 14px;
        border: 1px solid rgba(124, 93, 59, 0.12);
        background: linear-gradient(180deg, rgba(255,255,255,0.78), rgba(255,248,239,0.98));
        font-size: 24px;
        line-height: 1.5;
      }

      #${PANEL_ID} .bili-transcript-current.is-empty,
      #${PANEL_ID} .bili-transcript-status {
        color: #6f5a46;
        font-size: 14px;
      }

      #${PANEL_ID} .bili-transcript-controls {
        display: flex;
        align-items: center;
        gap: 8px;
      }

      #${PANEL_ID} .bili-transcript-capture-controls {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
      }

      #${PANEL_ID} .bili-transcript-select {
        flex: 1;
        min-width: 0;
        border-radius: 999px;
        border: 1px solid rgba(124, 93, 59, 0.22);
        background: rgba(255,255,255,0.84);
        color: #2f241b;
        padding: 8px 12px;
        font-size: 13px;
      }

      #${PANEL_ID} .bili-transcript-list {
        display: flex;
        flex-direction: column;
        gap: 8px;
        overflow: auto;
        max-height: min(52vh, 420px);
        padding-right: 4px;
      }

      #${PANEL_ID} .bili-capture-section {
        display: flex;
        flex-direction: column;
        gap: 8px;
        padding-top: 4px;
        border-top: 1px solid rgba(124, 93, 59, 0.12);
      }

      #${PANEL_ID} .bili-capture-heading {
        margin: 0;
        font-size: 13px;
        line-height: 1.4;
        font-weight: 700;
        color: #4d3a2c;
      }

      #${PANEL_ID} .bili-capture-status {
        color: #6f5a46;
        font-size: 13px;
        line-height: 1.5;
      }

      #${PANEL_ID} .bili-capture-list {
        display: flex;
        flex-direction: column;
        gap: 8px;
        overflow: auto;
        max-height: min(34vh, 280px);
        padding-right: 4px;
      }

      #${PANEL_ID} .bili-capture-card {
        display: grid;
        grid-template-columns: 92px 1fr;
        gap: 10px;
        width: 100%;
        padding: 8px;
        border: 1px solid rgba(124, 93, 59, 0.12);
        border-radius: 12px;
        background: rgba(255,255,255,0.72);
        text-align: left;
      }

      #${PANEL_ID} .bili-capture-thumb {
        width: 92px;
        height: 52px;
        border-radius: 8px;
        object-fit: cover;
        background: rgba(56, 39, 25, 0.08);
      }

      #${PANEL_ID} .bili-capture-meta {
        display: flex;
        flex-direction: column;
        gap: 4px;
        min-width: 0;
      }

      #${PANEL_ID} .bili-capture-time {
        color: #7b644c;
        font-size: 12px;
      }

      #${PANEL_ID} .bili-capture-label {
        color: #201812;
        font-size: 14px;
        line-height: 1.45;
        word-break: break-word;
      }

      #${PANEL_ID} .bili-capture-empty {
        padding: 10px 12px;
        border-radius: 12px;
        background: rgba(255,255,255,0.58);
        color: #6f5a46;
        font-size: 13px;
      }

      #${PANEL_ID} .bili-transcript-line {
        display: grid;
        grid-template-columns: 56px 1fr;
        gap: 10px;
        width: 100%;
        padding: 10px 12px;
        border: 1px solid transparent;
        border-radius: 12px;
        background: rgba(255,255,255,0.76);
        cursor: pointer;
        text-align: left;
      }

      #${PANEL_ID} .bili-transcript-line:hover {
        border-color: rgba(161, 106, 53, 0.28);
        background: rgba(255, 250, 243, 0.98);
      }

      #${PANEL_ID} .${ACTIVE_CLASS} {
        border-color: rgba(161, 106, 53, 0.34);
        background: rgba(245, 228, 198, 0.94);
      }

      #${PANEL_ID} .bili-transcript-time {
        color: #7b644c;
        font-size: 12px;
        line-height: 1.4;
        padding-top: 2px;
      }

      #${PANEL_ID} .bili-transcript-text {
        color: #201812;
        font-size: 15px;
        line-height: 1.55;
        white-space: pre-wrap;
      }

      #${PANEL_ID} button {
        border: none;
        border-radius: 999px;
        background: rgba(56, 39, 25, 0.08);
        color: #382719;
        padding: 8px 10px;
        cursor: pointer;
        font-size: 12px;
      }

      #${PANEL_ID} button:hover {
        background: rgba(56, 39, 25, 0.16);
      }

      #${PANEL_ID} button:disabled {
        opacity: 0.45;
        cursor: not-allowed;
      }

      @media (max-width: 960px) {
        #${PANEL_ID} {
          top: auto;
          right: 12px;
          bottom: 12px;
          left: 12px;
          width: auto;
          max-height: 60vh;
        }

        #${PANEL_ID}.is-collapsed {
          left: auto;
          width: 228px;
        }

        #${PANEL_ID} .bili-transcript-current {
          font-size: 20px;
        }
      }
    `;

    document.head.appendChild(style);
  }

  function formatTime(seconds) {
    const safeSeconds = Math.max(0, Math.floor(Number(seconds) || 0));
    const minutes = Math.floor(safeSeconds / 60);
    const remainder = safeSeconds % 60;
    return `${String(minutes).padStart(2, '0')}:${String(remainder).padStart(2, '0')}`;
  }

  function normalizeSubtitleUrl(url) {
    if (!url) {
      return '';
    }

    try {
      return new URL(url, API_ORIGIN).toString();
    } catch (error) {
      return url;
    }
  }

  function getTrackUrl(track) {
    return normalizeSubtitleUrl(track.subtitle_url || track.url || '');
  }

  function getTrackLabel(track, index) {
    return track.lan_doc || track.lan || `Subtitle ${index + 1}`;
  }

  function getVideoTitle() {
    return document.title.replace(/_哔哩哔哩_bilibili$/, '').trim() || 'Transcript';
  }

  function readPanelState() {
    try {
      return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
    } catch (error) {
      return {};
    }
  }

  function savePanelState(panel) {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({ collapsed: panel.classList.contains('is-collapsed') })
    );
  }

  function clearRetryTimer() {
    if (state.retryTimer) {
      window.clearTimeout(state.retryTimer);
      state.retryTimer = null;
    }
  }

  function detachPlaybackHandler() {
    if (state.video && state.playbackHandler) {
      state.video.removeEventListener('timeupdate', state.playbackHandler);
    }
    state.playbackHandler = null;
    state.video = null;
  }

  function createButton(label, onClick) {
    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = label;
    button.addEventListener('click', onClick);
    return button;
  }

  function setCaptureStatus(panel, message) {
    state.captureStatus = message;
    const status = panel.querySelector('.bili-capture-status');
    if (status) {
      status.textContent = message;
    }
  }

  function updateCaptureButtons(panel) {
    const startButton = panel.querySelector('[data-capture-action="start"]');
    const stopButton = panel.querySelector('[data-capture-action="stop"]');
    const snapshotButton = panel.querySelector('[data-capture-action="snapshot"]');
    const exportButton = panel.querySelector('[data-capture-action="export"]');
    const clearButton = panel.querySelector('[data-capture-action="clear"]');
    const hasItems = state.captureItems.length > 0;

    if (startButton) {
      startButton.disabled = state.captureRunning;
    }
    if (stopButton) {
      stopButton.disabled = !state.captureRunning;
    }
    if (snapshotButton) {
      snapshotButton.disabled = !state.video;
    }
    if (exportButton) {
      exportButton.disabled = !hasItems;
    }
    if (clearButton) {
      clearButton.disabled = !hasItems;
    }
  }

  function renderCaptureList(panel) {
    const list = panel.querySelector('.bili-capture-list');
    if (!list) {
      return;
    }

    list.innerHTML = '';

    if (state.captureItems.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'bili-capture-empty';
      empty.textContent = '还没有采集到画面。开始播放后点击“开始采集”，脚本会按语音片段自动截取。';
      list.appendChild(empty);
      updateCaptureButtons(panel);
      return;
    }

    state.captureItems.forEach((item, index) => {
      const row = document.createElement('button');
      row.type = 'button';
      row.className = 'bili-capture-card';

      const thumb = document.createElement('img');
      thumb.className = 'bili-capture-thumb';
      thumb.src = item.imageDataUrl;
      thumb.alt = `画面 ${index + 1}`;

      const meta = document.createElement('div');
      meta.className = 'bili-capture-meta';

      const time = document.createElement('span');
      time.className = 'bili-capture-time';
      time.textContent = formatTime(item.time);

      const label = document.createElement('span');
      label.className = 'bili-capture-label';
      label.textContent = item.label ? `${index + 1}. ${item.label}` : `第 ${index + 1} 张画面`;

      meta.appendChild(time);
      meta.appendChild(label);
      row.appendChild(thumb);
      row.appendChild(meta);
      row.addEventListener('click', () => {
        if (!state.video) {
          return;
        }

        state.video.currentTime = item.time;
        state.video.pause();
      });
      list.appendChild(row);
    });

    updateCaptureButtons(panel);
  }

  function clearCapturePendingShot() {
    if (state.capturePendingShotTimer) {
      window.clearTimeout(state.capturePendingShotTimer);
      state.capturePendingShotTimer = null;
    }
  }

  function stopCaptureLoop() {
    if (state.captureLoopHandle) {
      window.clearInterval(state.captureLoopHandle);
      state.captureLoopHandle = 0;
    }
    clearCapturePendingShot();
  }

  function startCaptureLoop(panel) {
    stopCaptureLoop();
    state.captureLoopHandle = window.setInterval(() => runCaptureLoop(panel), CAPTURE_LOOP_INTERVAL_MS);
  }

  function getCaptureAudioBinding(video) {
    if (!video) {
      return null;
    }

    return video[CAPTURE_AUDIO_BINDING_KEY] || null;
  }

  function setCaptureAudioBinding(video, binding) {
    if (!video) {
      return;
    }

    if (binding) {
      video[CAPTURE_AUDIO_BINDING_KEY] = binding;
      return;
    }

    delete video[CAPTURE_AUDIO_BINDING_KEY];
  }

  function releaseCaptureAudioBinding(video) {
    const binding = getCaptureAudioBinding(video);
    if (!binding) {
      return;
    }

    if (binding.sourceNode) {
      binding.sourceNode.disconnect();
    }

    if (binding.analyser) {
      binding.analyser.disconnect();
    }

    if (binding.audioContext && binding.audioContext.state !== 'closed') {
      binding.audioContext.close().catch(() => {});
    }

    setCaptureAudioBinding(video, null);
  }

  function teardownCaptureAudio(options = {}) {
    const { release = false } = options;
    stopCaptureLoop();
    state.captureRunning = false;
    state.captureVoiceActive = false;
    state.captureSilenceFrames = 0;
    state.captureSilenceStartedAt = -1;
    state.captureCapturedPreRoll = false;
    state.capturePreRollSamples = 0;
    state.captureFirstVoiceSegmentStarted = false;
    state.captureLastShotAt = -Infinity;

    const boundVideo = state.captureBoundVideo;
    if (release && boundVideo) {
      releaseCaptureAudioBinding(boundVideo);
    } else if (state.captureAudioContext && state.captureAudioContext.state === 'running') {
      state.captureAudioContext.suspend().catch(() => {});
    }

    state.captureSourceNode = null;
    state.captureAnalyser = null;
    state.captureAudioContext = null;
    state.captureBoundVideo = null;
    state.captureSampleBuffer = null;
  }

  function resetCaptureState(panel) {
    teardownCaptureAudio();
    state.captureItems = [];
    state.captureStatus = '准备好后点击“开始采集”，脚本会按语音片段自动截取画面。';
    state.captureMode = 'audio';
    state.captureSeenCueIndices = new Set();
    state.capturePendingCueIndex = -1;
    state.captureSegmentStartTime = -1;
    state.captureSegmentPeakTime = -1;
    state.captureSegmentPeakLevel = 0;
    state.captureSegmentSnapshot = null;
    state.captureSegmentTailSnapshot = null;
    state.captureSilenceStartedAt = -1;

    if (panel) {
      setCaptureStatus(panel, state.captureStatus);
      renderCaptureList(panel);
    }
  }

  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function sanitizeFileName(text) {
    return String(text || 'capture')
      .replace(/[\\/:*?"<>|]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function downloadBlob(filename, content, type) {
    const blob = new Blob([content], { type });
    downloadExistingBlob(filename, blob);
  }

  function downloadExistingBlob(filename, blob) {
    const blobUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = blobUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(blobUrl), 1000);
  }

  function getPptxConstructor() {
    if (typeof PptxGenJS === 'function') {
      return PptxGenJS;
    }

    if (typeof window.PptxGenJS === 'function') {
      return window.PptxGenJS;
    }

    if (typeof globalThis.PptxGenJS === 'function') {
      return globalThis.PptxGenJS;
    }

    return null;
  }

  async function ensurePptxLibrary() {
    const existingConstructor = getPptxConstructor();
    if (existingConstructor) {
      return existingConstructor;
    }

    if (!state.pptxLibraryPromise) {
      state.pptxLibraryPromise = requestText(PPTX_LIB_URL)
        .then((source) => {
          const loadedConstructor = Function(
            `${source}\nreturn typeof PptxGenJS === 'function' ? PptxGenJS : (typeof module !== 'undefined' && typeof module.exports === 'function' ? module.exports : null);`
          )();

          if (typeof loadedConstructor === 'function') {
            window.PptxGenJS = loadedConstructor;
            globalThis.PptxGenJS = loadedConstructor;
          }

          return loadedConstructor || null;
        })
        .then((PptxGenJS) => {
          if (typeof PptxGenJS !== 'function') {
            throw new Error('PPTX 导出库加载后不可用');
          }

          return PptxGenJS;
        })
        .catch((error) => {
          state.pptxLibraryPromise = null;
          throw error;
        });
    }

    return state.pptxLibraryPromise;
  }

  function sampleCaptureLevel() {
    if (!state.captureAnalyser || !state.captureSampleBuffer) {
      return 0;
    }

    state.captureAnalyser.getByteTimeDomainData(state.captureSampleBuffer);
    let sum = 0;

    for (const value of state.captureSampleBuffer) {
      const normalized = (value - 128) / 128;
      sum += normalized * normalized;
    }

    return Math.sqrt(sum / state.captureSampleBuffer.length);
  }

  async function ensureCaptureAudio(video) {
    if (!video) {
      throw new Error('还没有找到视频元素。');
    }

    const existingBinding = getCaptureAudioBinding(video);
    if (existingBinding) {
      if (existingBinding.audioContext.state === 'suspended') {
        await existingBinding.audioContext.resume();
      }

      state.captureAudioContext = existingBinding.audioContext;
      state.captureAnalyser = existingBinding.analyser;
      state.captureSourceNode = existingBinding.sourceNode;
      state.captureBoundVideo = video;
      state.captureSampleBuffer = existingBinding.sampleBuffer;
      return;
    }

    if (state.captureBoundVideo && state.captureBoundVideo !== video) {
      releaseCaptureAudioBinding(state.captureBoundVideo);
    }

    teardownCaptureAudio();

    const AudioContextClass = window.AudioContext || window.webkitAudioContext;
    if (!AudioContextClass) {
      throw new Error('当前浏览器不支持音频分析，无法自动采集。');
    }

    const audioContext = new AudioContextClass();
    const analyser = audioContext.createAnalyser();
    analyser.fftSize = 2048;
    analyser.smoothingTimeConstant = 0.85;

    const sourceNode = audioContext.createMediaElementSource(video);
    sourceNode.connect(analyser);
    analyser.connect(audioContext.destination);

    if (audioContext.state === 'suspended') {
      await audioContext.resume();
    }

    const binding = {
      audioContext,
      analyser,
      sourceNode,
      sampleBuffer: new Uint8Array(analyser.fftSize),
    };

    setCaptureAudioBinding(video, binding);

    state.captureAudioContext = binding.audioContext;
    state.captureAnalyser = binding.analyser;
    state.captureSourceNode = binding.sourceNode;
    state.captureBoundVideo = video;
    state.captureSampleBuffer = binding.sampleBuffer;
  }

  function createCaptureSnapshot(label) {
    if (!state.video) {
      throw new Error('当前没有可截图的视频。');
    }

    const video = state.video;
    const width = video.videoWidth || CAPTURE_FALLBACK_WIDTH;
    const aspectRatio = video.videoWidth && video.videoHeight ? video.videoHeight / video.videoWidth : 9 / 16;
    const height = video.videoHeight || Math.max(720, Math.round(width * aspectRatio));
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;

    const context = canvas.getContext('2d');
    context.imageSmoothingEnabled = true;
    context.imageSmoothingQuality = 'high';
    context.drawImage(video, 0, 0, width, height);

    return {
      id: `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
      time: video.currentTime,
      label,
      imageDataUrl: canvas.toDataURL('image/png'),
    };
  }

  function appendCaptureItem(panel, item, message) {
    state.captureItems.push(item);
    state.captureLastShotAt = item.time;
    setCaptureStatus(panel, message || `已采集 ${state.captureItems.length} 张画面，可继续播放后导出文档。`);
    renderCaptureList(panel);
  }

  function captureCurrentFrame(panel, label) {
    const item = createCaptureSnapshot(label);
    appendCaptureItem(panel, item);
  }

  function captureCueFrame(panel, cueIndex) {
    if (!state.video || cueIndex < 0 || cueIndex >= state.cues.length) {
      return;
    }

    if (state.captureSeenCueIndices.has(cueIndex)) {
      return;
    }

    const cue = state.cues[cueIndex];
    captureCurrentFrame(panel, cue.content);
    state.captureSeenCueIndices.add(cueIndex);
    setCaptureStatus(panel, `正在按字幕逐句采集，已记录 ${state.captureSeenCueIndices.size} / ${state.cues.length} 句。`);
  }

  function scheduleSpeechCapture(panel) {
    if (!state.video) {
      return;
    }

    try {
      const snapshot = createCaptureSnapshot('speech');
      if (!state.captureSegmentSnapshot || state.video.currentTime - state.captureLastShotAt >= CAPTURE_MIN_GAP_SECONDS) {
        state.captureSegmentSnapshot = snapshot;
      } else {
        state.captureSegmentSnapshot = snapshot;
      }
      state.captureSegmentPeakTime = snapshot.time;
    } catch (error) {
      setCaptureStatus(panel, error instanceof Error ? error.message : '自动截图失败');
      stopAutoCapture(panel);
    }
  }

  function finalizeAudioSegment(panel) {
    if (!state.captureVoiceActive || !state.captureSegmentSnapshot) {
      return;
    }

    const segmentEndTime = state.video ? state.video.currentTime : state.captureSegmentPeakTime;
    const duration = segmentEndTime - state.captureSegmentStartTime;

    if (duration >= CAPTURE_SEGMENT_MIN_SECONDS) {
      const representativeSnapshot = state.captureSegmentTailSnapshot
        && state.captureSegmentTailSnapshot.time > state.captureSegmentSnapshot.time
        && duration <= 1.2
        ? state.captureSegmentTailSnapshot
        : state.captureSegmentSnapshot;

      const item = {
        ...representativeSnapshot,
        label: `语音段 ${formatTime(state.captureSegmentStartTime)} - ${formatTime(segmentEndTime)}`,
      };
      appendCaptureItem(
        panel,
        item,
        `已按语音段采集 ${state.captureItems.length + 1} 张画面，当前段落 ${formatTime(state.captureSegmentStartTime)} - ${formatTime(segmentEndTime)}。`
      );
    }

    state.captureSegmentStartTime = -1;
    state.captureSegmentPeakTime = -1;
    state.captureSegmentPeakLevel = 0;
    state.captureSegmentSnapshot = null;
    state.captureSegmentTailSnapshot = null;
  }

  function runCaptureLoop(panel) {
    if (!state.captureRunning) {
      return;
    }

    const currentTime = state.video ? state.video.currentTime : 0;

    // Capture first 3 pre-roll frames in the first 150ms of video
    if (state.captureMode === 'audio' && state.capturePreRollSamples < 3 && currentTime <= 0.15) {
      try {
        captureCurrentFrame(panel, 'pre-roll');
        state.capturePreRollSamples += 1;
      } catch (err) {
        console.debug('pre-roll frame capture error', err);
      }
    }

    if (state.captureMode === 'cue') {
      const cueIndex = state.video ? getActiveCueIndex(state.video.currentTime) : -1;

      if (cueIndex >= 0 && cueIndex !== state.capturePendingCueIndex && !state.captureSeenCueIndices.has(cueIndex)) {
        state.capturePendingCueIndex = cueIndex;
        captureCueFrame(panel, cueIndex);
      }

      return;
    }

    const level = sampleCaptureLevel();

    if (level >= CAPTURE_AUDIO_THRESHOLD) {
      state.captureSilenceFrames = 0;
      state.captureSilenceStartedAt = -1;
      if (!state.captureVoiceActive) {
        state.captureVoiceActive = true;
        state.captureSegmentStartTime = currentTime;
        state.captureSegmentPeakLevel = level;
        // On first voice segment, capture a few extra frames around the speech boundary
        if (!state.captureFirstVoiceSegmentStarted) {
          state.captureFirstVoiceSegmentStarted = true;
          try {
            // Sample 3 frames before and at speech start to capture pre-speech transitions
            for (let i = 0; i < 3; i++) {
              captureCurrentFrame(panel, 'pre-speech');
            }
          } catch (err) {
            console.debug('pre-speech frame capture error', err);
          }
        }
        scheduleSpeechCapture(panel);
      } else if (level >= state.captureSegmentPeakLevel) {
        state.captureSegmentPeakLevel = level;
        scheduleSpeechCapture(panel);
      }

      if (state.captureVoiceActive && currentTime - state.captureSegmentStartTime >= CAPTURE_MAX_SEGMENT_SECONDS) {
        finalizeAudioSegment(panel);
        state.captureVoiceActive = true;
        state.captureSegmentStartTime = currentTime;
        state.captureSegmentPeakLevel = level;
        state.captureSilenceStartedAt = -1;
        state.captureSegmentTailSnapshot = null;
        scheduleSpeechCapture(panel);
      }
    } else if (state.captureVoiceActive) {
      state.captureSilenceFrames += 1;
      if (state.captureSilenceStartedAt < 0) {
        state.captureSilenceStartedAt = currentTime;
        try {
          state.captureSegmentTailSnapshot = createCaptureSnapshot('speech-tail');
        } catch (_error) {
          state.captureSegmentTailSnapshot = null;
        }
      }
      if (currentTime - state.captureSilenceStartedAt >= CAPTURE_SILENCE_HOLD_SECONDS) {
        finalizeAudioSegment(panel);
        state.captureVoiceActive = false;
        state.captureSilenceStartedAt = -1;
      }
    }

  }

  async function startAutoCapture(panel) {
    if (state.captureRunning) {
      return;
    }

    const video = state.video || document.querySelector('video');
    if (!video) {
      setCaptureStatus(panel, '页面还没加载出视频，暂时不能采集。');
      updateCaptureButtons(panel);
      return;
    }

    state.captureSeenCueIndices = new Set();
    state.capturePendingCueIndex = -1;

    if (state.cues.length > 0) {
      state.captureRunning = true;
      state.captureMode = 'cue';
      setCaptureStatus(panel, `正在按字幕逐句采集，共 ${state.cues.length} 句。请从开头播放，脚本会每句截一张图。`);
      updateCaptureButtons(panel);
      startCaptureLoop(panel);
      return;
    }

    try {
      await ensureCaptureAudio(video);
      state.captureRunning = true;
      state.captureMode = 'audio';
      state.captureVoiceActive = false;
      state.captureSilenceFrames = CAPTURE_SILENCE_FRAMES;
      state.captureSilenceStartedAt = -1;
      state.captureSegmentStartTime = -1;
      state.captureSegmentPeakTime = -1;
      state.captureSegmentPeakLevel = 0;
      state.captureSegmentSnapshot = null;
      state.captureSegmentTailSnapshot = null;
      state.video = video;
      // Reset video to start position so we can capture pre-roll frames
      video.currentTime = 0;
      state.capturePreRollSamples = 0;
      state.captureFirstVoiceSegmentStarted = false;
      setCaptureStatus(panel, `正在按语音段采集，脚本会在每段语音结束后选一张代表画面。当前已记录 ${state.captureItems.length} 张。`);
      updateCaptureButtons(panel);
      startCaptureLoop(panel);
    } catch (error) {
      setCaptureStatus(panel, error instanceof Error ? error.message : '自动采集启动失败');
      updateCaptureButtons(panel);
    }
  }

  function stopAutoCapture(panel) {
    if (!state.captureRunning) {
      setCaptureStatus(panel, state.captureItems.length > 0 ? `已暂停采集，当前有 ${state.captureItems.length} 张画面。` : '采集尚未开始。');
      updateCaptureButtons(panel);
      return;
    }

    stopCaptureLoop();
    if (state.captureMode === 'audio' && state.captureVoiceActive) {
      finalizeAudioSegment(panel);
    }
    state.captureRunning = false;
    state.captureVoiceActive = false;
    state.captureSilenceFrames = 0;
    state.captureSilenceStartedAt = -1;
    state.capturePendingCueIndex = -1;
    state.captureSegmentStartTime = -1;
    state.captureSegmentPeakTime = -1;
    state.captureSegmentPeakLevel = 0;
    state.captureSegmentSnapshot = null;
    state.captureSegmentTailSnapshot = null;
    setCaptureStatus(
      panel,
      state.captureItems.length > 0
        ? state.captureMode === 'cue'
          ? `已暂停逐句采集，当前有 ${state.captureItems.length} 张画面。`
          : `已暂停采集，当前有 ${state.captureItems.length} 张画面。`
        : '已暂停采集。'
    );
    updateCaptureButtons(panel);
  }

  function exportCaptureDocument(panel) {
    if (state.captureItems.length === 0) {
      setCaptureStatus(panel, '还没有可导出的画面。');
      updateCaptureButtons(panel);
      return;
    }

    const title = getVideoTitle();
    const entries = state.captureItems.map((item, index) => `
      <section class="card">
        <img src="${item.imageDataUrl}" alt="${escapeHtml(`画面 ${index + 1}`)}" />
      </section>
    `).join('');

    const html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${escapeHtml(title)} - 画面采集文档</title>
  <style>
    body { margin: 0; padding: 32px; background: #f4efe7; color: #2c2017; font-family: "PingFang SC", "Noto Serif SC", serif; }
    h1 { margin: 0 0 8px; font-size: 30px; }
    p { margin: 0 0 24px; color: #6f5a46; }
    .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 20px; }
    .card { background: rgba(255,255,255,0.9); border-radius: 18px; padding: 14px; box-shadow: 0 12px 28px rgba(35, 25, 17, 0.08); }
    img { display: block; width: 100%; border-radius: 12px; }
  </style>
</head>
<body>
  <h1>${escapeHtml(title)}</h1>
  <p>来源：${escapeHtml(location.href)} | 共 ${state.captureItems.length} 张画面</p>
  <div class="grid">${entries}</div>
</body>
</html>`;

    const filename = `${sanitizeFileName(title)}-画面采集.html`;
    downloadBlob(filename, html, 'text/html;charset=utf-8');
    setCaptureStatus(panel, `已导出 HTML 备份，共 ${state.captureItems.length} 张画面。`);
    updateCaptureButtons(panel);
  }

  async function exportCaptureSlides(panel) {
    if (state.captureItems.length === 0) {
      setCaptureStatus(panel, '还没有可导出的画面。');
      updateCaptureButtons(panel);
      return;
    }

    let PptxGenJS = getPptxConstructor();
    if (typeof PptxGenJS !== 'function') {
      setCaptureStatus(panel, '正在加载 PPTX 导出库，请稍候...');
      updateCaptureButtons(panel);

      try {
        PptxGenJS = await ensurePptxLibrary();
      } catch (error) {
        setCaptureStatus(
          panel,
          error instanceof Error
            ? `PPTX 导出库加载失败: ${error.message}`
            : 'PPTX 导出库加载失败'
        );
        updateCaptureButtons(panel);
        return;
      }
    }

    const title = getVideoTitle();
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = 'GitHub Copilot';
    pptx.company = 'GitHub Copilot';
    pptx.subject = 'Bilibili storybook capture';
    pptx.title = `${title} - 画面采集`;
    pptx.lang = 'zh-CN';
    pptx.theme = {
      headFontFace: 'Aptos Display',
      bodyFontFace: 'Aptos',
      lang: 'zh-CN',
    };

    state.captureItems.forEach((item, index) => {
      const slide = pptx.addSlide();
      slide.background = { color: 'FFFFFF' };
      slide.addImage({
        data: item.imageDataUrl,
        x: 0,
        y: 0,
        w: 13.333,
        h: 7.5,
        sizing: 'contain',
      });
    });

    const filename = `${sanitizeFileName(title)}-画面采集.pptx`;

    try {
      setCaptureStatus(panel, `正在生成 PPTX，共 ${state.captureItems.length} 张画面，请稍候...`);
      const blob = await pptx.write({ outputType: 'blob' });
      downloadExistingBlob(filename, blob);
      setCaptureStatus(panel, `已导出 PPTX，共 ${state.captureItems.length} 张画面，可直接导入 Google Slides。`);
    } catch (error) {
      setCaptureStatus(panel, error instanceof Error ? `PPTX 导出失败: ${error.message}` : 'PPTX 导出失败');
    }

    updateCaptureButtons(panel);
  }

  function clearCaptureItems(panel) {
    state.captureItems = [];
    state.captureLastShotAt = -Infinity;
    setCaptureStatus(panel, '已清空画面采集结果。');
    renderCaptureList(panel);
  }

  function buildPanel() {
    const existing = document.getElementById(PANEL_ID);
    if (existing) {
      updateCaptureButtons(existing);
      return existing;
    }

    const panel = document.createElement('aside');
    panel.id = PANEL_ID;

    const header = document.createElement('div');
    header.className = 'bili-transcript-header';

    const title = document.createElement('h1');
    title.className = 'bili-transcript-title';
    title.textContent = getVideoTitle();
    header.appendChild(title);

    const toggle = createButton('收起', () => {
      panel.classList.toggle('is-collapsed');
      toggle.textContent = panel.classList.contains('is-collapsed') ? '展开' : '收起';
      savePanelState(panel);
    });
    header.appendChild(toggle);

    const body = document.createElement('div');
    body.className = 'bili-transcript-body';

    const current = document.createElement('div');
    current.className = 'bili-transcript-current is-empty';
    current.textContent = '正在加载字幕...';

    const controls = document.createElement('div');
    controls.className = 'bili-transcript-controls';

    const select = document.createElement('select');
    select.className = 'bili-transcript-select';
    select.setAttribute('aria-label', '选择字幕轨道');
    controls.appendChild(select);

    const status = document.createElement('div');
    status.className = 'bili-transcript-status';
    status.textContent = '正在读取页面字幕信息...';

    const list = document.createElement('div');
    list.className = 'bili-transcript-list';

    const captureSection = document.createElement('section');
    captureSection.className = 'bili-capture-section';

    const captureHeading = document.createElement('h2');
    captureHeading.className = 'bili-capture-heading';
    captureHeading.textContent = '绘本采集';

    const captureControls = document.createElement('div');
    captureControls.className = 'bili-transcript-capture-controls';

    const startCaptureButton = createButton('开始采集', () => {
      startAutoCapture(panel);
    });
    startCaptureButton.dataset.captureAction = 'start';

    const stopCaptureButton = createButton('停止', () => {
      stopAutoCapture(panel);
    });
    stopCaptureButton.dataset.captureAction = 'stop';

    const snapshotButton = createButton('手动截取', () => {
      try {
        captureCurrentFrame(panel, 'manual');
      } catch (error) {
        setCaptureStatus(panel, error instanceof Error ? error.message : '手动截图失败');
        updateCaptureButtons(panel);
      }
    });
    snapshotButton.dataset.captureAction = 'snapshot';

    const exportButton = createButton('导出PPTX', () => {
      exportCaptureSlides(panel);
    });
    exportButton.dataset.captureAction = 'export';

    const exportHtmlButton = createButton('导出HTML', () => {
      exportCaptureDocument(panel);
    });

    const clearButton = createButton('清空', () => {
      clearCaptureItems(panel);
    });
    clearButton.dataset.captureAction = 'clear';

    captureControls.appendChild(startCaptureButton);
    captureControls.appendChild(stopCaptureButton);
    captureControls.appendChild(snapshotButton);
    captureControls.appendChild(exportButton);
    captureControls.appendChild(exportHtmlButton);
    captureControls.appendChild(clearButton);

    const captureStatus = document.createElement('div');
    captureStatus.className = 'bili-capture-status';
    captureStatus.textContent = state.captureStatus;

    const captureList = document.createElement('div');
    captureList.className = 'bili-capture-list';

    body.appendChild(current);
    body.appendChild(controls);
    body.appendChild(status);
    body.appendChild(list);
    captureSection.appendChild(captureHeading);
    captureSection.appendChild(captureControls);
    captureSection.appendChild(captureStatus);
    captureSection.appendChild(captureList);
    body.appendChild(captureSection);

    panel.appendChild(header);
    panel.appendChild(body);

    const persistedState = readPanelState();
    if (persistedState.collapsed) {
      panel.classList.add('is-collapsed');
      toggle.textContent = '展开';
    }

    document.body.appendChild(panel);
    renderCaptureList(panel);
    updateCaptureButtons(panel);
    return panel;
  }

  function updatePanelTitle(panel) {
    const title = panel.querySelector('.bili-transcript-title');
    if (title) {
      title.textContent = getVideoTitle();
    }
  }

  function setCurrentText(panel, text) {
    const current = panel.querySelector('.bili-transcript-current');
    current.textContent = text || '等待台词...';
    current.classList.toggle('is-empty', !text);
  }

  function renderStatus(panel, message) {
    const status = panel.querySelector('.bili-transcript-status');
    const list = panel.querySelector('.bili-transcript-list');
    status.textContent = message;
    list.innerHTML = '';
    setCurrentText(panel, '');
  }

  function getCurrentCid(initialState) {
    if (initialState.epInfo && initialState.epInfo.cid) {
      return initialState.epInfo.cid;
    }

    const url = new URL(location.href);
    const page = Number(url.searchParams.get('p') || '1');
    const pages = initialState.videoData && initialState.videoData.pages;

    if (Array.isArray(pages) && pages[page - 1] && pages[page - 1].cid) {
      return pages[page - 1].cid;
    }

    return initialState.videoData && initialState.videoData.cid ? initialState.videoData.cid : null;
  }

  function requestText(url) {
    if (typeof GM_xmlhttpRequest === 'function') {
      return new Promise((resolve, reject) => {
        GM_xmlhttpRequest({
          method: 'GET',
          url,
          anonymous: false,
          onload: (response) => {
            if (response.status < 200 || response.status >= 300) {
              reject(new Error(`请求失败: ${response.status}`));
              return;
            }

            resolve(response.responseText || '');
          },
          onerror: () => {
            reject(new Error('请求失败: failed to fetch'));
          },
        });
      });
    }

    return fetch(url, { credentials: 'include' }).then(async (response) => {
      if (!response.ok) {
        throw new Error(`请求失败: ${response.status}`);
      }

      return response.text();
    });
  }

  async function requestJson(url, failureLabel) {
    const text = await requestText(url);

    try {
      return JSON.parse(text);
    } catch (error) {
      const snippet = text.slice(0, 80).replace(/\s+/g, ' ');
      throw new Error(`${failureLabel}: 返回内容不是 JSON${snippet ? ` (${snippet})` : ''}`);
    }
  }

  async function getSubtitleTracks() {
    const playInfo = window.__playinfo__ || window.__PLAYINFO__;
    const directTracks = playInfo && playInfo.data && playInfo.data.subtitle && playInfo.data.subtitle.subtitles;

    if (Array.isArray(directTracks) && directTracks.length > 0) {
      return directTracks;
    }

    const initialState = window.__INITIAL_STATE__;
    if (!initialState) {
      return [];
    }

    const bvid = initialState.bvid || (initialState.videoData && initialState.videoData.bvid);
    const cid = getCurrentCid(initialState);

    if (!bvid || !cid) {
      return [];
    }

    const payload = await requestJson(
      `${API_ORIGIN}/x/player/v2?bvid=${encodeURIComponent(bvid)}&cid=${encodeURIComponent(cid)}`,
      '字幕接口返回异常'
    );

    return payload && payload.data && payload.data.subtitle && Array.isArray(payload.data.subtitle.subtitles)
      ? payload.data.subtitle.subtitles
      : [];
  }

  async function loadCues(trackUrl) {
    const payload = await requestJson(trackUrl, '字幕文件返回异常');
    const body = Array.isArray(payload.body) ? payload.body : [];

    return body
      .map((entry) => ({
        from: Number(entry.from || 0),
        to: Number(entry.to || 0),
        content: String(entry.content || '').trim(),
      }))
      .filter((entry) => entry.content);
  }

  function choosePreferredTrack(tracks) {
    const englishTrack = tracks.find((track, index) => /英文|英语|english|en/i.test(getTrackLabel(track, index)));
    if (englishTrack) {
      return englishTrack;
    }

    const chineseTrack = tracks.find((track, index) => /中文|汉语|普通话|zh/i.test(getTrackLabel(track, index)));
    if (chineseTrack) {
      return chineseTrack;
    }

    return tracks[0];
  }

  function renderCueList(panel) {
    const list = panel.querySelector('.bili-transcript-list');
    const status = panel.querySelector('.bili-transcript-status');
    list.innerHTML = '';

    state.cues.forEach((cue) => {
      const row = document.createElement('button');
      row.type = 'button';
      row.className = 'bili-transcript-line';

      const time = document.createElement('span');
      time.className = 'bili-transcript-time';
      time.textContent = formatTime(cue.from);

      const text = document.createElement('span');
      text.className = 'bili-transcript-text';
      text.textContent = cue.content;

      row.appendChild(time);
      row.appendChild(text);
      row.addEventListener('click', () => {
        if (!state.video) {
          return;
        }

        state.video.currentTime = cue.from;
        state.video.play().catch(() => {});
      });

      list.appendChild(row);
    });

    status.textContent = `已加载 ${state.cues.length} 句台词，可点击跳转。`;
    setCurrentText(panel, state.cues[0] ? state.cues[0].content : '');
  }

  function getActiveCueIndex(currentTime) {
    return state.cues.findIndex((cue) => currentTime >= cue.from && currentTime <= cue.to);
  }

  function syncActiveCue(panel) {
    if (!state.video || state.cues.length === 0) {
      return;
    }

    const nextIndex = getActiveCueIndex(state.video.currentTime);
    if (nextIndex === state.activeCueIndex) {
      return;
    }

    state.activeCueIndex = nextIndex;
    const rows = panel.querySelectorAll('.bili-transcript-line');
    rows.forEach((row, index) => {
      const isActive = index === nextIndex;
      row.classList.toggle(ACTIVE_CLASS, isActive);
      if (isActive) {
        row.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      }
    });

    setCurrentText(panel, nextIndex >= 0 ? state.cues[nextIndex].content : '');
  }

  function attachVideoSync(panel, video) {
    detachPlaybackHandler();
    state.video = video;
    state.playbackHandler = () => syncActiveCue(panel);
    state.video.addEventListener('timeupdate', state.playbackHandler);
    syncActiveCue(panel);
  }

  function scheduleRetry() {
    clearRetryTimer();
    state.retryTimer = window.setTimeout(() => {
      initializeViewer();
    }, 3000);
  }

  async function populateTracks(panel) {
    const select = panel.querySelector('.bili-transcript-select');
    const tracks = await getSubtitleTracks();

    if (tracks.length === 0) {
      renderStatus(panel, '这个视频没有公开字幕，暂时无法显示台词。');
      return;
    }

    select.innerHTML = '';
    tracks.forEach((track, index) => {
      const option = document.createElement('option');
      option.value = getTrackUrl(track);
      option.textContent = getTrackLabel(track, index);
      select.appendChild(option);
    });

    const preferredTrack = choosePreferredTrack(tracks);
    state.selectedTrackUrl = getTrackUrl(preferredTrack);
    select.value = state.selectedTrackUrl;

    select.onchange = async () => {
      if (!select.value || select.value === state.selectedTrackUrl) {
        return;
      }

      state.selectedTrackUrl = select.value;
      renderStatus(panel, '正在切换字幕轨道...');
      state.cues = await loadCues(state.selectedTrackUrl);
      state.activeCueIndex = -1;
      renderCueList(panel);
      syncActiveCue(panel);
    };

    state.cues = await loadCues(state.selectedTrackUrl);
    state.activeCueIndex = -1;
    renderCueList(panel);
  }

  async function initializeViewer() {
    ensureStyles();
    const panel = buildPanel();
    updatePanelTitle(panel);
    clearRetryTimer();

    const video = document.querySelector('video');
    if (!video) {
      renderStatus(panel, '页面还没加载出视频，稍后自动重试。');
      scheduleRetry();
      return;
    }

    try {
      renderStatus(panel, '正在读取页面字幕信息...');
      await populateTracks(panel);
      attachVideoSync(panel, video);
    } catch (error) {
      console.error('[Bilibili Transcript Viewer]', error);
      renderStatus(panel, error instanceof Error ? error.message : '字幕加载失败');
      scheduleRetry();
    }
  }

  function watchLocationChanges() {
    if (state.routeWatcherStarted) {
      return;
    }

    state.routeWatcherStarted = true;
    window.setInterval(() => {
      if (location.href === state.lastUrl) {
        return;
      }

      state.lastUrl = location.href;
      state.cues = [];
      state.activeCueIndex = -1;
      state.selectedTrackUrl = '';
      const panel = document.getElementById(PANEL_ID);
      if (state.captureItems.length > 0) {
        stopCaptureLoop();
        state.captureRunning = false;
        state.captureVoiceActive = false;
        state.captureSilenceFrames = 0;
        state.captureSilenceStartedAt = -1;
        state.capturePendingCueIndex = -1;
        state.captureSegmentStartTime = -1;
        state.captureSegmentPeakTime = -1;
        state.captureSegmentPeakLevel = 0;
        state.captureSegmentSnapshot = null;
        state.captureSegmentTailSnapshot = null;
        state.captureMode = 'audio';
        if (panel) {
          setCaptureStatus(panel, `检测到视频已切换，已保留上一段采集的 ${state.captureItems.length} 张画面。请先导出或手动清空后再开始新一轮采集。`);
          renderCaptureList(panel);
        }
      } else {
        resetCaptureState(panel);
      }
      detachPlaybackHandler();
      initializeViewer();
    }, 1000);
  }

  function bootstrap() {
    const start = () => {
      if (!document.body || !document.head) {
        return;
      }

      initializeViewer();
      watchLocationChanges();
    };

    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', start, { once: true });
      return;
    }

    start();
  }

  bootstrap();
})();
