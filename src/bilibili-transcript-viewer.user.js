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
  const CAPTURE_SHARPNESS_SAMPLE_INTERVAL_SECONDS = 0.12;
  const CAPTURE_SHARPNESS_DOWNSCALE_WIDTH = 160;
  const CAPTURE_SHARPNESS_IMPROVEMENT_RATIO = 1.04;
  const CAPTURE_PREVIEW_ADJUST_STEP_SECONDS = 0.02;
  const CAPTURE_PREVIEW_ADJUST_COARSE_STEP_SECONDS = 0.1;
  const PPTX_LIB_URL = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
  const CAPTURE_AUDIO_BINDING_KEY = '__biliTranscriptCaptureAudioBinding';
  const CAPTURE_PREVIEW_ID = 'bili-transcript-viewer-capture-preview';

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
    captureStatus: '点击“开始自动采集”后，脚本会按字幕或语音片段自动截取画面。',
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
    captureSegmentBestSharpness: 0,
    captureSegmentLastSharpnessSampleAt: -1,
    captureCapturedPreRoll: false,
    capturePreRollSamples: 0,
    captureFirstVoiceSegmentStarted: false,
    capturePreviewIndex: -1,
    pptxLibraryPromise: null,
  };
  let capturePreviewTimeOffset = 0;
  let captureSharpnessCanvas = null;
  let captureSharpnessContext = null;

  function ensureStyles() {
    if (document.getElementById(STYLE_ID)) {
      return;
    }

    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = `
      #${PANEL_ID} {
        --bili-panel-bg: rgba(250, 246, 239, 0.96);
        --bili-panel-stroke: rgba(124, 93, 59, 0.18);
        --bili-text-main: #1e1812;
        --bili-text-subtle: #6f5a46;
        --bili-accent: #a16a35;
        position: fixed;
        top: 18px;
        right: 18px;
        z-index: 999999;
        width: 390px;
        height: auto;
        max-width: calc(100vw - 36px);
        max-height: calc(100vh - 36px);
        display: flex;
        flex-direction: column;
        border-radius: 18px;
        overflow: hidden;
        border: 1px solid var(--bili-panel-stroke);
        box-shadow: 0 22px 64px rgba(30, 24, 18, 0.24);
        background:
          radial-gradient(130% 60% at 0% 0%, rgba(245, 227, 193, 0.34), transparent 56%),
          radial-gradient(120% 70% at 100% 0%, rgba(255, 255, 255, 0.42), transparent 50%),
          var(--bili-panel-bg);
        color: var(--bili-text-main);
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
        padding: 14px 16px 13px;
        background: linear-gradient(140deg, rgba(225, 193, 149, 0.44) 0%, rgba(255, 246, 231, 0.82) 100%);
        border-bottom: 1px solid rgba(124, 93, 59, 0.16);
        cursor: move;
        user-select: none;
      }

      #${PANEL_ID} .bili-transcript-title {
        flex: 1;
        min-width: 0;
        margin: 0;
        font-size: 15px;
        line-height: 1.35;
        font-weight: 700;
        letter-spacing: 0.2px;
      }

      #${PANEL_ID} .bili-transcript-body {
        display: flex;
        flex-direction: column;
        gap: 12px;
        padding: 14px 16px 16px;
      }

      #${PANEL_ID} .bili-transcript-section {
        display: flex;
        flex-direction: column;
        gap: 8px;
      }

      #${PANEL_ID} .bili-transcript-section-heading {
        margin: 0;
        font-size: 13px;
        line-height: 1.4;
        font-weight: 700;
        color: #4d3a2c;
      }

      #${PANEL_ID} .bili-transcript-section-note {
        color: var(--bili-text-subtle);
        font-size: 12px;
        line-height: 1.5;
      }

      #${PANEL_ID}.is-collapsed .bili-transcript-body {
        display: none;
      }

      #${PANEL_ID} .bili-transcript-current {
        min-height: 76px;
        padding: 14px;
        border-radius: 14px;
        border: 1px solid rgba(124, 93, 59, 0.16);
        background: linear-gradient(180deg, rgba(255, 255, 255, 0.82), rgba(255, 248, 239, 0.99));
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.75);
        font-size: 24px;
        line-height: 1.5;
      }

      #${PANEL_ID} .bili-transcript-current.is-empty,
      #${PANEL_ID} .bili-transcript-status {
        color: var(--bili-text-subtle);
        font-size: 14px;
      }

      #${PANEL_ID} .bili-transcript-controls {
        display: flex;
        align-items: center;
        gap: 8px;
      }

      #${PANEL_ID} .bili-transcript-capture-controls {
        display: flex;
        flex-direction: column;
        gap: 8px;
      }

      #${PANEL_ID} .bili-capture-controls-row {
        display: flex;
        gap: 8px;
      }

      #${PANEL_ID} .bili-capture-controls-row button {
        flex: 1;
        min-height: 34px;
      }

      #${PANEL_ID} .bili-capture-controls-row .bili-btn-danger {
        color: #c64c3c;
        background: rgba(220, 100, 80, 0.12);
        border-color: rgba(220, 100, 80, 0.34);
      }

      #${PANEL_ID} .bili-capture-controls-row .bili-btn-danger:hover {
        background: rgba(220, 100, 80, 0.22);
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
        color: var(--bili-text-subtle);
        font-size: 13px;
        line-height: 1.5;
      }

      #${PANEL_ID} .bili-capture-list {
        display: flex;
        flex-direction: column;
        gap: 8px;
        overflow: auto;
        max-height: min(40vh, 320px);
        padding-right: 4px;
      }

      #${PANEL_ID} .bili-capture-card {
        display: grid;
        grid-template-columns: 108px 1fr;
        gap: 10px;
        width: 100%;
        padding: 8px;
        border: 1px solid rgba(124, 93, 59, 0.14);
        border-radius: 12px;
        background: rgba(255,255,255,0.82);
        text-align: left;
        transition: transform 0.16s ease, box-shadow 0.16s ease, border-color 0.16s ease;
      }

      #${PANEL_ID} .bili-capture-card:hover {
        border-color: rgba(161, 106, 53, 0.34);
        transform: translateY(-1px);
        box-shadow: 0 6px 16px rgba(64, 45, 29, 0.12);
      }

      #${PANEL_ID} .bili-capture-thumb {
        width: 108px;
        height: 60px;
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

      #${PANEL_ID} .bili-capture-item-container {
        display: grid;
        grid-template-columns: 1fr 32px;
        gap: 6px;
        align-items: stretch;
      }

      #${PANEL_ID} .bili-capture-delete-btn {
        padding: 0;
        background: rgba(220, 100, 80, 0.12);
        border: 1px solid rgba(220, 100, 80, 0.28);
        border-radius: 8px;
        color: #c64c3c;
        font-size: 18px;
        font-weight: bold;
        cursor: pointer;
        transition: all 0.2s ease;
        display: flex;
        align-items: center;
        justify-content: center;
      }

      #${PANEL_ID} .bili-capture-delete-btn:hover {
        background: rgba(220, 100, 80, 0.22);
        border-color: rgba(220, 100, 80, 0.48);
      }

      #${PANEL_ID} .bili-capture-delete-btn:active {
        background: rgba(220, 100, 80, 0.32);
      }

      #${CAPTURE_PREVIEW_ID} {
        position: fixed;
        inset: 0;
        z-index: 1000001;
        display: none;
        align-items: center;
        justify-content: center;
        background: rgba(12, 10, 8, 0.74);
        padding: 20px;
      }

      #${CAPTURE_PREVIEW_ID}.is-open {
        display: flex;
      }

      #${CAPTURE_PREVIEW_ID} .bili-capture-preview-dialog {
        width: min(1200px, calc(100vw - 40px));
        max-height: calc(100vh - 40px);
        height: min(920px, calc(100vh - 40px));
        display: flex;
        flex-direction: column;
        border-radius: 16px;
        border: 1px solid rgba(255, 255, 255, 0.18);
        background: rgba(28, 22, 17, 0.96);
        box-shadow: 0 24px 64px rgba(0, 0, 0, 0.36);
        overflow: hidden;
      }

      #${CAPTURE_PREVIEW_ID} .bili-capture-preview-toolbar {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 10px 12px;
        border-bottom: 1px solid rgba(255, 255, 255, 0.12);
      }

      #${CAPTURE_PREVIEW_ID} .bili-capture-preview-info {
        flex: 1;
        min-width: 0;
        color: #f6e8d8;
        font-size: 13px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }

      #${CAPTURE_PREVIEW_ID} .bili-capture-preview-btn {
        border: 1px solid rgba(255, 255, 255, 0.24);
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.1);
        color: #f4ede6;
        padding: 6px 12px;
        cursor: pointer;
        font-size: 12px;
      }

      #${CAPTURE_PREVIEW_ID} .bili-capture-preview-btn:hover {
        background: rgba(255, 255, 255, 0.18);
      }

      #${CAPTURE_PREVIEW_ID} .bili-capture-preview-btn:disabled {
        opacity: 0.45;
        cursor: not-allowed;
      }

      #${CAPTURE_PREVIEW_ID} .bili-capture-preview-btn.is-danger {
        border-color: rgba(246, 120, 104, 0.62);
        color: #ffd8d3;
        background: rgba(246, 120, 104, 0.16);
      }

      #${CAPTURE_PREVIEW_ID} .bili-capture-preview-stage {
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 18px;
        min-height: 0;
        flex: 1;
      }

      #${CAPTURE_PREVIEW_ID} .bili-capture-preview-image {
        max-width: 100%;
        max-height: calc(100vh - 110px);
        border-radius: 10px;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.32);
        background: rgba(255, 255, 255, 0.08);
      }

      #${PANEL_ID}::after {
        content: '';
        position: absolute;
        bottom: 0;
        right: 0;
        width: 16px;
        height: 16px;
        cursor: se-resize;
        background: linear-gradient(135deg, transparent 50%, rgba(124, 93, 59, 0.3) 50%);
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
        border: 1px solid rgba(124, 93, 59, 0.2);
        border-radius: 999px;
        background: linear-gradient(180deg, rgba(255, 255, 255, 0.82), rgba(241, 230, 214, 0.62));
        color: #382719;
        padding: 8px 11px;
        cursor: pointer;
        font-size: 12px;
        font-weight: 600;
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.72);
        transition: background 0.16s ease, border-color 0.16s ease, transform 0.1s ease;
      }

      #${PANEL_ID} button:hover {
        background: linear-gradient(180deg, rgba(255, 255, 255, 0.92), rgba(242, 228, 207, 0.88));
        border-color: rgba(161, 106, 53, 0.35);
      }

      #${PANEL_ID} button:disabled {
        opacity: 0.45;
        cursor: not-allowed;
        transform: none;
      }

      #${PANEL_ID} .bili-btn-primary {
        background: linear-gradient(180deg, #f6e6cb, #edd5ab);
        border-color: rgba(161, 106, 53, 0.46);
      }

      #${PANEL_ID} .bili-btn-primary:hover {
        background: linear-gradient(180deg, #f8ebd5, #f0d9b2);
      }

      #${PANEL_ID} .bili-btn-stop {
        background: linear-gradient(180deg, #f8ddcf, #efc6b3);
        border-color: rgba(200, 98, 74, 0.44);
        color: #5a261c;
      }

      #${PANEL_ID} .bili-btn-capture {
        background: linear-gradient(180deg, #e8f0f8, #d4e4f3);
        border-color: rgba(85, 120, 160, 0.36);
        color: #2a3f56;
      }

      #${PANEL_ID} .bili-btn-export,
      #${PANEL_ID} .bili-btn-export-html {
        background: linear-gradient(180deg, #e7f2e6, #d8ead6);
        border-color: rgba(88, 132, 82, 0.36);
        color: #2e4f2c;
      }

      #${PANEL_ID} .bili-capture-list::-webkit-scrollbar,
      #${PANEL_ID} .bili-transcript-list::-webkit-scrollbar {
        width: 8px;
      }

      #${PANEL_ID} .bili-capture-list::-webkit-scrollbar-thumb,
      #${PANEL_ID} .bili-transcript-list::-webkit-scrollbar-thumb {
        border-radius: 999px;
        background: rgba(124, 93, 59, 0.26);
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

        #${PANEL_ID} .bili-capture-card {
          grid-template-columns: 92px 1fr;
        }

        #${PANEL_ID} .bili-capture-thumb {
          width: 92px;
          height: 52px;
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
    const left = typeof panel.style.left === 'string' && panel.style.left.endsWith('px') ? panel.style.left : null;
    const top = typeof panel.style.top === 'string' && panel.style.top.endsWith('px') ? panel.style.top : null;
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        collapsed: panel.classList.contains('is-collapsed'),
        left,
        top,
      })
    );
  }

  function applyPanelPosition(panel, persistedState) {
    if (!persistedState || typeof persistedState.left !== 'string' || typeof persistedState.top !== 'string') {
      return;
    }

    panel.style.left = persistedState.left;
    panel.style.top = persistedState.top;
    panel.style.right = 'auto';
    panel.style.bottom = 'auto';
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
      empty.textContent = '还没有采集到画面。播放视频后点击“开始自动采集”，脚本会按字幕或语音片段自动截取。';
      list.appendChild(empty);
      updateCaptureButtons(panel);
      return;
    }

    state.captureItems.forEach((item, index) => {
      const container = document.createElement('div');
      container.className = 'bili-capture-item-container';

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
        openCapturePreview(panel, index);
      });

      const deleteBtn = document.createElement('button');
      deleteBtn.type = 'button';
      deleteBtn.className = 'bili-capture-delete-btn';
      deleteBtn.textContent = '×';
      deleteBtn.title = '删除此画面';
      deleteBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        deleteCaptureItemAt(panel, index);
      });

      container.appendChild(row);
      container.appendChild(deleteBtn);
      list.appendChild(container);
    });

    updateCaptureButtons(panel);
  }

  function ensureCapturePreviewOverlay(panel) {
    const existing = document.getElementById(CAPTURE_PREVIEW_ID);
    if (existing) {
      return existing;
    }

    const overlay = document.createElement('div');
    overlay.id = CAPTURE_PREVIEW_ID;

    const dialog = document.createElement('div');
    dialog.className = 'bili-capture-preview-dialog';

    const toolbar = document.createElement('div');
    toolbar.className = 'bili-capture-preview-toolbar';

    const info = document.createElement('div');
    info.className = 'bili-capture-preview-info';

    const prevButton = document.createElement('button');
    prevButton.type = 'button';
    prevButton.className = 'bili-capture-preview-btn';
    prevButton.dataset.previewAction = 'prev';
    prevButton.textContent = '上一张';

    const nextButton = document.createElement('button');
    nextButton.type = 'button';
    nextButton.className = 'bili-capture-preview-btn';
    nextButton.dataset.previewAction = 'next';
    nextButton.textContent = '下一张';

    const adjustBackButton = document.createElement('button');
    adjustBackButton.type = 'button';
    adjustBackButton.className = 'bili-capture-preview-btn';
    adjustBackButton.dataset.previewAction = 'adjust-back';
    adjustBackButton.textContent = '← 0.02s';
    adjustBackButton.title = '向前调整 0.02 秒（按住 Shift 为 0.1 秒）';

    const adjustForwardButton = document.createElement('button');
    adjustForwardButton.type = 'button';
    adjustForwardButton.className = 'bili-capture-preview-btn';
    adjustForwardButton.dataset.previewAction = 'adjust-forward';
    adjustForwardButton.textContent = '0.02s →';
    adjustForwardButton.title = '向后调整 0.02 秒（按住 Shift 为 0.1 秒）';

    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'bili-capture-preview-btn is-danger';
    deleteButton.dataset.previewAction = 'delete';
    deleteButton.textContent = '删除';

    const closeButton = document.createElement('button');
    closeButton.type = 'button';
    closeButton.className = 'bili-capture-preview-btn';
    closeButton.dataset.previewAction = 'close';
    closeButton.textContent = '关闭';

    toolbar.appendChild(info);
    toolbar.appendChild(prevButton);
    toolbar.appendChild(nextButton);
    toolbar.appendChild(adjustBackButton);
    toolbar.appendChild(adjustForwardButton);
    toolbar.appendChild(deleteButton);
    toolbar.appendChild(closeButton);

    const stage = document.createElement('div');
    stage.className = 'bili-capture-preview-stage';

    const image = document.createElement('img');
    image.className = 'bili-capture-preview-image';
    image.alt = '预览画面';
    stage.appendChild(image);

    dialog.appendChild(toolbar);
    dialog.appendChild(stage);
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    overlay.addEventListener('click', (event) => {
      if (event.target === overlay) {
        closeCapturePreview();
      }
    });

    prevButton.addEventListener('click', () => {
      if (state.capturePreviewIndex > 0) {
        state.capturePreviewIndex -= 1;
        capturePreviewTimeOffset = 0;
        syncCapturePreview();
      }
    });

    nextButton.addEventListener('click', () => {
      if (state.capturePreviewIndex < state.captureItems.length - 1) {
        state.capturePreviewIndex += 1;
        capturePreviewTimeOffset = 0;
        syncCapturePreview();
      }
    });

    adjustBackButton.addEventListener('click', (event) => {
      const step = event.shiftKey ? CAPTURE_PREVIEW_ADJUST_COARSE_STEP_SECONDS : CAPTURE_PREVIEW_ADJUST_STEP_SECONDS;
      capturePreviewTimeOffset = Math.round((capturePreviewTimeOffset - step) * 1000) / 1000;
      syncCapturePreview();
    });

    adjustForwardButton.addEventListener('click', (event) => {
      const step = event.shiftKey ? CAPTURE_PREVIEW_ADJUST_COARSE_STEP_SECONDS : CAPTURE_PREVIEW_ADJUST_STEP_SECONDS;
      capturePreviewTimeOffset = Math.round((capturePreviewTimeOffset + step) * 1000) / 1000;
      syncCapturePreview();
    });

    deleteButton.addEventListener('click', () => {
      const panelElement = panel || document.getElementById(PANEL_ID);
      if (panelElement) {
        deleteCaptureItemAt(panelElement, state.capturePreviewIndex);
      }
    });

    closeButton.addEventListener('click', () => {
      closeCapturePreview();
    });

    document.addEventListener('keydown', (event) => {
      if (!overlay.classList.contains('is-open')) {
        return;
      }

      if (event.key === 'Escape') {
        closeCapturePreview();
        return;
      }

      if (event.key === 'ArrowLeft') {
        if (state.capturePreviewIndex > 0) {
          state.capturePreviewIndex -= 1;
          syncCapturePreview();
        }
        return;
      }

      if (event.key === 'ArrowRight') {
        if (state.capturePreviewIndex < state.captureItems.length - 1) {
          state.capturePreviewIndex += 1;
          syncCapturePreview();
        }
        return;
      }

      if (event.key === 'Delete' || event.key === 'Backspace') {
        const panelElement = panel || document.getElementById(PANEL_ID);
        if (panelElement) {
          deleteCaptureItemAt(panelElement, state.capturePreviewIndex);
        }
      }
    });

    return overlay;
  }

  function syncCapturePreview() {
    const overlay = document.getElementById(CAPTURE_PREVIEW_ID);
    if (!overlay) {
      return;
    }

    if (state.capturePreviewIndex < 0 || state.capturePreviewIndex >= state.captureItems.length) {
      closeCapturePreview();
      return;
    }

    const item = state.captureItems[state.capturePreviewIndex];
    const info = overlay.querySelector('.bili-capture-preview-info');
    const image = overlay.querySelector('.bili-capture-preview-image');
    const prevButton = overlay.querySelector('[data-preview-action="prev"]');
    const nextButton = overlay.querySelector('[data-preview-action="next"]');
    const adjustBackButton = overlay.querySelector('[data-preview-action="adjust-back"]');
    const adjustForwardButton = overlay.querySelector('[data-preview-action="adjust-forward"]');

    if (adjustBackButton && adjustForwardButton) {
      adjustBackButton.disabled = !state.video;
      adjustForwardButton.disabled = !state.video;
    }

    if (image) {
      if (capturePreviewTimeOffset !== 0 && state.video) {
        const adjustedTime = Math.max(0, item.time + capturePreviewTimeOffset);
        const originalTime = state.video.currentTime;
        state.video.currentTime = adjustedTime;
        try {
          const adjustedSnapshot = createCaptureSnapshot(`preview-adjust`);
          image.src = adjustedSnapshot.imageDataUrl;
        } catch (err) {
          console.debug('preview adjustment screenshot failed', err);
          image.src = item.imageDataUrl;
        }
        state.video.currentTime = originalTime;
      } else {
        image.src = item.imageDataUrl;
        capturePreviewTimeOffset = 0;
      }
      image.alt = `画面 ${state.capturePreviewIndex + 1}`;
    }

    if (info) {
      const label = item.label || '未命名';
      const offsetText = capturePreviewTimeOffset !== 0 ? ` (${capturePreviewTimeOffset > 0 ? '+' : ''}${capturePreviewTimeOffset.toFixed(2)}s)` : '';
      info.textContent = `${state.capturePreviewIndex + 1}/${state.captureItems.length} · ${formatTime(item.time)}${offsetText} · ${label}`;
    }

    if (prevButton) {
      prevButton.disabled = state.capturePreviewIndex <= 0;
    }
    if (nextButton) {
      nextButton.disabled = state.capturePreviewIndex >= state.captureItems.length - 1;
    }
  }

  function openCapturePreview(panel, index) {
    if (index < 0 || index >= state.captureItems.length) {
      return;
    }

    const overlay = ensureCapturePreviewOverlay(panel);
    state.capturePreviewIndex = index;
    overlay.classList.add('is-open');
    syncCapturePreview();
  }

  function closeCapturePreview() {
    const overlay = document.getElementById(CAPTURE_PREVIEW_ID);
    if (overlay) {
      overlay.classList.remove('is-open');
    }
    state.capturePreviewIndex = -1;
    capturePreviewTimeOffset = 0;
  }

  function deleteCaptureItemAt(panel, index) {
    if (index < 0 || index >= state.captureItems.length) {
      return;
    }

    state.captureItems.splice(index, 1);

    if (state.captureItems.length === 0) {
      closeCapturePreview();
    } else if (state.capturePreviewIndex >= state.captureItems.length) {
      state.capturePreviewIndex = state.captureItems.length - 1;
    }

    renderCaptureList(panel);
    setCaptureStatus(panel, `已删除一张画面，当前有 ${state.captureItems.length} 张。`);

    if (state.capturePreviewIndex >= 0) {
      syncCapturePreview();
    }
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
    state.captureStatus = '点击“开始自动采集”后，脚本会按字幕或语音片段自动截取画面。';
    state.captureMode = 'audio';
    state.captureSeenCueIndices = new Set();
    state.capturePendingCueIndex = -1;
    state.captureSegmentStartTime = -1;
    state.captureSegmentPeakTime = -1;
    state.captureSegmentPeakLevel = 0;
    state.captureSegmentSnapshot = null;
    state.captureSegmentTailSnapshot = null;
    state.captureSegmentBestSharpness = 0;
    state.captureSegmentLastSharpnessSampleAt = -1;
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

  function computeCurrentFrameSharpness(video) {
    if (!video || !video.videoWidth || !video.videoHeight) {
      return 0;
    }

    const width = CAPTURE_SHARPNESS_DOWNSCALE_WIDTH;
    const height = Math.max(90, Math.round(width * (video.videoHeight / video.videoWidth)));

    if (!captureSharpnessCanvas) {
      captureSharpnessCanvas = document.createElement('canvas');
      captureSharpnessContext = captureSharpnessCanvas.getContext('2d', { willReadFrequently: true });
    }

    if (!captureSharpnessContext) {
      return 0;
    }

    if (captureSharpnessCanvas.width !== width || captureSharpnessCanvas.height !== height) {
      captureSharpnessCanvas.width = width;
      captureSharpnessCanvas.height = height;
    }

    captureSharpnessContext.drawImage(video, 0, 0, width, height);
    const data = captureSharpnessContext.getImageData(0, 0, width, height).data;

    // Fast edge-energy estimate (sample every other pixel) to approximate frame sharpness.
    let score = 0;
    let samples = 0;
    for (let y = 2; y < height; y += 2) {
      for (let x = 2; x < width; x += 2) {
        const idx = (y * width + x) * 4;
        const leftIdx = idx - 8;
        const upIdx = idx - width * 4 * 2;

        const luma = (data[idx] * 77 + data[idx + 1] * 150 + data[idx + 2] * 29) >> 8;
        const lumaLeft = (data[leftIdx] * 77 + data[leftIdx + 1] * 150 + data[leftIdx + 2] * 29) >> 8;
        const lumaUp = (data[upIdx] * 77 + data[upIdx + 1] * 150 + data[upIdx + 2] * 29) >> 8;

        score += Math.abs(luma - lumaLeft) + Math.abs(luma - lumaUp);
        samples += 1;
      }
    }

    return samples > 0 ? score / samples : 0;
  }

  function updateSpeechSegmentSnapshot(panel, currentTime, forceCapture) {
    if (!state.video) {
      return;
    }

    const shouldSample = forceCapture
      || state.captureSegmentLastSharpnessSampleAt < 0
      || currentTime - state.captureSegmentLastSharpnessSampleAt >= CAPTURE_SHARPNESS_SAMPLE_INTERVAL_SECONDS;
    if (!shouldSample) {
      return;
    }

    state.captureSegmentLastSharpnessSampleAt = currentTime;
    const sharpness = computeCurrentFrameSharpness(state.video);
    const needsUpdate = !state.captureSegmentSnapshot
      || forceCapture
      || sharpness > state.captureSegmentBestSharpness * CAPTURE_SHARPNESS_IMPROVEMENT_RATIO;
    if (!needsUpdate) {
      return;
    }

    const snapshot = createCaptureSnapshot('speech');
    state.captureSegmentSnapshot = snapshot;
    state.captureSegmentPeakTime = snapshot.time;
    state.captureSegmentBestSharpness = sharpness;
  }

  function scheduleSpeechCapture(panel, currentTime, forceCapture = false) {
    if (!state.video) {
      return;
    }

    try {
      updateSpeechSegmentSnapshot(panel, currentTime, forceCapture);
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
    state.captureSegmentBestSharpness = 0;
    state.captureSegmentLastSharpnessSampleAt = -1;
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
        state.captureSegmentBestSharpness = 0;
        state.captureSegmentLastSharpnessSampleAt = -1;
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
        scheduleSpeechCapture(panel, currentTime, true);
      } else if (level >= state.captureSegmentPeakLevel) {
        state.captureSegmentPeakLevel = level;
        scheduleSpeechCapture(panel, currentTime);
      }

      if (state.captureVoiceActive) {
        scheduleSpeechCapture(panel, currentTime);
      }

      if (state.captureVoiceActive && currentTime - state.captureSegmentStartTime >= CAPTURE_MAX_SEGMENT_SECONDS) {
        finalizeAudioSegment(panel);
        state.captureVoiceActive = true;
        state.captureSegmentStartTime = currentTime;
        state.captureSegmentPeakLevel = level;
        state.captureSegmentBestSharpness = 0;
        state.captureSegmentLastSharpnessSampleAt = -1;
        state.captureSilenceStartedAt = -1;
        state.captureSegmentTailSnapshot = null;
        scheduleSpeechCapture(panel, currentTime, true);
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
      setCaptureStatus(panel, `正在按字幕逐句采集，共 ${state.cues.length} 句。建议从头播放，脚本会尽量为每句保留一张代表画面。`);
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
      state.captureSegmentBestSharpness = 0;
      state.captureSegmentLastSharpnessSampleAt = -1;
      state.video = video;
      // Reset video to start position so we can capture pre-roll frames
      video.currentTime = 0;
      state.capturePreRollSamples = 0;
      state.captureFirstVoiceSegmentStarted = false;
      setCaptureStatus(panel, `正在按语音段采集，脚本会在每段语音结束后挑选一张代表画面。当前已记录 ${state.captureItems.length} 张。`);
      updateCaptureButtons(panel);
      startCaptureLoop(panel);
    } catch (error) {
      setCaptureStatus(panel, error instanceof Error ? error.message : '自动采集启动失败');
      updateCaptureButtons(panel);
    }
  }

  function stopAutoCapture(panel) {
    if (!state.captureRunning) {
      setCaptureStatus(panel, state.captureItems.length > 0 ? `当前已暂停采集，已保留 ${state.captureItems.length} 张画面。` : '还没有开始采集。');
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
    state.captureSegmentBestSharpness = 0;
    state.captureSegmentLastSharpnessSampleAt = -1;
    setCaptureStatus(
      panel,
      state.captureItems.length > 0
        ? state.captureMode === 'cue'
          ? `已暂停逐句采集，当前已保留 ${state.captureItems.length} 张画面。`
          : `已暂停采集，当前已保留 ${state.captureItems.length} 张画面。`
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
    closeCapturePreview();
    setCaptureStatus(panel, '已清空当前采集结果。');
    renderCaptureList(panel);
  }

  function buildPanel() {
    const existing = document.getElementById(PANEL_ID);
    if (existing) {
      ensureCapturePreviewOverlay(existing);
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

    const transcriptSection = document.createElement('section');
    transcriptSection.className = 'bili-transcript-section';

    const transcriptHeading = document.createElement('h2');
    transcriptHeading.className = 'bili-transcript-section-heading';
    transcriptHeading.textContent = '字幕/台词阅读';

    const transcriptNote = document.createElement('div');
    transcriptNote.className = 'bili-transcript-section-note';
    transcriptNote.textContent = '自动跟随播放进度高亮当前句，也可以点击任意一句快速跳转。';

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
    captureHeading.textContent = '画面采集';

    const captureNote = document.createElement('div');
    captureNote.className = 'bili-transcript-section-note';
    captureNote.textContent = '有字幕时按句采集，没字幕时按语音片段采集；可手动补抓、预览和删除。';

    const captureControls = document.createElement('div');
    captureControls.className = 'bili-transcript-capture-controls';

    const capturePrimaryRow = document.createElement('div');
    capturePrimaryRow.className = 'bili-capture-controls-row';

    const captureExportRow = document.createElement('div');
    captureExportRow.className = 'bili-capture-controls-row';

    const startCaptureButton = createButton('开始自动采集', () => {
      startAutoCapture(panel);
    });
    startCaptureButton.dataset.captureAction = 'start';
    startCaptureButton.classList.add('bili-btn-primary');

    const stopCaptureButton = createButton('暂停采集', () => {
      stopAutoCapture(panel);
    });
    stopCaptureButton.dataset.captureAction = 'stop';
    stopCaptureButton.classList.add('bili-btn-stop');

    const snapshotButton = createButton('手动截图', () => {
      try {
        captureCurrentFrame(panel, 'manual');
      } catch (error) {
        setCaptureStatus(panel, error instanceof Error ? error.message : '手动截图失败');
        updateCaptureButtons(panel);
      }
    });
    snapshotButton.dataset.captureAction = 'snapshot';
    snapshotButton.classList.add('bili-btn-capture');

    const exportButton = createButton('导出 PPTX', () => {
      exportCaptureSlides(panel);
    });
    exportButton.dataset.captureAction = 'export';
    exportButton.classList.add('bili-btn-export');

    const exportHtmlButton = createButton('导出 HTML', () => {
      exportCaptureDocument(panel);
    });
    exportHtmlButton.classList.add('bili-btn-export-html');

    const clearButton = createButton('清空结果', () => {
      clearCaptureItems(panel);
    });
    clearButton.dataset.captureAction = 'clear';
    clearButton.classList.add('bili-btn-danger');

    capturePrimaryRow.appendChild(startCaptureButton);
    capturePrimaryRow.appendChild(stopCaptureButton);
    capturePrimaryRow.appendChild(snapshotButton);

    captureExportRow.appendChild(exportButton);
    captureExportRow.appendChild(exportHtmlButton);
    captureExportRow.appendChild(clearButton);

    captureControls.appendChild(capturePrimaryRow);
    captureControls.appendChild(captureExportRow);

    const captureStatus = document.createElement('div');
    captureStatus.className = 'bili-capture-status';
    captureStatus.textContent = state.captureStatus;

    const captureList = document.createElement('div');
    captureList.className = 'bili-capture-list';

    transcriptSection.appendChild(transcriptHeading);
    transcriptSection.appendChild(transcriptNote);
    transcriptSection.appendChild(current);
    transcriptSection.appendChild(controls);
    transcriptSection.appendChild(status);
    transcriptSection.appendChild(list);
    body.appendChild(transcriptSection);
    captureSection.appendChild(captureHeading);
    captureSection.appendChild(captureNote);
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
    applyPanelPosition(panel, persistedState);

    document.body.appendChild(panel);
    ensureCapturePreviewOverlay(panel);
    renderCaptureList(panel);
    updateCaptureButtons(panel);
    setupPanelDragging(panel);
    setupPanelResizing(panel);
    return panel;
  }

  function setupPanelDragging(panel) {
    const header = panel.querySelector('.bili-transcript-header');
    if (!header) {
      return;
    }

    let dragging = false;
    let offsetX = 0;
    let offsetY = 0;

    const onMouseDown = (event) => {
      if (window.matchMedia('(max-width: 960px)').matches) {
        return;
      }

      if (event.target instanceof Element && event.target.closest('button, select, input, textarea, a')) {
        return;
      }

      const rect = panel.getBoundingClientRect();
      dragging = true;
      offsetX = event.clientX - rect.left;
      offsetY = event.clientY - rect.top;

      panel.style.right = 'auto';
      panel.style.bottom = 'auto';
      event.preventDefault();
    };

    const onMouseMove = (event) => {
      if (!dragging) {
        return;
      }

      const maxLeft = Math.max(0, window.innerWidth - panel.offsetWidth);
      const maxTop = Math.max(0, window.innerHeight - panel.offsetHeight);
      const nextLeft = Math.min(Math.max(0, event.clientX - offsetX), maxLeft);
      const nextTop = Math.min(Math.max(0, event.clientY - offsetY), maxTop);

      panel.style.left = `${nextLeft}px`;
      panel.style.top = `${nextTop}px`;
    };

    const onMouseUp = () => {
      if (!dragging) {
        return;
      }

      dragging = false;
      savePanelState(panel);
    };

    header.addEventListener('mousedown', onMouseDown);
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  }

  function setupPanelResizing(panel) {
    let isResizing = false;
    let startX = 0;
    let startY = 0;
    let startWidth = 0;
    let startHeight = 0;

    const handleMouseDown = (e) => {
      // Only trigger if clicking near the bottom-right corner (resize handle area)
      const rect = panel.getBoundingClientRect();
      const distFromBottom = rect.bottom - e.clientY;
      const distFromRight = rect.right - e.clientX;
      
      if (distFromBottom < 20 && distFromRight < 20) {
        isResizing = true;
        startX = e.clientX;
        startY = e.clientY;
        startWidth = panel.offsetWidth;
        startHeight = panel.offsetHeight;
        e.preventDefault();
      }
    };

    const handleMouseMove = (e) => {
      if (!isResizing) return;

      const deltaX = e.clientX - startX;
      const deltaY = e.clientY - startY;

      const newWidth = Math.max(250, startWidth + deltaX);
      const newHeight = Math.max(200, startHeight + deltaY);

      panel.style.width = newWidth + 'px';
      panel.style.height = newHeight + 'px';
    };

    const handleMouseUp = () => {
      isResizing = false;
    };

    panel.addEventListener('mousedown', handleMouseDown);
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
  }

  function updatePanelTitle(panel) {
    const title = panel.querySelector('.bili-transcript-title');
    if (title) {
      title.textContent = getVideoTitle();
    }
  }

  function setCurrentText(panel, text) {
    const current = panel.querySelector('.bili-transcript-current');
    current.textContent = text || '等待当前台词...';
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
      renderStatus(panel, '这个视频没有公开字幕，暂时无法显示“字幕/台词阅读”内容。');
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
