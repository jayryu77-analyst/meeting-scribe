// 녹음 → 텍스트 변환 (Whisper, 휴대폰·PC 안에서 처리 · 외부 서버로 음성을 보내지 않음)
// 처음 한 번 모델을 내려받아 브라우저에 저장해 두고, 다음부터는 저장된 모델을 씁니다.
import { pipeline, env } from "https://cdn.jsdelivr.net/npm/@huggingface/transformers@3.8.1/dist/transformers.min.js";
env.allowLocalModels = false;

let asr = null, loaded = "", device = "wasm";

async function hasWebGPU() {
  try { return !!(navigator.gpu && await navigator.gpu.requestAdapter()); } catch (e) { return false; }
}
self.addEventListener("error", e => postMessage({ type: "fatal", message: String(e.message || e) }));
self.addEventListener("unhandledrejection", e => postMessage({ type: "fatal", message: String((e.reason && e.reason.message) || e.reason) }));
async function load(model, pref) {
  if (asr && loaded === model + pref) return;
  const progress = p => { if (p.status === "progress" && p.total) postMessage({ type: "dl", file: p.file, loaded: p.loaded, total: p.total }); };
  device = pref !== "wasm" && (await hasWebGPU()) ? "webgpu" : "wasm";   // 휴대폰은 그래픽 가속이 불안정해 일반 모드 사용
  try {
    asr = await pipeline("automatic-speech-recognition", model, {
      device, progress_callback: progress,
      dtype: device === "webgpu" ? { encoder_model: "fp32", decoder_model_merged: "q4" } : "q8",
    });
  } catch (e) {
    if (device !== "webgpu") throw e;
    device = "wasm";   // 그래픽 가속이 안 되면 일반 방식으로 다시
    asr = await pipeline("automatic-speech-recognition", model, { device, dtype: "q8", progress_callback: progress });
  }
  loaded = model + pref;
  postMessage({ type: "ready", device });
}

self.onmessage = async e => {
  const m = e.data;
  try {
    if (m.type === "run") {
      await load(m.model, m.device || "auto");
      postMessage({ type: "stage", stage: "run", device });
      const opt = { task: m.task || "transcribe", chunk_length_s: 30, stride_length_s: 5, return_timestamps: true };
      if (m.lang) opt.language = m.lang;
      const out = await asr(m.audio, opt);
      postMessage({ type: "done", id: m.id, text: out.text || "", chunks: out.chunks || [] });
    }
  } catch (err) {
    postMessage({ type: "error", id: m.id, message: String((err && err.message) || err) });
  }
};
