import type { AuthStorage, ModelRegistry } from "@mariozechner/pi-coding-agent";
import type {
  ImageApiType,
  ImageModelDescriptor,
  ImageModelCatalogResponse,
  ImageReasoningEffort,
  OfficeHost,
} from "@pi-office/pi-office-pack";

interface ImageModelEntry {
  provider: string;
  modelId: string;
  modelName: string;
  apiType: ImageApiType;
  supportsReasoningEffort: boolean;
  supportedAspectRatios: string[];
  supportedSizes: string[];
  priority: number;
}

const COMMON_ASPECT_RATIOS = ["1:1", "16:9", "9:16", "4:3", "3:4", "3:2", "2:3"];

const IMAGE_MODEL_CATALOG: ImageModelEntry[] = [
  {
    provider: "openai",
    modelId: "gpt-image-1",
    modelName: "GPT Image 1",
    apiType: "openai-images",
    supportsReasoningEffort: true,
    supportedAspectRatios: COMMON_ASPECT_RATIOS,
    supportedSizes: ["1K", "2K", "4K"],
    priority: 1,
  },
  {
    provider: "openai",
    modelId: "dall-e-3",
    modelName: "DALL-E 3",
    apiType: "openai-images",
    supportsReasoningEffort: false,
    supportedAspectRatios: ["1:1", "16:9", "9:16"],
    supportedSizes: ["1K", "2K"],
    priority: 2,
  },
  {
    provider: "openai",
    modelId: "dall-e-2",
    modelName: "DALL-E 2",
    apiType: "openai-images",
    supportsReasoningEffort: false,
    supportedAspectRatios: ["1:1"],
    supportedSizes: ["1K"],
    priority: 10,
  },
  {
    provider: "google",
    modelId: "gemini-2.0-flash-preview-image-generation",
    modelName: "Gemini 2.0 Flash Image",
    apiType: "google-generative-ai-image",
    supportsReasoningEffort: false,
    supportedAspectRatios: COMMON_ASPECT_RATIOS,
    supportedSizes: ["1K", "2K"],
    priority: 3,
  },
  {
    provider: "google",
    modelId: "imagen-3.0-generate-002",
    modelName: "Imagen 3",
    apiType: "google-generative-ai-image",
    supportsReasoningEffort: false,
    supportedAspectRatios: COMMON_ASPECT_RATIOS,
    supportedSizes: ["1K", "2K"],
    priority: 4,
  },
  {
    provider: "openrouter",
    modelId: "openai/gpt-5-image",
    modelName: "GPT-5 Image (OpenRouter)",
    apiType: "openai-chat-image",
    supportsReasoningEffort: true,
    supportedAspectRatios: COMMON_ASPECT_RATIOS,
    supportedSizes: ["1K", "2K", "4K"],
    priority: 5,
  },
  {
    provider: "openrouter",
    modelId: "openai/gpt-5-image-mini",
    modelName: "GPT-5 Image Mini (OpenRouter)",
    apiType: "openai-chat-image",
    supportsReasoningEffort: true,
    supportedAspectRatios: COMMON_ASPECT_RATIOS,
    supportedSizes: ["1K", "2K"],
    priority: 6,
  },
  {
    provider: "openrouter",
    modelId: "google/gemini-2.5-flash-image",
    modelName: "Gemini 2.5 Flash Image (OpenRouter)",
    apiType: "openai-chat-image",
    supportsReasoningEffort: false,
    supportedAspectRatios: COMMON_ASPECT_RATIOS,
    supportedSizes: ["1K", "2K"],
    priority: 7,
  },
  {
    provider: "openrouter",
    modelId: "google/gemini-3-pro-image-preview",
    modelName: "Gemini 3 Pro Image (OpenRouter)",
    apiType: "openai-chat-image",
    supportsReasoningEffort: false,
    supportedAspectRatios: COMMON_ASPECT_RATIOS,
    supportedSizes: ["1K", "2K"],
    priority: 8,
  },
  {
    provider: "openrouter",
    modelId: "google/gemini-2.5-flash-image-preview",
    modelName: "Gemini 2.5 Flash Image Preview (OpenRouter)",
    apiType: "openai-chat-image",
    supportsReasoningEffort: false,
    supportedAspectRatios: COMMON_ASPECT_RATIOS,
    supportedSizes: ["1K", "2K"],
    priority: 9,
  },
  {
    provider: "openrouter",
    modelId: "black-forest-labs/flux.2-pro",
    modelName: "FLUX.2 Pro (OpenRouter)",
    apiType: "openai-chat-image",
    supportsReasoningEffort: false,
    supportedAspectRatios: COMMON_ASPECT_RATIOS,
    supportedSizes: ["1K", "2K"],
    priority: 10,
  },
  {
    provider: "openrouter",
    modelId: "black-forest-labs/flux.2-max",
    modelName: "FLUX.2 Max (OpenRouter)",
    apiType: "openai-chat-image",
    supportsReasoningEffort: false,
    supportedAspectRatios: COMMON_ASPECT_RATIOS,
    supportedSizes: ["1K", "2K", "4K"],
    priority: 11,
  },
  {
    provider: "openrouter",
    modelId: "bytedance-seed/seedream-4.5",
    modelName: "Seedream 4.5 (OpenRouter)",
    apiType: "openai-chat-image",
    supportsReasoningEffort: false,
    supportedAspectRatios: COMMON_ASPECT_RATIOS,
    supportedSizes: ["1K", "2K"],
    priority: 12,
  },
];

const SIZE_DIMENSIONS: Record<string, { short: number; long: number }> = {
  "1K": { short: 1024, long: 1024 },
  "2K": { short: 1536, long: 2048 },
  "4K": { short: 2160, long: 3840 },
};

const ASPECT_RATIO_MAP: Record<string, [number, number]> = {
  "1:1": [1, 1],
  "16:9": [16, 9],
  "9:16": [9, 16],
  "4:3": [4, 3],
  "3:4": [3, 4],
  "3:2": [3, 2],
  "2:3": [2, 3],
  "4:5": [4, 5],
  "5:4": [5, 4],
  "21:9": [21, 9],
};

export interface ImageGenParams {
  prompt: string;
  modelKey?: string | undefined;
  aspectRatio?: string | undefined;
  size?: string | undefined;
  quality?: string | undefined;
  reasoningEffort?: ImageReasoningEffort | undefined;
  host?: OfficeHost | undefined;
}

export interface ImageGenResult {
  base64: string;
  mimeType: string;
  width: number;
  height: number;
  modelKey: string;
  modelName: string;
}

// OpenAI size mapping from aspect ratios (matches reference implementation)
const OPENAI_SIZES: Record<string, string> = {
  "1:1": "1024x1024",
  "16:9": "1536x1024",
  "9:16": "1024x1536",
  "4:3": "1536x1024",
  "3:4": "1024x1536",
  "3:2": "1536x1024",
  "2:3": "1024x1536",
};

function resolveOpenAiSize(aspectRatio: string, modelId: string): string {
  if (modelId === "dall-e-2") return "1024x1024";
  return OPENAI_SIZES[aspectRatio] || "1024x1024";
}

function resolveGoogleAspectLabel(aspectRatio: string): string {
  return aspectRatio.replace(":", ":");
}

function extractDataUriBase64(value: string): string | undefined {
  const match = value.match(/^data:[^;]+;base64,(.+)$/s);
  return match?.[1];
}

function extractBase64FromImageArray(images: unknown): string | undefined {
  if (!Array.isArray(images) || !images.length) return undefined;
  for (const img of images) {
    if (!img || typeof img !== "object") continue;
    const entry = img as Record<string, unknown>;
    // { b64_json: "..." }
    if (typeof entry.b64_json === "string" && entry.b64_json) return entry.b64_json;
    // { url: "data:..." }
    if (typeof entry.url === "string" && entry.url) return extractDataUriBase64(entry.url) ?? entry.url;
    // { data: "..." }
    if (typeof entry.data === "string" && entry.data) return extractDataUriBase64(entry.data) ?? entry.data;
    // { base64: "..." }
    if (typeof entry.base64 === "string" && entry.base64) return entry.base64;
  }
  return undefined;
}

function extractBase64FromContentParts(content: unknown): string | undefined {
  if (!Array.isArray(content)) return undefined;
  for (const part of content) {
    if (!part || typeof part !== "object") continue;
    const p = part as Record<string, unknown>;
    const pType = String(p.type ?? "");

    // OpenAI native: { type: "image_url", image_url: { url: "data:..." } }
    if (pType === "image_url" && p.image_url && typeof p.image_url === "object") {
      const url = String((p.image_url as Record<string, unknown>).url ?? "");
      if (url) return extractDataUriBase64(url) ?? url;
    }

    // OpenRouter variant: { type: "image", image_url: { url: "data:..." } }
    if (pType === "image" && p.image_url && typeof p.image_url === "object") {
      const url = String((p.image_url as Record<string, unknown>).url ?? "");
      if (url) return extractDataUriBase64(url) ?? url;
    }

    // Anthropic-style: { type: "image", source: { type: "base64", data: "..." } }
    if (pType === "image" && p.source && typeof p.source === "object") {
      const data = String((p.source as Record<string, unknown>).data ?? "");
      if (data) return data;
    }

    // Direct data field: { type: "image", data: "..." }
    if (pType === "image" && typeof p.data === "string" && p.data) {
      return extractDataUriBase64(p.data) ?? p.data;
    }

    // OpenAI output_image: { type: "output_image", output_image: { b64_json: "..." } }
    if (pType === "output_image" && p.output_image && typeof p.output_image === "object") {
      const b64 = String((p.output_image as Record<string, unknown>).b64_json ?? "");
      if (b64) return b64;
    }

    // Generic b64_json at part level
    if (typeof p.b64_json === "string" && p.b64_json) return p.b64_json;

    // Generic url at part level
    if (typeof p.url === "string" && p.url) {
      const extracted = extractDataUriBase64(p.url);
      if (extracted) return extracted;
    }
  }
  return undefined;
}

function extractBase64FromChatResponse(message: Record<string, unknown>): string | undefined {
  // OpenAI/OpenRouter newer format: message.images array (content is null)
  if (message.images) {
    const result = extractBase64FromImageArray(message.images);
    if (result) return result;
  }

  // Standard content array format
  if (Array.isArray(message.content)) {
    const result = extractBase64FromContentParts(message.content);
    if (result) return result;
  }

  // String content that might be a data URI
  if (typeof message.content === "string" && message.content.startsWith("data:image")) {
    return extractDataUriBase64(message.content);
  }

  // Check top-level data/output fields as fallback
  if (typeof message.data === "string" && message.data) {
    return extractDataUriBase64(message.data) ?? message.data;
  }

  return undefined;
}

function summarizeResponse(value: unknown): unknown {
  if (value === null || value === undefined) return value;
  if (typeof value === "string") {
    if (value.length > 200) return `[string, ${value.length} chars, starts: ${value.slice(0, 80)}...]`;
    return value;
  }
  if (typeof value !== "object") return value;
  if (Array.isArray(value)) {
    return value.slice(0, 5).map(summarizeResponse);
  }
  const result: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
    result[k] = summarizeResponse(v);
  }
  return result;
}

export function resolveImageDimensions(
  host: OfficeHost | undefined,
  aspectRatio: string | undefined,
  sizeTier: string | undefined,
): { width: number; height: number; aspectRatio: string; sizeTier: string } {
  const effectiveAspect = aspectRatio ?? (host === "powerpoint" ? "16:9" : "4:3");
  const [aw, ah] = ASPECT_RATIO_MAP[effectiveAspect] ?? [4, 3];
  const isLandscape = aw >= ah;

  let effectiveSize: string;
  if (sizeTier && SIZE_DIMENSIONS[sizeTier]) {
    effectiveSize = sizeTier;
  } else if (host === "powerpoint") {
    effectiveSize = "2K";
  } else if (host === "excel") {
    effectiveSize = "1K";
  } else {
    effectiveSize = "2K";
  }

  const dim = SIZE_DIMENSIONS[effectiveSize] ?? SIZE_DIMENSIONS["2K"]!;
  let width: number;
  let height: number;

  if (aw === ah) {
    width = dim.short;
    height = dim.short;
  } else if (isLandscape) {
    width = dim.long;
    height = Math.round((dim.long * ah) / aw);
  } else {
    height = dim.long;
    width = Math.round((dim.long * aw) / ah);
  }

  width = Math.round(width / 2) * 2;
  height = Math.round(height / 2) * 2;

  return { width, height, aspectRatio: effectiveAspect, sizeTier: effectiveSize };
}

function buildDocumentContextHint(host: OfficeHost | undefined, width: number, height: number): string {
  if (host === "powerpoint") {
    return `This image will be used in a presentation slide (${width}x${height}). Design for visual impact at presentation scale with bold, clear visuals.`;
  }
  if (host === "excel") {
    return `This image will be placed in a spreadsheet. Keep it compact, clear, and professional.`;
  }
  return `This image will be inserted inline in a document. Keep it clean and professional.`;
}

export class ImageGenService {
  constructor(
    private readonly authStorage: AuthStorage,
    private readonly modelRegistry: ModelRegistry,
  ) {}

  listAvailableModels(): ImageModelDescriptor[] {
    const configuredProviders = new Set<string>();
    for (const model of this.modelRegistry.getAvailable()) {
      if (model.provider) configuredProviders.add(model.provider);
    }
    const stored = new Set(this.authStorage.list());
    const available = new Set([...configuredProviders, ...stored]);

    return IMAGE_MODEL_CATALOG
      .sort((a, b) => a.priority - b.priority)
      .map((entry) => ({
        provider: entry.provider,
        modelId: entry.modelId,
        modelName: entry.modelName,
        apiType: entry.apiType,
        supportsReasoningEffort: entry.supportsReasoningEffort,
        supportedAspectRatios: entry.supportedAspectRatios,
        supportedSizes: entry.supportedSizes,
        configured: available.has(entry.provider),
      }));
  }

  getCatalog(): ImageModelCatalogResponse {
    const models = this.listAvailableModels();
    const firstConfigured = models.find((m) => m.configured);
    return {
      models,
      defaultModelKey: firstConfigured ? `${firstConfigured.provider}::${firstConfigured.modelId}` : "",
    };
  }

  resolveModel(preferredKey: string | undefined): ImageModelEntry | undefined {
    if (preferredKey) {
      const [provider, modelId] = preferredKey.split("::");
      const match = IMAGE_MODEL_CATALOG.find((e) => e.provider === provider && e.modelId === modelId);
      if (match) return match;
    }
    const available = this.listAvailableModels().filter((m) => m.configured);
    if (!available.length) return undefined;
    const entry = IMAGE_MODEL_CATALOG.find(
      (e) => e.provider === available[0]!.provider && e.modelId === available[0]!.modelId,
    );
    return entry;
  }

  async generate(params: ImageGenParams): Promise<ImageGenResult> {
    const model = this.resolveModel(params.modelKey);
    if (!model) {
      throw new Error("No image generation model is available. Configure an AI provider with image support first.");
    }

    const { width, height, aspectRatio, sizeTier } = resolveImageDimensions(
      params.host,
      params.aspectRatio,
      params.size,
    );

    const contextHint = buildDocumentContextHint(params.host, width, height);
    const augmentedPrompt = `${contextHint}\n\n${params.prompt}`;

    const apiKey = await this.getApiKey(model.provider);
    if (!apiKey) {
      throw new Error(`No API key available for provider "${model.provider}". Add one in Settings > AI Providers.`);
    }

    const quality = params.quality ?? "auto";
    const reasoningEffort = model.supportsReasoningEffort ? (params.reasoningEffort ?? "high") : undefined;

    let result: ImageGenResult;
    switch (model.apiType) {
      case "openai-images":
        result = await this.generateOpenAiImages(model, apiKey, augmentedPrompt, width, height, aspectRatio, quality, reasoningEffort);
        break;
      case "openai-chat-image":
        result = await this.generateOpenAiChatImage(model, apiKey, augmentedPrompt, width, height, quality, aspectRatio, sizeTier);
        break;
      case "google-generative-ai-image":
        result = await this.generateGoogleImage(model, apiKey, augmentedPrompt, aspectRatio, sizeTier);
        break;
      default:
        throw new Error(`Unsupported image API type: ${model.apiType}`);
    }

    return result;
  }

  private async getApiKey(provider: string): Promise<string | undefined> {
    // Try auth storage first (most direct path for API keys)
    try {
      const stored = this.authStorage.get(provider);
      if (stored && typeof stored === "object" && "key" in stored && stored.key) {
        return stored.key as string;
      }
    } catch {
      // fall through
    }

    // Try model registry (covers OAuth and other auth methods)
    try {
      const available = this.modelRegistry.getAvailable();
      const providerModel = available.find((m) => m.provider === provider);
      if (providerModel) {
        const auth = await this.modelRegistry.getApiKeyAndHeaders(providerModel);
        if (auth.ok && auth.apiKey) return auth.apiKey;
      }
    } catch {
      // fall through
    }

    // Check common env vars as last resort
    const envKeys: Record<string, string> = {
      openai: "OPENAI_API_KEY",
      google: "GEMINI_API_KEY",
      openrouter: "OPENROUTER_API_KEY",
    };
    const envKey = envKeys[provider];
    if (envKey && process.env[envKey]) return process.env[envKey];

    return undefined;
  }

  private getBaseUrl(provider: string): string {
    if (provider === "openrouter") return "https://openrouter.ai/api";
    if (provider === "google") return "https://generativelanguage.googleapis.com";
    return "https://api.openai.com";
  }

  private async generateOpenAiImages(
    model: ImageModelEntry,
    apiKey: string,
    prompt: string,
    width: number,
    height: number,
    aspectRatio: string,
    quality: string,
    reasoningEffort: string | undefined,
  ): Promise<ImageGenResult> {
    const size = resolveOpenAiSize(aspectRatio, model.modelId);
    const baseUrl = this.getBaseUrl(model.provider);
    const body: Record<string, unknown> = {
      model: model.modelId,
      prompt,
      n: 1,
      size,
      response_format: "b64_json",
    };
    // gpt-image-1 accepts quality directly; dall-e-3 maps to "hd"/"standard"
    if (model.modelId === "gpt-image-1") {
      if (quality !== "auto") body.quality = quality;
    } else if (model.modelId === "dall-e-3") {
      body.quality = quality === "high" ? "hd" : "standard";
    }
    if (reasoningEffort) body.reasoning = { effort: reasoningEffort };

    const response = await fetch(`${baseUrl}/v1/images/generations`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`OpenAI image generation failed (${response.status}): ${text}`);
    }

    const json = (await response.json()) as { data: Array<{ b64_json: string; revised_prompt?: string }> };
    const imageData = json.data[0];
    if (!imageData?.b64_json) throw new Error("No image data in OpenAI response.");

    const [w, h] = size.split("x").map(Number);
    return {
      base64: imageData.b64_json,
      mimeType: "image/png",
      width: w ?? width,
      height: h ?? height,
      modelKey: `${model.provider}::${model.modelId}`,
      modelName: model.modelName,
    };
  }

  private async generateOpenAiChatImage(
    model: ImageModelEntry,
    apiKey: string,
    prompt: string,
    width: number,
    height: number,
    quality: string,
    aspectRatio?: string,
    sizeTier?: string,
  ): Promise<ImageGenResult> {
    const baseUrl = this.getBaseUrl(model.provider);
    const body: Record<string, unknown> = {
      model: model.modelId,
      messages: [{ role: "user", content: prompt }],
      // Reference: modalities: ["image"] — image-only so image endpoints route successfully
      modalities: ["image"],
    };

    // Add image_config for OpenRouter aspect ratio and size control
    const imageConfig: Record<string, string> = {};
    if (aspectRatio && aspectRatio !== "1:1") {
      imageConfig.aspect_ratio = aspectRatio;
    }
    if (sizeTier && sizeTier !== "1K") {
      imageConfig.image_size = sizeTier;
    }
    if (Object.keys(imageConfig).length > 0) {
      body.image_config = imageConfig;
    }

    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    };
    if (model.provider === "openrouter") {
      headers["HTTP-Referer"] = "https://pi-office.local";
      headers["X-Title"] = "Pi-Office";
    }

    const response = await fetch(`${baseUrl}/v1/chat/completions`, {
      method: "POST",
      headers,
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Chat image generation failed (${response.status}): ${text}`);
    }

    const json = await response.json();
    console.log("[generate_image] chat-image raw response structure:", JSON.stringify(summarizeResponse(json), null, 2));

    const choices = (json as Record<string, unknown>).choices;
    if (!Array.isArray(choices) || !choices.length) {
      throw new Error("No choices in chat image response.");
    }

    const message = (choices[0] as Record<string, unknown>).message as Record<string, unknown> | undefined;
    if (!message) throw new Error("No message in chat image response choice.");

    // Reference impl: images are in message.images[], each with { type, image_url: { url: "data:..." } }
    if (Array.isArray(message.images) && (message.images as unknown[]).length > 0) {
      const images = message.images as Array<Record<string, unknown>>;
      const firstImg = images[0]!;
      const imageUrl = firstImg.image_url as Record<string, unknown> | undefined;
      if (imageUrl) {
        const url = String(imageUrl.url ?? "");
        if (url.startsWith("data:")) {
          const commaIdx = url.indexOf(",");
          const header = url.slice(0, commaIdx);
          const data = url.slice(commaIdx + 1);
          const mimeType = header.match(/data:(.*);/)?.[1] || "image/png";
          return {
            base64: data,
            mimeType,
            width,
            height,
            modelKey: `${model.provider}::${model.modelId}`,
            modelName: model.modelName,
          };
        }
        // URL-based: fetch and convert to base64
        if (url) {
          const imgResponse = await fetch(url);
          const buffer = await imgResponse.arrayBuffer();
          return {
            base64: Buffer.from(buffer).toString("base64"),
            mimeType: imgResponse.headers.get("content-type") || "image/png",
            width,
            height,
            modelKey: `${model.provider}::${model.modelId}`,
            modelName: model.modelName,
          };
        }
      }
    }

    // Fallback: try generic extraction from all known response formats
    const base64 = extractBase64FromChatResponse(message);
    if (!base64) {
      throw new Error(
        `No image data found in chat completion response. Response keys: ${Object.keys(message).join(", ")}. ` +
        `Content type: ${typeof message.content}. ` +
        `Content preview: ${JSON.stringify(summarizeResponse(message.content)).slice(0, 500)}`,
      );
    }

    return {
      base64,
      mimeType: "image/png",
      width,
      height,
      modelKey: `${model.provider}::${model.modelId}`,
      modelName: model.modelName,
    };
  }

  private async generateGoogleImage(
    model: ImageModelEntry,
    apiKey: string,
    prompt: string,
    aspectRatio: string,
    sizeTier: string,
  ): Promise<ImageGenResult> {
    const baseUrl = this.getBaseUrl(model.provider);
    const body = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: {
        responseModalities: ["TEXT", "IMAGE"],
      },
    };

    const response = await fetch(
      `${baseUrl}/v1beta/models/${model.modelId}:generateContent?key=${apiKey}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      },
    );

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Google image generation failed (${response.status}): ${text}`);
    }

    const json = (await response.json()) as {
      candidates: Array<{
        content: {
          parts: Array<{ inlineData?: { mimeType: string; data: string }; text?: string }>;
        };
      }>;
    };

    const candidate = json.candidates?.[0];
    if (!candidate) throw new Error("No candidate in Google response.");

    const imagePart = candidate.content.parts.find((p) => p.inlineData?.data);
    if (!imagePart?.inlineData) throw new Error("No image data in Google response.");

    const dims = resolveImageDimensions(undefined, aspectRatio, sizeTier);
    return {
      base64: imagePart.inlineData.data,
      mimeType: imagePart.inlineData.mimeType || "image/png",
      width: dims.width,
      height: dims.height,
      modelKey: `${model.provider}::${model.modelId}`,
      modelName: model.modelName,
    };
  }
}
