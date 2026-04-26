import { useEffect, useState } from "react";
import type { SimpleIcon } from "simple-icons";
import {
  siAirtable,
  siConfluence,
  siFigma,
  siGithub,
  siGitlab,
  siGoogledrive,
  siJira,
  siNotion,
  siObsidian,
  siPaypal,
  siPerplexity,
  siPostgresql,
  siPosthog,
  siQdrant,
  siShopify,
  siStripe,
} from "simple-icons";

export function GearIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M19.4 13a7.7 7.7 0 0 0 .1-1 7.7 7.7 0 0 0-.1-1l2.1-1.6-2-3.4-2.5 1a7 7 0 0 0-1.7-1L15 2h-6l-.3 3a7 7 0 0 0-1.7 1l-2.5-1-2 3.4L4.6 11a7.7 7.7 0 0 0-.1 1 7.7 7.7 0 0 0 .1 1l-2.1 1.6 2 3.4 2.5-1a7 7 0 0 0 1.7 1L9 22h6l.3-3a7 7 0 0 0 1.7-1l2.5 1 2-3.4zM12 15.2A3.2 3.2 0 1 1 12 8.8a3.2 3.2 0 0 1 0 6.4Z"
        fill="currentColor"
      />
    </svg>
  );
}

export function CloseIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M6.7 5.3 12 10.6l5.3-5.3 1.4 1.4L13.4 12l5.3 5.3-1.4 1.4L12 13.4l-5.3 5.3-1.4-1.4 5.3-5.3-5.3-5.3z"
        fill="currentColor"
      />
    </svg>
  );
}

export function ArrowUpIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M12 5.2 6.7 10.5l1.4 1.4 2.9-2.9V18h2V9l2.9 2.9 1.4-1.4z"
        fill="currentColor"
      />
    </svg>
  );
}

export function StopIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <rect x="7" y="7" width="10" height="10" rx="2" fill="currentColor" />
    </svg>
  );
}

export function ChevronDownIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" width="12" height="12">
      <path d="M7 10l5 5 5-5z" fill="currentColor" />
    </svg>
  );
}

export function ChevronRightIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" width="12" height="12">
      <path d="M10 7l5 5-5 5z" fill="currentColor" />
    </svg>
  );
}

export function CheckIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" width="14" height="14">
      <path d="M9 16.2L4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z" fill="currentColor" />
    </svg>
  );
}

export function SearchIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" width="14" height="14">
      <path
        d="M15.5 14h-.79l-.28-.27A6.47 6.47 0 0 0 16 9.5 6.5 6.5 0 1 0 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"
        fill="currentColor"
      />
    </svg>
  );
}

export function BackArrowIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z" fill="currentColor" />
    </svg>
  );
}

export function UserIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"
        fill="currentColor"
      />
    </svg>
  );
}

export function ProviderIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4zm0 10.99h7c-.53 4.12-3.28 7.79-7 8.94V12H5V6.3l7-3.11v8.8z"
        fill="currentColor"
      />
    </svg>
  );
}

export function ModelIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M21 10.12h-6.78l2.74-2.82-2.73-2.7L12 6.83 9.77 4.6 7.04 7.3l2.74 2.82H3v2.12h6.78l-2.74 2.82 2.73 2.7L12 15.53l2.23 2.23 2.73-2.7-2.74-2.82H21v-2.12z"
        fill="currentColor"
      />
    </svg>
  );
}

export function ToolIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M22.7 19l-9.1-9.1c.9-2.3.4-5-1.5-6.9-2-2-5-2.4-7.4-1.3L9 6 6 9 1.6 4.7C.4 7.1.9 10.1 2.9 12.1c1.9 1.9 4.6 2.4 6.9 1.5l9.1 9.1c.4.4 1 .4 1.4 0l2.3-2.3c.5-.4.5-1.1.1-1.4z"
        fill="currentColor"
      />
    </svg>
  );
}

export function PrefsIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M3 17v2h6v-2H3zM3 5v2h10V5H3zm10 16v-2h8v-2h-8v-2h-2v6h2zM7 9v2H3v2h4v2h2V9H7zm14 4v-2H11v2h10zm-6-4h2V7h4V5h-4V3h-2v6z"
        fill="currentColor"
      />
    </svg>
  );
}

export function IntegrationIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M8 6a3 3 0 1 1 0 6H5v2h3a5 5 0 1 0 0-10H5v2h3zm8 4H8v4h8v-4zm0-6h-3v2h3a3 3 0 1 1 0 6h-3v2h3a5 5 0 1 0 0-10zm-8 8H5v2h3v-2z"
        fill="currentColor"
      />
    </svg>
  );
}

export function LinkIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M10.6 13.4a1 1 0 0 1 0-1.4l3.4-3.4a3 3 0 0 1 4.2 4.2l-1.8 1.8-1.4-1.4 1.8-1.8a1 1 0 0 0-1.4-1.4L12 13.4a1 1 0 0 1-1.4 0zm2.8-2.8a1 1 0 0 1 0 1.4L10 15.4a3 3 0 0 1-4.2-4.2l1.8-1.8L9 10.8 7.2 12.6a1 1 0 0 0 1.4 1.4l3.4-3.4a1 1 0 0 1 1.4 0z"
        fill="currentColor"
      />
    </svg>
  );
}

export function ShieldIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M12 2 4 5v6c0 5.2 3.4 9.9 8 11 4.6-1.1 8-5.8 8-11V5l-8-3zm0 9.8 4.1-4.1 1.4 1.4-5.5 5.5-3-3 1.4-1.4 1.6 1.6z"
        fill="currentColor"
      />
    </svg>
  );
}

export function MagicIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="m5 17 2 2-2 2-2-2 2-2Zm8.3-13.3 1.4 3.1 3.1 1.4-3.1 1.4-1.4 3.1-1.4-3.1-3.1-1.4 3.1-1.4 1.4-3.1ZM19 14l.9 2.1L22 17l-2.1.9L19 20l-.9-2.1L16 17l2.1-.9.9-2.1Zm-8.6-1.4 1.4 1.4-6.8 6.8H3v-2l7.4-7.4Z"
        fill="currentColor"
      />
    </svg>
  );
}

export function DiagnosticsIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M3 18h4v3H3v-3Zm7-7h4v10h-4V11Zm7-8h4v18h-4V3Zm-7 4h4v2h-4V7Zm-7 7h4v2H3v-2Z"
        fill="currentColor"
      />
    </svg>
  );
}

export function NewChatIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        fillRule="evenodd"
        clipRule="evenodd"
        d="M20 2H4C2.9 2 2 2.9 2 4v18l4-4h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zM13 6h-2v3H8v2h3v3h2v-3h3v-2h-3z"
        fill="currentColor"
      />
    </svg>
  );
}

export function PinIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M15 3 21 9l-2 2-2-2-3 3v5l-2 4-2-4v-5L7 9 5 11 3 9l6-6z" fill="currentColor" />
    </svg>
  );
}

export function ScopeIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M11 3h2v4h-2V3Zm0 14h2v4h-2v-4ZM3 11h4v2H3v-2Zm14 0h4v2h-4v-2Zm-8.9-5.5 1.4-1.4 2.5 2.5 2.5-2.5 1.4 1.4-2.5 2.5 2.5 2.5-1.4 1.4-2.5-2.5-2.5 2.5-1.4-1.4 2.5-2.5-2.5-2.5Z" fill="currentColor" />
    </svg>
  );
}

export function UploadIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M5 20h14v-2H5v2Zm7-16 5 5h-3v6h-4V9H7l5-5Z" fill="currentColor" />
    </svg>
  );
}

export function HistoryIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M13 3a9 9 0 0 0-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7a6.93 6.93 0 0 1-4.95-2.05l-1.41 1.41A8.96 8.96 0 0 0 13 21a9 9 0 0 0 0-18zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8h-1.5z"
        fill="currentColor"
      />
    </svg>
  );
}

export function TrashIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" width="14" height="14">
      <path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM8 9h8v10H8V9zm7.5-5l-1-1h-5l-1 1H5v2h14V4h-3.5z" fill="currentColor" />
    </svg>
  );
}

export function DownloadIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M5 20h14v-2H5v2Zm7-16v8h3l-5 5-5-5h3V4h4Z" fill="currentColor" />
    </svg>
  );
}

export function ExternalLinkIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M14 3h7v7h-2V6.4l-8.3 8.3-1.4-1.4L17.6 5H14V3ZM5 5h6v2H7v10h10v-4h2v6H5V5Z" fill="currentColor" />
    </svg>
  );
}

function connectorGlyph(iconKey: string): string {
  const map: Record<string, string> = {
    jira: "J",
    confluence: "C",
    "microsoft-365": "M365",
    figma: "F",
    slack: "S",
    notion: "N",
    obsidian: "O",
    granola: "G",
    supermemory: "SM",
    posthog: "P",
    launchdarkly: "LD",
    shopify: "S",
    stripe: "S",
    paypal: "P",
    airtable: "A",
    perplexity: "PX",
    parallel: "PW",
    "parallel-web": "PW",
    exa: "E",
    "exa-search": "E",
    qdrant: "Q",
    gdrive: "GD",
    "google-drive": "GD",
    github: "GH",
    gitlab: "GL",
    postgres: "PG",
    "postgresql-reader": "PG",
    tavily: "TV",
    custom: "MCP",
  };
  return map[iconKey] ?? iconKey.slice(0, 2).toUpperCase();
}

const CONNECTOR_MARKS: Record<string, SimpleIcon | undefined> = {
  jira: siJira,
  confluence: siConfluence,
  figma: siFigma,
  notion: siNotion,
  obsidian: siObsidian,
  posthog: siPosthog,
  shopify: siShopify,
  stripe: siStripe,
  paypal: siPaypal,
  airtable: siAirtable,
  perplexity: siPerplexity,
  qdrant: siQdrant,
  gdrive: siGoogledrive,
  "google-drive": siGoogledrive,
  github: siGithub,
  gitlab: siGitlab,
  postgres: siPostgresql,
  "postgresql-reader": siPostgresql,
};

const CONNECTOR_LOGO_ASSETS: Record<string, string | undefined> = {
  figma: "/connectors/figma.png",
  slack: "/connectors/slack.png",
  notion: "/connectors/notion.png",
  obsidian: "/connectors/obsidian.png",
  supermemory: "/connectors/supermemory.png",
  posthog: "/connectors/posthog.png",
  launchdarkly: "/connectors/launchdarkly.png",
  shopify: "/connectors/shopify.png",
  stripe: "/connectors/stripe.png",
  paypal: "/connectors/paypal.png",
  airtable: "/connectors/airtable.png",
  perplexity: "/connectors/perplexity.png",
  parallel: "/connectors/parallel.svg",
  "parallel-web": "/connectors/parallel.svg",
  exa: "/connectors/exa.jpg",
  "exa-search": "/connectors/exa.jpg",
  qdrant: "/connectors/qdrant.png",
  gdrive: "/connectors/google-drive.png",
  "google-drive": "/connectors/google-drive.png",
  github: "/connectors/github.svg",
  granola: "/connectors/granola.png",
  gitlab: "/connectors/gitlab.png",
  postgres: "/connectors/postgresql-reader.png",
  "postgresql-reader": "/connectors/postgresql-reader.png",
  tavily: "/connectors/tavily.jpg",
  "microsoft-365": "/connectors/microsoft-365.svg",
};

export function ConnectorBrandIcon({ iconKey, label }: { iconKey: string; label: string }) {
  const [loadFailed, setLoadFailed] = useState(false);
  const logoSrc = CONNECTOR_LOGO_ASSETS[iconKey];
  const mark = CONNECTOR_MARKS[iconKey];

  useEffect(() => {
    setLoadFailed(false);
  }, [iconKey]);

  if (logoSrc && !loadFailed) {
    return (
      <span className="connector-logo connector-logo-image" aria-hidden="true" title={label}>
        <img src={logoSrc} alt="" loading="lazy" decoding="async" onError={() => setLoadFailed(true)} />
      </span>
    );
  }

  if (mark) {
    return (
      <span
        className="connector-logo connector-logo-mark"
        aria-hidden="true"
        title={label}
        style={{ color: `#${mark.hex}` }}
      >
        <svg viewBox="0 0 24 24" role="presentation">
          <path d={mark.path} fill="currentColor" />
        </svg>
      </span>
    );
  }

  return (
    <span className={`connector-logo connector-logo-fallback connector-logo-${iconKey}`} aria-hidden="true" title={label}>
      {connectorGlyph(iconKey)}
    </span>
  );
}
