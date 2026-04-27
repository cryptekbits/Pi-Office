export const OFFICE_HOSTS = ["word", "excel", "powerpoint"] as const;
export type OfficeHost = (typeof OFFICE_HOSTS)[number];

export const PROMPT_MODES = ["prompt", "steer", "followUp"] as const;
export type PromptMode = (typeof PROMPT_MODES)[number];

export const OFFICE_DOCUMENT_STATES = ["unsaved", "saved"] as const;
export type OfficeDocumentState = (typeof OFFICE_DOCUMENT_STATES)[number];

export const COMPANION_STATUSES = ["discovering", "connected", "unavailable", "error"] as const;
export type CompanionStatus = (typeof COMPANION_STATUSES)[number];

export const CONNECTOR_EXECUTION_ENVIRONMENTS = ["browser", "companion"] as const;
export type ConnectorExecutionEnvironment = (typeof CONNECTOR_EXECUTION_ENVIRONMENTS)[number];

export interface CompanionCapabilities {
  fileRead: boolean;
  localMcp: boolean;
  endpoint?: string | undefined;
  shell?: CompanionShellCapability | undefined;
  version?: string | undefined;
  agent?: CompanionAgentCapability | undefined;
  providerAuth?: CompanionProviderAuthCapability | undefined;
  nativeCapture?: CompanionNativeCaptureCapability | undefined;
  mcp?: CompanionMcpCapability | undefined;
  memory?: CompanionSimpleCapability | undefined;
}

export interface CompanionDiscoveryAttempt {
  endpoint: string;
  ok: boolean;
  message?: string | undefined;
  durationMs?: number | undefined;
}

export interface CompanionState {
  status: CompanionStatus;
  endpoint?: string | undefined;
  identity?: string | undefined;
  lastError?: string | undefined;
  manualEndpoint?: string | undefined;
  lastSuccessfulEndpoint?: string | undefined;
  lastDiscoveryAttempts?: CompanionDiscoveryAttempt[] | undefined;
  sessionId?: string | undefined;
  connectorToolNames?: string[] | undefined;
  capabilities: CompanionCapabilities;
}

/**
 * @deprecated Use OfficeDocumentState plus CompanionState instead.
 */
export type OfficeMode = "document-only" | "workspace";

export interface PromptImagePayload {
  data: string;
  mimeType: string;
  label?: string | undefined;
  width?: number | undefined;
  height?: number | undefined;
  kind?: string | undefined;
}

export interface OfficeVisualSnapshot extends PromptImagePayload {
  kind: "selection" | "inline-picture" | "floating-shape" | "slide" | "shape" | "worksheet" | "viewport" | "window";
}

export interface OfficeSelectionFontSummary {
  name?: string | undefined;
  size?: number | undefined;
  bold?: boolean | undefined;
  italic?: boolean | undefined;
  color?: string | undefined;
}

export interface OfficeSelectionMeta {
  paragraphCount?: number | undefined;
  firstParagraphStyle?: string | undefined;
  styleHistogram?: Record<string, number> | undefined;
  isInTable?: boolean | undefined;
  isListItem?: boolean | undefined;
  enclosingParagraphText?: string | undefined;
  fontSummary?: OfficeSelectionFontSummary | undefined;
}

export interface OfficeSelectionSummary {
  label: string;
  textPreview?: string | undefined;
  structuredPreview?: string | undefined;
  selectionMeta?: OfficeSelectionMeta | undefined;
  details?: string[] | undefined;
  kind?: "empty" | "text" | "mixed" | "images" | "objects" | undefined;
  imageCount?: number | undefined;
  objectCount?: number | undefined;
}

export interface OfficeDocumentDescriptor {
  id: string;
  title: string;
  saved: boolean;
  documentUrl?: string | undefined;
  documentPath?: string | undefined;
  workspaceDir?: string | undefined;
}

export const OFFICE_ANCHOR_KINDS = [
  "document",
  "selection",
  "heading",
  "paragraph",
  "comment",
  "revision",
  "footnote",
  "endnote",
  "field",
  "contentControl",
  "searchResult",
  "bookmark",
  "hyperlink",
  "cell",
  "range",
  "sheet",
  "worksheet",
  "workbook",
  "slide",
  "notesRegion",
  "layout",
  "slideMaster",
  "shape",
  "table",
  "chart",
  "pivotTable",
  "namedItem",
] as const;
export type OfficeAnchorKind = (typeof OFFICE_ANCHOR_KINDS)[number];

export interface OfficeAnchor {
  kind: OfficeAnchorKind;
  label?: string | undefined;
  id?: string | undefined;
  text?: string | undefined;
  searchQuery?: string | undefined;
  sheetName?: string | undefined;
  address?: string | undefined;
  paragraphId?: string | undefined;
  searchResultId?: string | undefined;
  searchResultIndex?: number | undefined;
  objectType?: string | undefined;
  occurrenceIndex?: number | undefined;
  commentId?: string | undefined;
  revisionId?: string | undefined;
  noteTarget?: "reference" | "body" | undefined;
  bookmarkName?: string | undefined;
  hyperlinkId?: string | undefined;
  hyperlinkAddress?: string | undefined;
  slideId?: string | undefined;
  slideIndex?: number | undefined;
  shapeId?: string | undefined;
  tableName?: string | undefined;
  chartName?: string | undefined;
  pivotTableName?: string | undefined;
  namedItemName?: string | undefined;
}

export type OfficeCellValue = string | number | boolean | null;

export interface OfficeHostAction {
  type: string;
  target?: OfficeAnchor | undefined;
  content?: string | undefined;
  format?: string | undefined;
  placement?: string | undefined;
  values?: OfficeCellValue[][] | undefined;
  formulas?: string[][] | undefined;
  numberFormat?: string[][] | undefined;
  options?: Record<string, unknown> | undefined;
  [key: string]: unknown;
}

export interface OfficeObjectReference extends OfficeAnchor {
  host: OfficeHost;
}

export type OfficeActionCompletion = "native" | "fallback" | "partial";

export interface OfficeActionResult {
  ok: true;
  host: OfficeHost;
  operation: string;
  summary?: string | undefined;
  touchedObjects?: OfficeObjectReference[] | undefined;
  createdObjects?: OfficeObjectReference[] | undefined;
  navigation?: OfficeAnchor | undefined;
  completion?: OfficeActionCompletion | undefined;
  fallbackStrategy?: string | undefined;
  nativeAttempted?: boolean | undefined;
  nativeFailure?: string | undefined;
  warnings?: string[] | undefined;
  data?: Record<string, unknown> | undefined;
}

export interface OfficeApplyEditParams {
  action?: OfficeHostAction | undefined;
  mode?: string | undefined;
  content?: string | undefined;
  format?: string | undefined;
}

export interface OfficeNavigateParams {
  anchor?: OfficeAnchor | undefined;
  target?: string | undefined;
  kind?: OfficeAnchorKind | string | undefined;
}

export interface OfficeStateUpdate {
  host: OfficeHost;
  document: OfficeDocumentDescriptor;
  selection: OfficeSelectionSummary;
  capabilities: string[];
  timestamp: string;
}

export interface OfficeSessionOpenRequest {
  host: OfficeHost;
  documentId: string;
  documentPath?: string | undefined;
  documentUrl?: string | undefined;
  saved: boolean;
  title: string;
  selectionSummary?: OfficeSelectionSummary | undefined;
  forceNew?: boolean | undefined;
  windowId?: string | undefined;
}

export interface OfficeSessionOpenResponse {
  sessionId: string;
  documentState: OfficeDocumentState;
  companion: CompanionState;
  origin: string;
  eventsPath: string;
}

export interface OfficeSessionStateResponse {
  ok: true;
  documentState: OfficeDocumentState;
  companion: CompanionState;
}

export interface CompanionSessionOpenRequest {
  browserSessionId: string;
  windowId?: string | undefined;
  host: OfficeHost;
  documentId: string;
  documentPath?: string | undefined;
  documentUrl?: string | undefined;
  saved: boolean;
  title: string;
  connectors: CompanionConnectorDefinition[];
}

export interface CompanionSessionOpenResponse {
  ok: true;
  sessionId: string;
  companion: CompanionState;
  connectors: ConnectorStatus[];
}

export const COMPANION_CAPABILITY_STATES = ["unavailable", "available", "degraded"] as const;
export type CompanionCapabilityState = (typeof COMPANION_CAPABILITY_STATES)[number];

export interface CompanionSimpleCapability {
  state: CompanionCapabilityState;
  available: boolean;
  reason?: string | undefined;
  version?: string | undefined;
}

export interface CompanionAgentCapability extends CompanionSimpleCapability {
  officeToolProxy: boolean;
  providerAuth: boolean;
  smartAuto: boolean;
}

export interface CompanionProviderAuthCapability extends CompanionSimpleCapability {
  explicitMigrationRequired: boolean;
  supportedAuthMethods?: ProviderAuthMethod[] | undefined;
}

export interface CompanionMcpCapability extends CompanionSimpleCapability {
  readOnly: boolean;
  toolCount?: number | undefined;
}

export interface CompanionNativeCaptureCapability extends CompanionSimpleCapability {
  hosts: OfficeHost[];
  platform?: string | undefined;
  trueViewportScreenshot: boolean;
  includeWindowFrame: boolean;
}

export interface CompanionNativeCaptureRequest {
  host?: OfficeHost | undefined;
  includeWindowFrame?: boolean | undefined;
}

export interface CompanionNativeCaptureResponse {
  ok: boolean;
  visual?: OfficeVisualSnapshot | undefined;
  details?: Record<string, unknown> | undefined;
  error?: string | undefined;
}

export const COMPANION_SHELL_STATES = ["unavailable", "available", "degraded"] as const;
export type CompanionShellState = (typeof COMPANION_SHELL_STATES)[number];

export type CompanionShellBackend =
  | "none"
  | "docker"
  | "wsl2"
  | "linux-bubblewrap"
  | "macos-container"
  | "test";

export type CompanionShellNetworkPolicy = "disabled" | "enabled";

export interface CompanionShellPolicy {
  version: string;
  readableRoots: string[];
  scratchRoot?: string | undefined;
  deniedPatterns: string[];
  network: CompanionShellNetworkPolicy;
  timeoutMs: number;
  outputByteLimit: number;
}

export interface CompanionShellProbeResult {
  id: string;
  description: string;
  ok: boolean;
  blocked: boolean;
  detail: string;
}

export interface CompanionShellCapability {
  state: CompanionShellState;
  backend: CompanionShellBackend;
  detectedAt: string;
  reason?: string | undefined;
  policy: CompanionShellPolicy;
  probes: CompanionShellProbeResult[];
}

export interface CompanionShellExecuteRequest {
  command: string;
  cwd?: string | undefined;
  category?: "read-only" | "scratch-write" | undefined;
  approvalId?: string | undefined;
  timeoutMs?: number | undefined;
}

export interface CompanionShellExecuteResponse {
  ok: boolean;
  auditId: string;
  backend: CompanionShellBackend;
  policyVersion: string;
  durationMs: number;
  exitCode?: number | null | undefined;
  stdout: string;
  stderr: string;
  capped: boolean;
  deniedPaths: string[];
  touchedPaths: string[];
  error?: string | undefined;
}

export interface PromptRequest {
  text: string;
  mode?: PromptMode | undefined;
  images?: PromptImagePayload[] | undefined;
}

export type PromptSuggestionRole = "user" | "assistant";

export interface PromptSuggestionMessage {
  role: PromptSuggestionRole;
  text: string;
}

export interface PromptSuggestion {
  id: string;
  text: string;
}

export interface PromptSuggestionRequest {
  generationId: string;
  latestAssistantText: string;
  recentMessages: PromptSuggestionMessage[];
}

export interface PromptSuggestionResponse {
  generationId: string;
  suggestions: PromptSuggestion[];
}

export interface DeriveSubjectRequest {
  messages: { role: string; text: string }[];
}

export interface DeriveSubjectResponse {
  subject: string;
}

export interface SetModelRequest {
  provider: string;
  modelId: string;
}

export interface SaveApiKeyRequest {
  provider: string;
  apiKey: string;
}

export const CONNECTOR_CATEGORIES = [
  "knowledge",
  "collaboration",
  "productivity",
  "engineering",
  "analytics",
  "commerce",
  "research",
  "data",
] as const;
export type ConnectorCategory = (typeof CONNECTOR_CATEGORIES)[number];

export const CONNECTOR_SETUP_DIFFICULTIES = ["easy", "guided", "advanced"] as const;
export type ConnectorSetupDifficulty = (typeof CONNECTOR_SETUP_DIFFICULTIES)[number];

export const CONNECTOR_SETUP_KINDS = [
  "remote_oauth",
  "remote_api_key",
  "remote_url_token",
  "local_node",
  "local_python",
  "local_executable_or_docker",
] as const;
export type ConnectorSetupKind = (typeof CONNECTOR_SETUP_KINDS)[number];

export const CONNECTOR_AUTH_METHODS = [
  "none",
  "api_key",
  "bearer_token",
  "oauth",
] as const;
export type ConnectorAuthMethod = (typeof CONNECTOR_AUTH_METHODS)[number];

export const CONNECTOR_TRANSPORTS = ["remote_http", "local_stdio"] as const;
export type ConnectorTransport = (typeof CONNECTOR_TRANSPORTS)[number];

export const CONNECTOR_MATURITY_LEVELS = ["ready", "beta", "custom_mcp_only"] as const;
export type ConnectorMaturity = (typeof CONNECTOR_MATURITY_LEVELS)[number];

export const CONNECTOR_CREDENTIAL_SOURCES = [
  "none",
  "env",
  "manual",
  "detected_env",
  "oauth",
] as const;
export type ConnectorCredentialSource = (typeof CONNECTOR_CREDENTIAL_SOURCES)[number];

export const CONNECTOR_SCOPE_TARGETS = ["global", "workspace", "document"] as const;
export type ConnectorScopeTarget = (typeof CONNECTOR_SCOPE_TARGETS)[number];

export const CONNECTOR_HEALTH_STATES = [
  "ready",
  "stale",
  "offline",
  "auth_required",
  "auth_expired",
  "degraded",
  "unverified",
] as const;
export type ConnectorHealthState = (typeof CONNECTOR_HEALTH_STATES)[number];

export const CONNECTOR_DIAGNOSTIC_LEVELS = ["info", "warning", "error"] as const;
export type ConnectorDiagnosticLevel = (typeof CONNECTOR_DIAGNOSTIC_LEVELS)[number];

export const CONNECTOR_RUNTIME_CHECKS = [
  "node",
  "npm",
  "npx",
  "python",
  "uv",
  "docker",
  "git",
] as const;
export type ConnectorRuntimeCheckKey = (typeof CONNECTOR_RUNTIME_CHECKS)[number];

export interface ConnectorRuntimeRequirement {
  key: ConnectorRuntimeCheckKey;
  label: string;
  optional?: boolean | undefined;
  description?: string | undefined;
}

export interface ConnectorEnvHint {
  key: string;
  label: string;
  description?: string | undefined;
  required?: boolean | undefined;
}

/** Static HTTP header for remote MCP (applied by the companion). */
export interface ConnectorRemoteHttpHeader {
  name: string;
  value: string;
}

/** HTTP header whose value is read from the companion process environment. */
export interface ConnectorRemoteHttpHeaderFromEnv {
  name: string;
  envVarName: string;
}

export interface ConnectorReadPolicy {
  mode: "hard-read-only";
  allowResources: boolean;
  allowPrompts: boolean;
  allowToolPatterns: string[];
  blockToolPatterns: string[];
  allowPromptPatterns?: string[] | undefined;
  blockPromptPatterns?: string[] | undefined;
}

export interface ConnectorConfigTemplate {
  transport: ConnectorTransport;
  command?: string | undefined;
  args?: string[] | undefined;
  cwd?: string | undefined;
  url?: string | undefined;
  env?: Record<string, string> | undefined;
}

export type ConnectorSetupProfileOfficialness =
  | "official"
  | "official_preview"
  | "community"
  | "provider_reference"
  | "deprecated"
  | "experimental"
  | "planned";

export type ConnectorSetupProfileAvailability = "available" | "needs_companion" | "planned" | "advanced";
export type ConnectorBrowserDirectSupport = "supported" | "unsupported" | "unknown";
export type ConnectorOAuthBroker = "taskpane" | "companion";
export type ConnectorOAuthLaunchMode = "popup" | "system_browser";

export interface ConnectorOAuthProfileSettings {
  broker: ConnectorOAuthBroker;
  launchMode?: ConnectorOAuthLaunchMode | undefined;
  metadataUrl?: string | undefined;
  authorizationUrl?: string | undefined;
  tokenUrl?: string | undefined;
  registrationUrl?: string | undefined;
  redirectPath?: string | undefined;
  clientName?: string | undefined;
  scopes?: string[] | undefined;
}

export interface ConnectorSetupProfile {
  id: string;
  label: string;
  description?: string | undefined;
  transport: ConnectorTransport;
  setupKind?: ConnectorSetupKind | undefined;
  authMethod: ConnectorAuthMethod;
  /** Hosted MCP endpoint for remote profiles. */
  endpoint?: string | undefined;
  command?: string | undefined;
  args?: string[] | undefined;
  cwd?: string | undefined;
  env?: Record<string, string> | undefined;
  remoteHttpHeaders?: ConnectorRemoteHttpHeader[] | undefined;
  remoteHttpHeadersFromEnv?: ConnectorRemoteHttpHeaderFromEnv[] | undefined;
  credentialEnvKey?: string | undefined;
  requiresCompanion: boolean;
  officialness: ConnectorSetupProfileOfficialness;
  availability?: ConnectorSetupProfileAvailability | undefined;
  browserDirect?: ConnectorBrowserDirectSupport | undefined;
  oauth?: ConnectorOAuthProfileSettings | undefined;
  docsUrl?: string | undefined;
  endpointEvidenceUrl?: string | undefined;
  authEvidenceUrl?: string | undefined;
  checkedAt?: string | undefined;
  setupDisabled?: boolean | undefined;
  riskNotes?: string[] | undefined;
  defaultWhenCompanionAbsent?: boolean | undefined;
  defaultWhenCompanionPresent?: boolean | undefined;
  privacyNotes?: string[] | undefined;
  simpleFields: string[];
  advancedFields: string[];
}

export const CONNECTOR_TOOL_CLASSIFICATIONS = [
  "read",
  "sensitive_read",
  "costly_read",
  "write",
  "destructive",
  "unknown",
] as const;
export type ConnectorToolClassification = (typeof CONNECTOR_TOOL_CLASSIFICATIONS)[number];

export interface ConnectorMcpToolAnnotations {
  title?: string | undefined;
  readOnlyHint?: boolean | undefined;
  destructiveHint?: boolean | undefined;
  idempotentHint?: boolean | undefined;
  openWorldHint?: boolean | undefined;
}

export interface ConnectorToolInventoryItem {
  name: string;
  rawName?: string | undefined;
  description?: string | undefined;
  inputSchema?: unknown;
  annotations?: ConnectorMcpToolAnnotations | undefined;
  classification: ConnectorToolClassification;
  defaultEnabled: boolean;
  enabled: boolean;
  reason: string;
  source: "mcp_tool" | "resource" | "catalog_hint";
}

export interface ConnectorToolPolicyOverride {
  toolName: string;
  enabled: boolean;
  warningAcknowledged?: boolean | undefined;
  updatedAt?: string | undefined;
}

export interface ConnectorCatalogItem {
  id: string;
  name: string;
  vendor: string;
  iconKey: string;
  category: ConnectorCategory;
  maturity: ConnectorMaturity;
  setupKind: ConnectorSetupKind;
  authMethod: ConnectorAuthMethod;
  transport: ConnectorTransport;
  readOnly: boolean;
  customOnly?: boolean | undefined;
  summary: string;
  officeValue: string;
  tags: string[];
  capabilityHints: string[];
  requirements: ConnectorRuntimeRequirement[];
  envHints: ConnectorEnvHint[];
  readPolicy: ConnectorReadPolicy;
  template?: ConnectorConfigTemplate | undefined;
  setupProfiles?: ConnectorSetupProfile[] | undefined;
  docsUrl?: string | undefined;
  authUrl?: string | undefined;
  setupNotes?: string[] | undefined;
  recommendedHosts: OfficeHost[];
  setupDifficulty: ConnectorSetupDifficulty;
  catalogRevision?: string | undefined;
}

export interface CompanionConnectorDefinition {
  id: string;
  connectorId: string;
  name: string;
  source: "library" | "custom";
  category: ConnectorCategory;
  maturity: ConnectorMaturity;
  setupKind: ConnectorSetupKind;
  authMethod: ConnectorAuthMethod;
  transport: ConnectorTransport;
  credentialSource: ConnectorCredentialSource;
  url?: string | undefined;
  command?: string | undefined;
  args?: string[] | undefined;
  cwd?: string | undefined;
  setupProfileId?: string | undefined;
  env?: Record<string, string> | undefined;
  /** Extra host env var names copied into the stdio child (companion), in addition to SDK safe inherited vars. */
  stdioEnvPassthrough?: string[] | undefined;
  /** Static request headers for Streamable HTTP / SSE MCP (companion). */
  remoteHttpHeaders?: ConnectorRemoteHttpHeader[] | undefined;
  /** Request headers populated from companion environment variable values. */
  remoteHttpHeadersFromEnv?: ConnectorRemoteHttpHeaderFromEnv[] | undefined;
  /** OAuth broker metadata copied from the selected setup profile. */
  oauth?: ConnectorOAuthProfileSettings | undefined;
  secret?: string | undefined;
  secretEnvKey?: string | undefined;
  useDetectedEnvKey?: string | undefined;
  readOnly: boolean;
  readPolicy: ConnectorReadPolicy;
  toolPolicyOverrides?: ConnectorToolPolicyOverride[] | undefined;
  catalogRevision?: string | undefined;
}

export interface ConnectorCapabilitySummary {
  tools: string[];
  allowedTools: string[];
  blockedTools: string[];
  toolInventory?: ConnectorToolInventoryItem[] | undefined;
  resourceToolNames: string[];
  promptNames: string[];
  allowedPrompts: string[];
  blockedPrompts: string[];
  resourceCount: number;
  promptCount: number;
  inventoryHash?: string | undefined;
}

export interface ConnectorScopeContext {
  host: OfficeHost;
  documentId?: string | undefined;
  documentTitle?: string | undefined;
  documentSaved: boolean;
  documentUrl?: string | undefined;
  workspaceId?: string | undefined;
}

export interface ConnectorScopeState {
  target: ConnectorScopeTarget;
  enabled: boolean;
  applies: boolean;
  label: string;
  scopeKey?: string | undefined;
  reason?: string | undefined;
}

export interface ConnectorVerificationSnapshot {
  verifiedAt: string;
  catalogRevision: string;
  inventoryHash: string;
  toolNames: string[];
  allowedTools: string[];
  blockedTools: string[];
  toolInventory?: ConnectorToolInventoryItem[] | undefined;
  resourceToolNames: string[];
  promptNames: string[];
  allowedPrompts: string[];
  blockedPrompts: string[];
  resourceCount: number;
  promptCount: number;
  staleReason?: string | undefined;
  lastFailure?: string | undefined;
}

export interface ConnectorConflict {
  key: string;
  kind: "duplicate_connector" | "duplicate_remote_url" | "duplicate_local_runtime" | "import_collision";
  existingConnectorId: string;
  existingName: string;
  incomingConnectorId: string;
  title: string;
  message: string;
  resolution: "open_existing" | "replace" | "skip";
}

export interface ConnectorLogSummary {
  total: number;
  lastAt?: string | undefined;
  lastLevel?: ConnectorDiagnosticLevel | undefined;
}

export interface ConnectorAuditPreference {
  enabled: boolean;
}

export interface ConnectorLogEntry {
  id: string;
  connectorId?: string | undefined;
  connectorName?: string | undefined;
  kind: "setup" | "test" | "reverify" | "scope" | "favorite" | "remove" | "import" | "export" | "use";
  channel: "diagnostic" | "audit";
  level: ConnectorDiagnosticLevel;
  message: string;
  timestamp: string;
  scopeTarget?: ConnectorScopeTarget | undefined;
  healthState?: ConnectorHealthState | undefined;
  durationMs?: number | undefined;
  errorCode?: string | undefined;
  redacted?: boolean | undefined;
}

export interface ConnectorStatus {
  id: string;
  connectorId: string;
  name: string;
  enabled: boolean;
  configured: boolean;
  connected: boolean;
  source: "library" | "custom";
  category: ConnectorCategory;
  maturity: ConnectorMaturity;
  setupKind: ConnectorSetupKind;
  authMethod: ConnectorAuthMethod;
  transport: ConnectorTransport;
  setupProfileId?: string | undefined;
  credentialSource: ConnectorCredentialSource;
  detectedEnvKey?: string | undefined;
  usesDetectedCredential?: boolean | undefined;
  lastTestedAt?: string | undefined;
  lastError?: string | undefined;
  lastHealthyAt?: string | undefined;
  healthState: ConnectorHealthState;
  staleReason?: string | undefined;
  activeScope: ConnectorScopeTarget;
  scopeStates: ConnectorScopeState[];
  favorite: boolean;
  recommendedHosts: OfficeHost[];
  setupDifficulty: ConnectorSetupDifficulty;
  needsCredential: boolean;
  verification?: ConnectorVerificationSnapshot | undefined;
  conflicts?: ConnectorConflict[] | undefined;
  logSummary?: ConnectorLogSummary | undefined;
  capabilities?: ConnectorCapabilitySummary | undefined;
  executionEnvironment?: ConnectorExecutionEnvironment | undefined;
  executionAvailable?: boolean | undefined;
  suppressNonReadToolWarning?: boolean | undefined;
}

export interface ConnectorRuntimeCheck {
  key: ConnectorRuntimeCheckKey;
  label: string;
  ok: boolean;
  detail: string;
}

export interface ConnectorEnvSuggestion {
  key: string;
  present: boolean;
  maskedValue?: string | undefined;
}

export interface ConnectorDiagnostic {
  level: ConnectorDiagnosticLevel;
  code: string;
  title: string;
  message: string;
  connectorId?: string | undefined;
}

export interface ConnectorCatalogResponse {
  connectors: ConnectorCatalogItem[];
}

export interface ConnectorStatusResponse {
  connectors: ConnectorStatus[];
}

export interface ConnectorDiagnosticsResponse {
  generatedAt: string;
  runtimes: ConnectorRuntimeCheck[];
  envSuggestions: ConnectorEnvSuggestion[];
  diagnostics: ConnectorDiagnostic[];
}

export interface ConnectorPrepareResponse {
  connector: ConnectorCatalogItem;
  existing?: ConnectorStatus | undefined;
  draft?: ConnectorSetupRequest | undefined;
  runtimes: ConnectorRuntimeCheck[];
  envSuggestions: ConnectorEnvSuggestion[];
  diagnostics: ConnectorDiagnostic[];
  conflicts?: ConnectorConflict[] | undefined;
  auditPreference?: ConnectorAuditPreference | undefined;
  executionEnvironment?: ConnectorExecutionEnvironment | undefined;
  executionAvailable?: boolean | undefined;
}

export interface ConnectorSetupRequest {
  connectorId: string;
  existingId?: string | undefined;
  name?: string | undefined;
  enabled?: boolean | undefined;
  setupKind?: ConnectorSetupKind | undefined;
  authMethod?: ConnectorAuthMethod | undefined;
  transport?: ConnectorTransport | undefined;
  setupProfileId?: string | undefined;
  credentialSource?: ConnectorCredentialSource | undefined;
  secret?: string | undefined;
  secretEnvKey?: string | undefined;
  useDetectedEnvKey?: string | undefined;
  url?: string | undefined;
  command?: string | undefined;
  args?: string[] | undefined;
  cwd?: string | undefined;
  env?: Record<string, string> | undefined;
  stdioEnvPassthrough?: string[] | undefined;
  remoteHttpHeaders?: ConnectorRemoteHttpHeader[] | undefined;
  remoteHttpHeadersFromEnv?: ConnectorRemoteHttpHeaderFromEnv[] | undefined;
  preserveStoredSecret?: boolean | undefined;
  replaceExisting?: boolean | undefined;
  favorite?: boolean | undefined;
  scopeTarget?: ConnectorScopeTarget | undefined;
  scopeContext?: ConnectorScopeContext | undefined;
}

export interface ConnectorSetupResponse {
  ok: true;
  status: ConnectorStatus;
  diagnostics: ConnectorDiagnostic[];
  conflicts?: ConnectorConflict[] | undefined;
}

export interface ConnectorTestResponse {
  ok: boolean;
  status: ConnectorStatus;
  diagnostics: ConnectorDiagnostic[];
  conflicts?: ConnectorConflict[] | undefined;
}

export interface ConnectorOAuthStartRequest {
  connectorId: string;
}

export interface ConnectorOAuthStartResponse {
  ok: true;
  connectorId: string;
  url?: string | undefined;
  callbackUrl?: string | undefined;
  state: string;
  expiresAt: string;
  broker?: ConnectorOAuthBroker | undefined;
  openMode?: ConnectorOAuthLaunchMode | undefined;
}

export interface CompanionConnectorOAuthStartRequest {
  definition: CompanionConnectorDefinition;
}

export interface CompanionConnectorOAuthStatusRequest {
  connectorId?: string | undefined;
  state?: string | undefined;
}

export interface CompanionConnectorOAuthStatusResponse {
  ok: true;
  connectorId?: string | undefined;
  state?: string | undefined;
  connected: boolean;
  pending: boolean;
  expiresAt?: string | undefined;
  error?: string | undefined;
  status?: ConnectorStatus | undefined;
  diagnostics?: ConnectorDiagnostic[] | undefined;
}

export interface ConnectorOAuthCredentialHandoff {
  accessToken: string;
  refreshToken?: string | undefined;
  tokenType?: string | undefined;
  scope?: string | undefined;
  expiresAt?: string | undefined;
  expiresInSeconds?: number | undefined;
}

export interface ConnectorOAuthCallbackRequest {
  connectorId?: string | undefined;
  state: string;
  code?: string | undefined;
  error?: string | undefined;
  credential?: ConnectorOAuthCredentialHandoff | undefined;
  expiresAt?: string | undefined;
  expiresInSeconds?: number | undefined;
}

export interface ConnectorOAuthCallbackResponse {
  ok: true;
  status: ConnectorStatus;
  diagnostics: ConnectorDiagnostic[];
}

export interface ConnectorScopeUpdateRequest {
  connectorId: string;
  enabled: boolean;
  scopeTarget: ConnectorScopeTarget;
  scopeContext?: ConnectorScopeContext | undefined;
}

export interface ConnectorToolPolicyUpdateRequest {
  connectorId: string;
  toolName: string;
  enabled: boolean;
  warningAcknowledged?: boolean | undefined;
  suppressWarning?: boolean | undefined;
  scopeContext?: ConnectorScopeContext | undefined;
}

export interface ConnectorToolPolicyUpdateResponse {
  ok: true;
  status: ConnectorStatus;
}

export interface ConnectorFavoriteRequest {
  connectorId: string;
  favorite: boolean;
}

export interface ConnectorLogResponse {
  connectorId: string;
  auditPreference: ConnectorAuditPreference;
  logs: ConnectorLogEntry[];
}

export interface ConnectorAuditPreferenceResponse {
  preference: ConnectorAuditPreference;
}

export interface ConnectorExportItem {
  storedId?: string | undefined;
  connectorId: string;
  name: string;
  source: "library" | "custom";
  category: ConnectorCategory;
  maturity: ConnectorMaturity;
  setupKind: ConnectorSetupKind;
  authMethod: ConnectorAuthMethod;
  transport: ConnectorTransport;
  setupProfileId?: string | undefined;
  credentialSource: ConnectorCredentialSource;
  secretEnvKey?: string | undefined;
  useDetectedEnvKey?: string | undefined;
  url?: string | undefined;
  command?: string | undefined;
  args?: string[] | undefined;
  cwd?: string | undefined;
  env?: Record<string, string> | undefined;
  stdioEnvPassthrough?: string[] | undefined;
  remoteHttpHeaders?: ConnectorRemoteHttpHeader[] | undefined;
  remoteHttpHeadersFromEnv?: ConnectorRemoteHttpHeaderFromEnv[] | undefined;
  toolPolicyOverrides?: ConnectorToolPolicyOverride[] | undefined;
  suppressNonReadToolWarning?: boolean | undefined;
  defaultEnabled: boolean;
}

export interface ConnectorExportScopeOverride {
  storedId?: string | undefined;
  connectorId: string;
  scopeTarget: Exclude<ConnectorScopeTarget, "global">;
  scopeKey: string;
  enabled: boolean;
}

export interface ConnectorExportBundle {
  version: number;
  exportedAt: string;
  connectors: ConnectorExportItem[];
  scopeOverrides: ConnectorExportScopeOverride[];
  favorites: string[];
  auditPreference: ConnectorAuditPreference;
}

export interface ConnectorImportPreviewRequest {
  bundle: ConnectorExportBundle;
}

export interface ConnectorImportPreviewResponse {
  ok: true;
  bundle: ConnectorExportBundle;
  conflicts: ConnectorConflict[];
}

export interface ConnectorImportApplyRequest {
  bundle: ConnectorExportBundle;
  resolutions?: Record<string, "skip" | "replace"> | undefined;
}

export interface ConnectorImportApplyResponse {
  ok: true;
  importedConnectorIds: string[];
  skippedConflictKeys: string[];
}

export interface ProviderModelDescriptor {
  provider: string;
  providerLabel: string;
  modelId: string;
  modelName: string;
  settingsVisibility: SettingsVisibility;
  lab: string;
  family?: string | undefined;
  recommended: boolean;
  recommendationReason?: string | undefined;
  defaultForProvider: boolean;
  requiresUnrecommendedWarning: boolean;
  supportStatus: ProviderSupportStatus;
  runtimeSurface: ProviderRuntimeSurface;
  authMethods: ProviderAuthMethod[];
  apiKeySupported: boolean;
  browserCallable: boolean;
  companionRequired: boolean;
  subscriptionBacked: boolean;
  imageGenerationSupported: boolean;
  capabilityNote?: string | undefined;
  authState: ProviderAuthState;
  credentialStored: boolean;
  verifiedUsable: boolean;
  verificationError?: string | undefined;
  verifiedAt?: string | undefined;
  configured: boolean;
  oauthSupported: boolean;
  usesApiKey: boolean;
  contextWindow?: number | undefined;
  costTier?: "$" | "$$" | "$$$" | undefined;
  supportsThinking?: boolean | undefined;
  supportsReasoningEffort?: boolean | undefined;
}

export const THINKING_LEVELS = ["off", "minimal", "low", "medium", "high", "xhigh"] as const;
export type ThinkingLevel = (typeof THINKING_LEVELS)[number];

export const AUTONOMY_LEVELS = ["off", "low", "medium", "high", "extreme"] as const;
export type AutonomyLevel = (typeof AUTONOMY_LEVELS)[number];

export type ToolCategory =
  | "read"
  | "write-doc"
  | "escape-hatch"
  | "write-external"
  | "read-external"
  | "connector"
  | "interaction";

export interface ToolPermissionOverride {
  toolName: string;
  autoApproveAtLevel: AutonomyLevel | "disabled";
}

export interface ToolPermissionDecision {
  toolName: string;
  allowed: boolean;
  scope: "once" | "session" | "workspace" | "always";
}

export interface ToolPermissionRequest {
  requestId: string;
  toolName: string;
  toolCategory: ToolCategory;
  params: Record<string, unknown>;
  expiresAt?: string | undefined;
}

export const TOOL_CATEGORY_MAP: Record<string, ToolCategory> = {
  office_get_context: "read",
  office_tool_search: "read",
  office_tool_get: "read",
  office_tool_call: "write-doc",
  office_batch_execute: "write-doc",
  mcp_tool_search: "read",
  mcp_batch_execute: "connector",
  mcp_result_get: "read",
  mcp_result_summarize: "read",
  mcp_result_clear: "connector",
  office_read_section: "read",
  word_search: "read",
  word_format_text: "write-doc",
  word_list_format: "write-doc",
  word_reference_inventory: "read",
  word_hyperlink: "write-doc",
  word_table: "write-doc",
  word_section_layout: "write-doc",
  word_equation: "write-doc",
  word_field_reference: "write-doc",
  word_content_control: "write-doc",
  word_building_block: "write-doc",
  word_annotation_review: "write-doc",
  word_redline_review: "write-doc",
  word_proofing_stats: "read",
  word_collab_guard: "read",
  office_capture_snapshot: "read",
  office_capture_viewport: "read",
  verify_doc: "read",
  verify_doc_visual: "read",
  get_cell_ranges: "read",
  get_all_objects: "read",
  search_data: "read",
  get_range_as_csv: "read",
  read_range_image: "read",
  extract_chart_xml: "read",
  get_presentation_structure: "read",
  get_slide: "read",
  list_slide_shapes: "read",
  search_icons: "read",
  verify_slides: "read",
  verify_slide_visual: "read",
  office_apply_edit: "write-doc",
  edit_doc_text: "write-doc",
  edit_doc_list: "write-doc",
  set_cell_range: "write-doc",
  clear_cell_range: "write-doc",
  resize_range: "write-doc",
  copy_to: "write-doc",
  modify_sheet_structure: "write-doc",
  modify_object: "write-doc",
  modify_presentation_structure: "write-doc",
  duplicate_slide: "write-doc",
  insert_slide_element: "write-doc",
  remove_slide_element: "write-doc",
  edit_slide_text: "write-doc",
  edit_slide_xml: "write-doc",
  edit_slide_master: "write-doc",
  edit_slide_chart: "write-doc",
  copy_image_between_slides: "write-doc",
  insert_icon: "write-doc",
  office_propose_edits: "write-doc",
  office_navigate: "write-doc",
  office_execute_js: "escape-hatch",
  generate_image: "write-doc",
  ask_user: "interaction",
  mcp: "connector",
  read: "read-external",
  grep: "read-external",
  find: "read-external",
  ls: "read-external",
  edit: "write-external",
  write: "write-external",
  bash: "write-external",
};

export type ToolCapabilityRuntime = "taskpane" | "companion" | "browser" | "office-js";
export type ToolCapabilityRisk = "low" | "medium" | "high" | "escape_hatch";
export type ToolCapabilitySupportStatus = "supported" | "desktop_only" | "preview" | "companion_required" | "fallback" | "unsupported";

export interface ToolCapabilityRequirement {
  name: string;
  version?: string | undefined;
  status?: ToolCapabilitySupportStatus | undefined;
  note?: string | undefined;
}

export interface ToolCapabilitySearchResult {
  id: string;
  toolName: string;
  label: string;
  description: string;
  host?: OfficeHost | "all" | undefined;
  category: ToolCategory;
  runtime: ToolCapabilityRuntime;
  risk: ToolCapabilityRisk;
  support: ToolCapabilitySupportStatus;
  requirements?: ToolCapabilityRequirement[] | undefined;
  keywords: string[];
  score?: number | undefined;
  source?: string | undefined;
  connectorId?: string | undefined;
  connectorName?: string | undefined;
  schemaAvailable: boolean;
  fallback?: string | undefined;
}

export interface ToolCapabilityDetail extends ToolCapabilitySearchResult {
  parameters?: unknown;
}

export interface OfficeToolSearchRequest {
  query?: string | undefined;
  host?: OfficeHost | undefined;
  category?: ToolCategory | undefined;
  limit?: number | undefined;
  includeUnsupported?: boolean | undefined;
}

export interface OfficeToolSearchResponse {
  host: OfficeHost;
  query?: string | undefined;
  results: ToolCapabilitySearchResult[];
}

export interface OfficeToolGetRequest {
  toolName?: string | undefined;
  id?: string | undefined;
}

export interface OfficeToolGetResponse {
  detail: ToolCapabilityDetail;
}

export interface OfficeToolCallRequest {
  toolName: OfficeToolName;
  arguments?: Record<string, unknown> | undefined;
}

export interface McpToolSearchRequest {
  query?: string | undefined;
  connectorId?: string | undefined;
  limit?: number | undefined;
}

export interface McpToolSearchResponse {
  query?: string | undefined;
  results: ToolCapabilitySearchResult[];
}

export type BatchStepStatus = "completed" | "blocked" | "failed";

export interface StructuredBatchStep {
  id?: string | undefined;
  type: string;
  toolName?: string | undefined;
  arguments?: Record<string, unknown> | undefined;
  query?: string | undefined;
  limit?: number | undefined;
}

export interface StructuredBatchStepResult {
  id?: string | undefined;
  index: number;
  type: string;
  toolName?: string | undefined;
  status: BatchStepStatus;
  summary: string;
  content?: unknown;
  error?: string | undefined;
}

export interface StructuredBatchRequest {
  steps: StructuredBatchStep[];
  outputByteLimit?: number | undefined;
}

export interface StructuredBatchResponse {
  ok: boolean;
  summary: string;
  steps: StructuredBatchStepResult[];
  blocked: number;
  failed: number;
}

export type OfficeBatchStep = StructuredBatchStep;
export type OfficeBatchStepResult = StructuredBatchStepResult;
export interface OfficeBatchExecuteRequest extends StructuredBatchRequest {}
export interface OfficeBatchExecuteResponse extends StructuredBatchResponse {}
export type McpBatchStep = StructuredBatchStep;
export type McpBatchStepResult = StructuredBatchStepResult;
export interface McpBatchExecuteRequest extends StructuredBatchRequest {}
export interface McpBatchExecuteResponse extends StructuredBatchResponse {}
export type BatchExecutionResult = StructuredBatchResponse;

export interface McpResultHandle {
  handleId: string;
  source: "browser" | "companion";
  connectorId?: string | undefined;
  connectorName?: string | undefined;
  toolName?: string | undefined;
  createdAt: string;
  expiresAt?: string | undefined;
  sizeBytes: number;
  pageSizeBytes: number;
  pageCount: number;
  summary: string;
  redacted: boolean;
}

export interface McpResultPageRequest {
  handleId: string;
  page?: number | undefined;
}

export interface McpResultPageResponse {
  ok: true;
  handle: McpResultHandle;
  page: number;
  content: string;
  hasNextPage: boolean;
}

export interface McpResultSummarizeRequest {
  handleId: string;
  query?: string | undefined;
  maxChars?: number | undefined;
}

export interface McpResultSummarizeResponse {
  ok: true;
  handle: McpResultHandle;
  summary: string;
}

export interface McpResultClearRequest {
  handleId?: string | undefined;
}

export interface McpResultClearResponse {
  ok: true;
  cleared: number;
}

export const AUTONOMY_LEVEL_AUTO_APPROVE: Record<AutonomyLevel, Set<ToolCategory>> = {
  off: new Set(["interaction"]),
  low: new Set(["interaction", "read"]),
  medium: new Set(["interaction", "read", "write-doc"]),
  high: new Set(["interaction", "read", "write-doc", "read-external"]),
  extreme: new Set(["interaction", "read", "write-doc", "read-external", "write-external", "connector"]),
};

export const AUTONOMY_LEVEL_LABELS: Record<AutonomyLevel, string> = {
  off: "Off",
  low: "Low",
  medium: "Medium",
  high: "High",
  extreme: "Full",
};

export const IMAGE_REASONING_EFFORTS = ["low", "medium", "high"] as const;
export type ImageReasoningEffort = (typeof IMAGE_REASONING_EFFORTS)[number];

export const IMAGE_API_TYPES = ["openai-images", "openai-chat-image", "google-generative-ai-image"] as const;
export type ImageApiType = (typeof IMAGE_API_TYPES)[number];

export interface ImageModelDescriptor {
  provider: string;
  modelId: string;
  modelName: string;
  apiType: ImageApiType;
  authState: ProviderAuthState;
  credentialStored: boolean;
  verifiedUsable: boolean;
  verificationError?: string | undefined;
  verifiedAt?: string | undefined;
  supportsReasoningEffort: boolean;
  supportedAspectRatios: string[];
  supportedSizes: string[];
  configured: boolean;
}

export interface ImageModelCatalogResponse {
  models: ImageModelDescriptor[];
  defaultModelKey: string;
}

export interface ModelSessionSettings {
  thinkingLevel?: ThinkingLevel | undefined;
}

export interface ThinkingCapabilities {
  supportsThinking: boolean;
  availableLevels: ThinkingLevel[];
  supportsXhigh: boolean;
  currentLevel: ThinkingLevel;
}

export interface UserPreferences {
  showThinkingTraces: boolean;
  autoAttachVisuals: boolean;
  showTokenUsage: boolean;
  compactMessages: boolean;
  defaultThinkingLevel: ThinkingLevel;
  imageGenerationEnabled: boolean;
  defaultImageModel: string;
  defaultModelByProvider: Record<string, string>;
  suppressUnrecommendedModelWarning: boolean;
  imageReasoningEffort: ImageReasoningEffort;
  experimentalRewindSnapshots: boolean;
  nextPromptSuggestionsEnabled: boolean;
  autonomyLevel: AutonomyLevel;
  toolPermissionOverrides: ToolPermissionOverride[];
}

export const DEFAULT_USER_PREFERENCES: UserPreferences = {
  showThinkingTraces: false,
  autoAttachVisuals: true,
  showTokenUsage: true,
  compactMessages: false,
  defaultThinkingLevel: "high",
  imageGenerationEnabled: false,
  defaultImageModel: "",
  defaultModelByProvider: {},
  suppressUnrecommendedModelWarning: false,
  imageReasoningEffort: "high",
  experimentalRewindSnapshots: false,
  nextPromptSuggestionsEnabled: true,
  autonomyLevel: "medium",
  toolPermissionOverrides: [],
};

export const PROVIDER_AUTH_STATES = ["not_configured", "credential_stored", "verified_usable", "verification_failed"] as const;
export type ProviderAuthState = (typeof PROVIDER_AUTH_STATES)[number];

export interface ProviderAuthDescriptor {
  provider: string;
  state: ProviderAuthState;
  credentialStored: boolean;
  verifiedUsable: boolean;
  verifiedAt?: string | undefined;
  lastVerificationAttemptAt?: string | undefined;
  lastVerificationError?: string | undefined;
}

export interface ProviderDescriptor {
  provider: string;
  label: string;
  settingsVisibility: SettingsVisibility;
  lab: string;
  defaultModelId?: string | undefined;
  supportStatus: ProviderSupportStatus;
  runtimeSurface: ProviderRuntimeSurface;
  authMethods: ProviderAuthMethod[];
  apiKeySupported: boolean;
  browserCallable: boolean;
  companionRequired: boolean;
  subscriptionBacked: boolean;
  imageGenerationSupported: boolean;
  capabilityNote?: string | undefined;
  authState: ProviderAuthState;
  credentialStored: boolean;
  verifiedUsable: boolean;
  verificationError?: string | undefined;
  verifiedAt?: string | undefined;
  configured: boolean;
  oauthSupported: boolean;
  models: ProviderModelDescriptor[];
}

export interface ProviderCatalogResponse {
  providers: ProviderDescriptor[];
}

export type ProviderSupportStatus = "supported" | "planned" | "blocked" | "research_only";
export type ProviderRuntimeSurface = "browser_taskpane" | "companion" | "not_implemented";
export type SettingsVisibility = "simple" | "advanced";
export type ProviderAuthMethod =
  | "api_key"
  | "oauth"
  | "manual_token"
  | "cloud_identity"
  | "aws_credentials";

export interface AuthStatusResponse {
  storedProviders: string[];
  oauthProviders: string[];
  configuredProviders: string[];
  verifiedProviders: string[];
  unverifiedProviders: string[];
  verificationFailedProviders: string[];
  providerStates: ProviderAuthDescriptor[];
}

export interface SessionUsageTotals {
  input: number;
  output: number;
  cacheRead: number;
  cacheWrite: number;
  total: number;
}

export interface ContextBreakdownEntry {
  label: string;
  tokens: number;
  color: string;
}

export interface SessionContextUsage {
  tokens: number | null;
  contextWindow: number;
  percent: number | null;
  breakdown?: ContextBreakdownEntry[] | undefined;
}

export interface SessionStatsResponse {
  sessionFile?: string | undefined;
  sessionId: string;
  userMessages: number;
  assistantMessages: number;
  toolCalls: number;
  toolResults: number;
  totalMessages: number;
  tokens: SessionUsageTotals;
  cost: number;
  contextUsage?: SessionContextUsage | undefined;
}

export interface CompanionHealthResponse {
  ok: true;
  endpoint: string;
  identity: string;
  capabilities: CompanionCapabilities;
}

export interface OfficeContextPayload {
  summary: string;
  state: OfficeStateUpdate;
  anchors?: OfficeAnchor[] | undefined;
  formatting?: Record<string, unknown> | undefined;
  snippets?: Record<string, unknown> | undefined;
  visuals?: OfficeVisualSnapshot[] | undefined;
}

export const OFFICE_PROPOSE_EDITS_SEARCH_TEXT_MAX_LENGTH = 200;

export interface OfficeProposedEdit {
  id: string;
  kind: "insert" | "replace" | "delete";
  anchor?: string | undefined;
  paragraphId?: string | undefined;
  searchText?: string | undefined;
  oldText?: string | undefined;
  newText?: string | undefined;
  explanation?: string | undefined;
  found?: boolean | undefined;
  contextPreview?: string | undefined;
}

export interface OfficeEditProposal {
  requestId: string;
  edits: OfficeProposedEdit[];
  summary: string;
}

export const EDIT_REJECT_REASONS = ["keep_original", "rewrite_differently", "not_relevant"] as const;
export type EditRejectReason = (typeof EDIT_REJECT_REASONS)[number];

export const EDIT_REJECT_REASON_LABELS: Record<EditRejectReason, string> = {
  keep_original: "Keep original",
  rewrite_differently: "Rewrite differently",
  not_relevant: "Not relevant",
};

export interface OfficeEditProposalDecision {
  requestId: string;
  decisions: Array<{
    editId: string;
    accepted: boolean;
    modifiedText?: string | undefined;
    rejectReason?: EditRejectReason | undefined;
    rejectNote?: string | undefined;
  }>;
  globalFeedback?: string | undefined;
  applicationResult?: {
    applied: number;
    failed: number;
    errors: string[];
  } | undefined;
}

export const OFFICE_TOOL_NAMES = [
  "office_get_context",
  "office_tool_search",
  "office_tool_get",
  "office_tool_call",
  "office_batch_execute",
  "mcp_tool_search",
  "mcp_batch_execute",
  "mcp_result_get",
  "mcp_result_summarize",
  "mcp_result_clear",
  "office_apply_edit",
  "edit_doc_text",
  "edit_doc_list",
  "get_cell_ranges",
  "set_cell_range",
  "clear_cell_range",
  "resize_range",
  "copy_to",
  "modify_sheet_structure",
  "modify_object",
  "get_all_objects",
  "search_data",
  "get_range_as_csv",
  "read_range_image",
  "extract_chart_xml",
  "office_navigate",
  "office_capture_snapshot",
  "office_capture_viewport",
  "office_read_section",
  "word_search",
  "word_format_text",
  "word_list_format",
  "word_reference_inventory",
  "word_hyperlink",
  "word_table",
  "word_section_layout",
  "word_equation",
  "word_field_reference",
  "word_content_control",
  "word_building_block",
  "word_annotation_review",
  "word_redline_review",
  "word_proofing_stats",
  "word_collab_guard",
  "verify_doc",
  "verify_doc_visual",
  "get_presentation_structure",
  "get_slide",
  "list_slide_shapes",
  "modify_presentation_structure",
  "duplicate_slide",
  "insert_slide_element",
  "remove_slide_element",
  "edit_slide_text",
  "edit_slide_xml",
  "edit_slide_master",
  "edit_slide_chart",
  "copy_image_between_slides",
  "search_icons",
  "insert_icon",
  "verify_slides",
  "verify_slide_visual",
  "office_execute_js",
  "office_propose_edits",
] as const;
export type OfficeToolName = (typeof OFFICE_TOOL_NAMES)[number];

export interface OfficeToolRequest {
  requestId: string;
  toolName: OfficeToolName;
  host: OfficeHost;
  params: Record<string, unknown>;
}

export interface OfficeToolResult {
  requestId: string;
  success: boolean;
  content?: unknown;
  error?: string;
}

export interface AskUserOption {
  title: string;
  description?: string | undefined;
}

export interface AskUserQuestion {
  id: string;
  question: string;
  context?: string | undefined;
  options: AskUserOption[];
}

export interface AskUserRequest {
  requestId: string;
  questions: AskUserQuestion[];
}

export interface AskUserQuestionAnswer {
  questionId: string;
  selectedOption: string | null;
  notes?: string | undefined;
}

export interface AskUserResponse {
  requestId: string;
  answers: AskUserQuestionAnswer[];
}

export interface CheckpointMetadata {
  id: string;
  timestamp: number;
  userPrompt: string;
  file: string;
}

export type BridgeServerMessage =
  | { type: "session_event"; event: unknown }
  | { type: "office_tool_call"; request: OfficeToolRequest }
  | { type: "ask_user_request"; request: AskUserRequest }
  | { type: "edit_proposal_request"; proposal: OfficeEditProposal }
  | { type: "tool_permission_request"; request: ToolPermissionRequest }
  | { type: "tool_permission_expired"; requestId: string; toolName: string }
  | { type: "connection_state"; state: "ready" | "reconnecting" | "offline" }
  | { type: "available_checkpoints"; checkpoints: CheckpointMetadata[] }
  | { type: "error"; message: string };

export type BridgeClientMessage =
  | { type: "client_ready" }
  | { type: "ping" }
  | { type: "office_tool_result"; result: OfficeToolResult }
  | { type: "ask_user_response"; response: AskUserResponse }
  | { type: "edit_proposal_decision"; decision: OfficeEditProposalDecision }
  | { type: "tool_permission_response"; requestId: string; decision: ToolPermissionDecision }
  | { type: "rewind_session"; targetMessageCount: number }
  | { type: "persist_checkpoint"; documentId: string; checkpoint: DocumentCheckpointPayload }
  | { type: "load_checkpoint"; documentId: string; checkpointId: string };

export interface DocumentCheckpointPayload {
  id: string;
  timestamp: number;
  host: "word" | "excel" | "powerpoint";
  userPrompt: string;
  messageCount: number;
  ooxml?: string;
  sheets?: Array<{
    name: string;
    usedRangeAddress: string;
    values: unknown[][];
    numberFormats: string[][];
    formulas: unknown[][];
  }>;
  presentationBase64?: string;
}
