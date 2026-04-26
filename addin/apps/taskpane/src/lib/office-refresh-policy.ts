export const OFFICE_REFRESH_ERROR_DEDUP_WINDOW_MS = 15_000;

export interface OfficeRefreshAttemptToken {
  sequence: number;
  sessionId: string | undefined;
}

export interface OfficeRefreshCurrentState {
  active: boolean;
  latestSequence: number;
  sessionId: string | undefined;
}

export interface OfficeRefreshErrorRecord {
  message: string;
  firstSeenAt: number;
  lastShownAt: number;
  suppressedCount: number;
}

export interface OfficeRefreshErrorDecision {
  shouldSurface: boolean;
  nextRecord: OfficeRefreshErrorRecord;
}

export function shouldApplyOfficeRefreshResult(
  attempt: OfficeRefreshAttemptToken,
  current: OfficeRefreshCurrentState,
): boolean {
  return current.active && attempt.sequence === current.latestSequence && attempt.sessionId === current.sessionId;
}

export function formatOfficeRefreshError(error: unknown): string {
  if (error instanceof Error && error.message.trim()) {
    return error.message.trim();
  }
  if (typeof error === "string" && error.trim()) {
    return error.trim();
  }
  try {
    return JSON.stringify(error);
  } catch {
    return "Unknown Office host error.";
  }
}

export function getOfficeRefreshErrorDecision(
  previous: OfficeRefreshErrorRecord | undefined,
  message: string,
  now = Date.now(),
  dedupWindowMs = OFFICE_REFRESH_ERROR_DEDUP_WINDOW_MS,
): OfficeRefreshErrorDecision {
  if (!previous || previous.message !== message) {
    return {
      shouldSurface: true,
      nextRecord: {
        message,
        firstSeenAt: now,
        lastShownAt: now,
        suppressedCount: 0,
      },
    };
  }

  if (now - previous.lastShownAt >= dedupWindowMs) {
    return {
      shouldSurface: true,
      nextRecord: {
        message,
        firstSeenAt: previous.firstSeenAt,
        lastShownAt: now,
        suppressedCount: 0,
      },
    };
  }

  return {
    shouldSurface: false,
    nextRecord: {
      ...previous,
      suppressedCount: previous.suppressedCount + 1,
    },
  };
}
