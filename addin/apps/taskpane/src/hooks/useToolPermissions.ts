import { useCallback, useRef } from "react";
import {
  AUTONOMY_LEVEL_AUTO_APPROVE,
  TOOL_CATEGORY_MAP,
  type AutonomyLevel,
  type ToolCategory,
  type ToolPermissionDecision,
  type ToolPermissionOverride,
} from "@pi-office/pi-office-pack/protocol";

const SESSION_APPROVALS_KEY = "pi-office-session-approvals";
const WORKSPACE_APPROVALS_PREFIX = "pi-office-workspace-approvals:";
const ALWAYS_APPROVALS_KEY = "pi-office-always-approvals";

function loadSet(storage: Storage, key: string): Set<string> {
  try {
    const raw = storage.getItem(key);
    return raw ? new Set(JSON.parse(raw)) : new Set();
  } catch {
    return new Set();
  }
}

function saveSet(storage: Storage, key: string, set: Set<string>): void {
  try {
    storage.setItem(key, JSON.stringify([...set]));
  } catch {}
}

function getToolCategory(toolName: string): ToolCategory {
  return TOOL_CATEGORY_MAP[toolName] ?? "connector";
}

function getEffectiveLevel(
  toolName: string,
  category: ToolCategory,
  autonomyLevel: AutonomyLevel,
  overrides: ToolPermissionOverride[],
): AutonomyLevel | "disabled" {
  const override = overrides.find((o) => o.toolName === toolName);
  if (override) return override.autoApproveAtLevel;
  return autonomyLevel;
}

export function useToolPermissions(
  autonomyLevel: AutonomyLevel,
  overrides: ToolPermissionOverride[],
  workspaceDir: string | undefined,
) {
  const sessionApprovalsRef = useRef(loadSet(sessionStorage, SESSION_APPROVALS_KEY));
  const workspaceApprovalsRef = useRef(
    workspaceDir ? loadSet(localStorage, `${WORKSPACE_APPROVALS_PREFIX}${workspaceDir}`) : new Set<string>(),
  );
  const alwaysApprovalsRef = useRef(loadSet(localStorage, ALWAYS_APPROVALS_KEY));

  const shouldAutoApprove = useCallback(
    (toolName: string): boolean => {
      const category = getToolCategory(toolName);
      if (category === "interaction") return true;

      if (alwaysApprovalsRef.current.has(toolName)) return true;
      if (workspaceDir && workspaceApprovalsRef.current.has(toolName)) return true;
      if (sessionApprovalsRef.current.has(toolName)) return true;

      const effectiveLevel = getEffectiveLevel(toolName, category, autonomyLevel, overrides);
      if (effectiveLevel === "disabled") return false;
      const approvedCategories = AUTONOMY_LEVEL_AUTO_APPROVE[effectiveLevel];
      return approvedCategories.has(category);
    },
    [autonomyLevel, overrides, workspaceDir],
  );

  const grantPermission = useCallback(
    (toolName: string, scope: ToolPermissionDecision["scope"]): void => {
      switch (scope) {
        case "session": {
          sessionApprovalsRef.current.add(toolName);
          saveSet(sessionStorage, SESSION_APPROVALS_KEY, sessionApprovalsRef.current);
          break;
        }
        case "workspace": {
          if (workspaceDir) {
            workspaceApprovalsRef.current.add(toolName);
            saveSet(localStorage, `${WORKSPACE_APPROVALS_PREFIX}${workspaceDir}`, workspaceApprovalsRef.current);
          } else {
            sessionApprovalsRef.current.add(toolName);
            saveSet(sessionStorage, SESSION_APPROVALS_KEY, sessionApprovalsRef.current);
          }
          break;
        }
        case "always": {
          alwaysApprovalsRef.current.add(toolName);
          saveSet(localStorage, ALWAYS_APPROVALS_KEY, alwaysApprovalsRef.current);
          break;
        }
      }
    },
    [workspaceDir],
  );

  const revokePermission = useCallback(
    (toolName: string): void => {
      sessionApprovalsRef.current.delete(toolName);
      saveSet(sessionStorage, SESSION_APPROVALS_KEY, sessionApprovalsRef.current);
      if (workspaceDir) {
        workspaceApprovalsRef.current.delete(toolName);
        saveSet(localStorage, `${WORKSPACE_APPROVALS_PREFIX}${workspaceDir}`, workspaceApprovalsRef.current);
      }
      alwaysApprovalsRef.current.delete(toolName);
      saveSet(localStorage, ALWAYS_APPROVALS_KEY, alwaysApprovalsRef.current);
    },
    [workspaceDir],
  );

  return { shouldAutoApprove, grantPermission, revokePermission, getToolCategory } as const;
}

export { getToolCategory };
