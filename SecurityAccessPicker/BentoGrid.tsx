import * as React from "react";
import "./BentoGrid.css";

import { IInputs } from "./generated/ManifestTypes";
import { Input } from "@fluentui/react-input";
import { Button } from "@fluentui/react-button";

export interface BentoGridProps {
  context: ComponentFramework.Context<IInputs>;
}

interface OneColItem {
  id: string;
  roleName: string;
  businessUnitName: string;
}
type MatrixColor = "#107C10" | "#2266E3" | "#A52A2A" | "#797775" | "#FFFFFF";
interface MatrixItem {
  id: string;
  entity: string;
  c1: MatrixColor;
  c2: MatrixColor;
  c3: MatrixColor;
  c4: MatrixColor;
  c5: MatrixColor;
  c6: MatrixColor;
  c7: MatrixColor;
  c8: MatrixColor;
}
interface RoleRecord {
  roleid: string;
  name: string;
  businessUnitName: string;
}
type DepthClass = "basic" | "local" | "deep" | "global";
type SecurityPrincipalType = "team" | "systemuser";
interface SecurityPrincipalConfig {
  roleRelationship: string;
  privilegeRequestName: string;
  metadataTypeName: string;
  displayNameColumn: string;
  defaultLabel: string;
}

interface SecurityPrincipalResolution {
  id: string;
  entityType: SecurityPrincipalType;
}

interface ContextInfoLike {
  entityId?: unknown;
  entityTypeName?: unknown;
  entityLogicalName?: unknown;
  entityName?: unknown;
}

interface ModeWithContextInfo extends ComponentFramework.Mode {
  contextInfo?: ContextInfoLike;
}

interface XrmExecuteResponse {
  ok: boolean;
  json: () => Promise<unknown>;
}

interface XrmWebApiLike {
  execute: (request: unknown) => Promise<XrmExecuteResponse>;
}

interface XrmEntityLike {
  getId?: () => string;
  getEntityName?: () => string;
}

interface XrmPageLike {
  data?: {
    entity?: XrmEntityLike;
  };
}

interface XrmLike {
  WebApi?: XrmWebApiLike;
  Page?: XrmPageLike;
  App?: XrmAppLike;
}

interface WindowWithXrm extends Window {
  Xrm?: XrmLike;
}

interface XrmGlobalNotification {
  type: number;
  level: number;
  message: string;
  showCloseButton: boolean;
}

interface XrmAppLike {
  addGlobalNotification: (notification: XrmGlobalNotification) => Promise<string>;
}

interface RawPrivilegeRecord {
  Depth?: unknown;
  depth?: unknown;
  PrivilegeDepth?: unknown;
  privilegeDepth?: unknown;
  PrivilegeName?: unknown;
  privilegeName?: unknown;
  name?: unknown;
}

interface NormalizedPrivilege {
  depth: unknown;
  privilegeName: string;
}

const PARSE_VERBS = ["Create", "Read", "Write", "Delete", "Assign", "AppendTo", "Append", "Share"];
const RANK: Record<DepthClass, number> = { basic: 0, local: 1, deep: 2, global: 3 };
const SECURITY_PRINCIPAL_CONFIG: Record<SecurityPrincipalType, SecurityPrincipalConfig> = {
  team: {
    roleRelationship: "teamroles_association",
    privilegeRequestName: "RetrieveTeamPrivileges",
    metadataTypeName: "mscrm.team",
    displayNameColumn: "name",
    defaultLabel: "Team",
  },
  systemuser: {
    roleRelationship: "systemuserroles_association",
    privilegeRequestName: "RetrieveUserPrivileges",
    metadataTypeName: "mscrm.systemuser",
    displayNameColumn: "fullname",
    defaultLabel: "User",
  },
};

const cleanGuid = (id: string): string => (id ?? "").replace(/[{}]/g, "");

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

const getStringProperty = (value: unknown, key: string): string | null => {
  if (!isRecord(value)) return null;
  const candidate = value[key];
  return typeof candidate === "string" ? candidate : null;
};

const getArrayProperty = (value: unknown, key: string): unknown[] => {
  if (!isRecord(value)) return [];
  const candidate = value[key];
  return Array.isArray(candidate) ? candidate : [];
};

const toRoleRecord = (value: unknown): RoleRecord | null => {
  if (!isRecord(value)) return null;
  const roleid = typeof value.roleid === "string" ? value.roleid : null;
  const name = typeof value.name === "string" ? value.name : null;
  if (!roleid || !name) return null;
  const buFromFormattedValue = getStringProperty(value, "_businessunitid_value@OData.Community.Display.V1.FormattedValue");
  const buFromExpandedRecord = getStringProperty(value.businessunitid, "name");
  const businessUnitName = buFromFormattedValue ?? buFromExpandedRecord ?? "";
  return { roleid, name, businessUnitName };
};

const getPageEntityId = (windowRef: WindowWithXrm | null | undefined): string | null => {
  const entity = windowRef?.Xrm?.Page?.data?.entity;
  if (!entity?.getId) return null;
  try {
    const id = entity.getId();
    return typeof id === "string" ? id : null;
  } catch {
    return null;
  }
};

const getPageEntityName = (windowRef: WindowWithXrm | null | undefined): string | null => {
  const entity = windowRef?.Xrm?.Page?.data?.entity;
  if (!entity?.getEntityName) return null;
  try {
    const entityName = entity.getEntityName();
    return typeof entityName === "string" ? entityName : null;
  } catch {
    return null;
  }
};

const getXrmWebApi = (): XrmWebApiLike | null => {
  const w = window as WindowWithXrm;
  const parentWindow = window.parent as WindowWithXrm | null;
  return w.Xrm?.WebApi ?? parentWindow?.Xrm?.WebApi ?? null;
};

const getXrmApp = (): XrmAppLike | null => {
  const w = window as WindowWithXrm;
  const parentWindow = window.parent as WindowWithXrm | null;
  return w.Xrm?.App ?? parentWindow?.Xrm?.App ?? null;
};

const normalizeSecurityPrincipalType = (value: unknown): SecurityPrincipalType | null => {
  if (typeof value !== "string") return null;
  const normalized = value.trim().toLowerCase();
  if (normalized === "team") return "team";
  if (normalized === "systemuser") return "systemuser";
  return null;
};

const getCurrentSecurityPrincipal = (
  context: ComponentFramework.Context<IInputs>
): SecurityPrincipalResolution | null => {
  const mode = context.mode as ModeWithContextInfo;
  const contextInfo = mode.contextInfo;
  const fromContextInfoId = typeof contextInfo?.entityId === "string" ? contextInfo.entityId : null;
  const fromContextInfoType = normalizeSecurityPrincipalType(
    contextInfo?.entityTypeName ?? contextInfo?.entityLogicalName ?? contextInfo?.entityName
  );

  const w = window as WindowWithXrm;
  const parentWindow = window.parent as WindowWithXrm | null;
  const fromPageId = getPageEntityId(w) ?? getPageEntityId(parentWindow);
  const fromPageType = normalizeSecurityPrincipalType(getPageEntityName(w) ?? getPageEntityName(parentWindow));

  const entityId =
    typeof fromContextInfoId === "string" && fromContextInfoId
      ? fromContextInfoId
      : typeof fromPageId === "string" && fromPageId
      ? fromPageId
      : null;

  if (!entityId) return null;

  return {
    id: cleanGuid(entityId),
    entityType: fromContextInfoType ?? fromPageType ?? "team",
  };
};

const depthToClass = (depth: unknown): DepthClass | null => {
  if (depth == null) return null;
  if (typeof depth === "string") {
    const s = depth.trim().toLowerCase();
    if (s === "basic") return "basic";
    if (s === "local") return "local";
    if (s === "deep") return "deep";
    if (s === "global") return "global";
    return null;
  }

  const d = Number(depth);
  if (!Number.isFinite(d)) return null;
  if (d >= 0 && d <= 3) return ["basic", "local", "deep", "global"][d] as DepthClass;
  if (d & 8) return "global";
  if (d & 4) return "deep";
  if (d & 2) return "local";
  if (d & 1) return "basic";
  return null;
};

const depthToColor = (depthClass: DepthClass | null | undefined): MatrixColor => {
  if (depthClass === "basic") return "#797775";
  if (depthClass === "local") return "#2266E3";
  if (depthClass === "deep") return "#A52A2A";
  if (depthClass === "global") return "#107C10";
  return "#FFFFFF";
};

const parsePrivilegeName = (name: string | undefined | null): { right: string; entityKey: string } | null => {
  if (!name || typeof name !== "string") return null;
  for (const v of PARSE_VERBS) {
    const prefix = `prv${v}`;
    if (name.startsWith(prefix)) {
      return { right: v, entityKey: name.substring(prefix.length) };
    }
  }
  return null;
};

const prettifyEntity = (entityKey: string): string =>
  (entityKey ?? "").replace(/^new_/i, "").replace(/([a-z])([A-Z])/g, "$1 $2");

const toNormalizedPrivileges = (result: unknown): NormalizedPrivilege[] => {
  const rolePrivilegesRaw = getArrayProperty(result, "RolePrivileges");
  const rolePrivilegesFallback = getArrayProperty(result, "rolePrivileges");
  const source = rolePrivilegesRaw.length > 0 ? rolePrivilegesRaw : rolePrivilegesFallback;

  return source
    .filter((entry): entry is RawPrivilegeRecord => isRecord(entry))
    .map((entry) => {
      const depth = entry.Depth ?? entry.depth ?? entry.PrivilegeDepth ?? entry.privilegeDepth;
      const privilegeNameCandidate = entry.PrivilegeName ?? entry.privilegeName ?? entry.name;
      const privilegeName = typeof privilegeNameCandidate === "string" ? privilegeNameCandidate : null;
      return privilegeName ? { depth, privilegeName } : null;
    })
    .filter((entry): entry is NormalizedPrivilege => entry !== null);
};

const errorMessageFromUnknown = (error: unknown): string => {
  if (error instanceof Error && error.message) return error.message;
  const directMessage = getStringProperty(error, "message");
  if (directMessage) return directMessage;
  const rawValue = getStringProperty(error, "raw");
  if (rawValue) {
    try {
      const parsed = JSON.parse(rawValue) as unknown;
      const parsedMessage = getStringProperty(parsed, "message");
      if (parsedMessage) return parsedMessage;
    } catch {
      return rawValue;
    }
  }
  return "An unexpected error occurred.";
};

export const BentoGrid: React.FC<BentoGridProps> = ({ context }) => {
  const [query, setQuery] = React.useState("");
  const [principalId, setPrincipalId] = React.useState<string | null>(null);
  const [principalType, setPrincipalType] = React.useState<SecurityPrincipalType>("team");
  const [allRoles, setAllRoles] = React.useState<RoleRecord[]>([]);
  const [assignedRoles, setAssignedRoles] = React.useState<RoleRecord[]>([]);
  const [principalName, setPrincipalName] = React.useState<string>("Team");
  const [selectedLeftId, setSelectedLeftId] = React.useState<string | null>(null);
  const [selectedRightId, setSelectedRightId] = React.useState<string | null>(null);
  const [matrixItems, setMatrixItems] = React.useState<MatrixItem[]>([]);
  const notifyError = React.useCallback(async (error: unknown) => {
    const message = errorMessageFromUnknown(error);
    const xrmApp = getXrmApp();
    if (xrmApp?.addGlobalNotification) {
      await xrmApp.addGlobalNotification({
        type: 2,
        level: 2,
        message,
        showCloseButton: true,
      });
      return;
    }
    console.error(message);
  }, []);

  const loadGrid = React.useCallback(async (currentPrincipalType: SecurityPrincipalType, currentPrincipalId: string) => {
    const xrmWebApi = getXrmWebApi();
    if (!xrmWebApi?.execute) {
      setMatrixItems([]);
      return;
    }
    const principalConfig = SECURITY_PRINCIPAL_CONFIG[currentPrincipalType];
    console.log("Loading grid with principal:", currentPrincipalType, currentPrincipalId);

    const req = {
      entity: { entityType: currentPrincipalType, id: cleanGuid(currentPrincipalId) },
      getMetadata: () => ({
        boundParameter: "entity",
        parameterTypes: { entity: { typeName: principalConfig.metadataTypeName, structuralProperty: 5 } },
        operationType: 1,
        operationName: principalConfig.privilegeRequestName,
      }),
    };

    const response = await xrmWebApi.execute(req);
    if (!response?.ok) {
      throw new Error(`${principalConfig.privilegeRequestName} failed.`);
    }

    const result = await response.json();
    const rolePrivileges = toNormalizedPrivileges(result);

    const matrix = new Map<string, Map<string, DepthClass>>();
    for (const rp of rolePrivileges) {
      const parsed = parsePrivilegeName(rp.privilegeName);
      if (!parsed) continue;

      const cls = depthToClass(rp.depth);
      if (!cls) continue;

      if (!matrix.has(parsed.entityKey)) {
        matrix.set(parsed.entityKey, new Map<string, DepthClass>());
      }

      const row = matrix.get(parsed.entityKey)!;
      const current = row.get(parsed.right);
      if (!current || RANK[cls] > RANK[current]) {
        row.set(parsed.right, cls);
      }
    }

    const rows = Array.from(matrix.entries())
      .map(([entityKey, rights]) => ({ entityKey, label: prettifyEntity(entityKey), rights }))
      .sort((a, b) => (a.label ?? "").localeCompare(b.label ?? ""));

    setMatrixItems(
      rows.map((row) => ({
        id: row.entityKey || row.label,
        entity: row.label,
        c1: depthToColor(row.rights.get("Create")),
        c2: depthToColor(row.rights.get("Read")),
        c3: depthToColor(row.rights.get("Write")),
        c4: depthToColor(row.rights.get("Delete")),
        c5: depthToColor(row.rights.get("Append")),
        c6: depthToColor(row.rights.get("AppendTo")),
        c7: depthToColor(row.rights.get("Assign")),
        c8: depthToColor(row.rights.get("Share")),
      }))
    );
  }, []);

  const loadInitialData = React.useCallback(
    async (currentPrincipalType: SecurityPrincipalType, currentPrincipalId: string) => {
      const principalConfig = SECURITY_PRINCIPAL_CONFIG[currentPrincipalType];
      const [principal, roles] = await Promise.all([
        context.webAPI.retrieveRecord(
          currentPrincipalType,
          currentPrincipalId,
          `?$select=${principalConfig.displayNameColumn}&$expand=${principalConfig.roleRelationship}($select=roleid,name,_businessunitid_value;$orderby=name asc)`
        ),
        context.webAPI.retrieveMultipleRecords("role", "?$select=roleid,name,_businessunitid_value&$orderby=name asc"),
      ]);

      const teamRoles = getArrayProperty(principal, principalConfig.roleRelationship)
        .map(toRoleRecord)
        .filter((record): record is RoleRecord => record !== null);
      const principalLabel = getStringProperty(principal, principalConfig.displayNameColumn) ?? principalConfig.defaultLabel;
      const roleEntities = isRecord(roles) && Array.isArray(roles.entities) ? roles.entities : [];
      const allRoleRecords = roleEntities.map(toRoleRecord).filter((record): record is RoleRecord => record !== null);

      setPrincipalName(principalLabel);
      setAssignedRoles(teamRoles.map((r) => ({ roleid: r.roleid, name: r.name, businessUnitName: r.businessUnitName })));
      setAllRoles(allRoleRecords.map((r) => ({ roleid: r.roleid, name: r.name, businessUnitName: r.businessUnitName })));
      await loadGrid(currentPrincipalType, currentPrincipalId);
      setSelectedLeftId(null);
      setSelectedRightId(null);
    },
    [context, loadGrid]
  );

  React.useEffect(() => {
    const principal = getCurrentSecurityPrincipal(context);
    setPrincipalId(principal?.id ?? null);
    if (!principal) return;
    setPrincipalType(principal.entityType);

    loadInitialData(principal.entityType, principal.id).catch((e: unknown) => {
      void notifyError(e);
    });
  }, [context, loadInitialData, notifyError]);

  const assignedSet = React.useMemo(() => new Set(assignedRoles.map((x) => x.roleid)), [assignedRoles]);
  const normalizedQuery = React.useMemo(() => query.trim().toLowerCase(), [query]);

  const leftItems = React.useMemo<OneColItem[]>(
    () =>
      allRoles
        .filter((r) => !assignedSet.has(r.roleid))
        .filter((r) => (r.name ?? "").toLowerCase().includes(normalizedQuery))
        .map((r) => ({ id: r.roleid, roleName: r.name ?? "(no name)", businessUnitName: r.businessUnitName ?? "" })),
    [allRoles, assignedSet, normalizedQuery]
  );

  const rightItems = React.useMemo<OneColItem[]>(
    () =>
      assignedRoles
        .filter((r) => (r.name ?? "").toLowerCase().includes(normalizedQuery))
        .map((r) => ({ id: r.roleid, roleName: r.name ?? "(no name)", businessUnitName: r.businessUnitName ?? "" })),
    [assignedRoles, normalizedQuery]
  );

  const selectedAvailable = React.useMemo(
    () => leftItems.find((x) => x.id === selectedLeftId),
    [leftItems, selectedLeftId]
  );
  const selectedAssigned = React.useMemo(
    () => rightItems.find((x) => x.id === selectedRightId),
    [rightItems, selectedRightId]
  );

  const handleAdd = React.useCallback(async () => {
    if (!principalId || !selectedAvailable) return;

    try {
      const xrmWebApi = getXrmWebApi();
      if (!xrmWebApi?.execute) return;
      const principalConfig = SECURITY_PRINCIPAL_CONFIG[principalType];

      const associateRequest = {
        target: { entityType: principalType, id: cleanGuid(principalId) },
        relatedEntities: [{ entityType: "role", id: cleanGuid(String(selectedAvailable.id)) }],
        relationship: principalConfig.roleRelationship,
        getMetadata: () => ({
          boundParameter: null,
          parameterTypes: {},
          operationType: 2,
          operationName: "Associate",
        }),
      };

      const response = await xrmWebApi.execute(associateRequest);
      if (!response?.ok) {
        throw new Error("Associate failed.");
      }

      setAssignedRoles((prev) => {
        if (prev.some((r) => r.roleid === selectedAvailable.id)) return prev;
        const next = [
          ...prev,
          { roleid: selectedAvailable.id, name: selectedAvailable.roleName, businessUnitName: selectedAvailable.businessUnitName },
        ];
        next.sort((a, b) => (a.name ?? "").localeCompare(b.name ?? ""));
        return next;
      });
      setSelectedLeftId(null);
      setSelectedRightId(selectedAvailable.id);
      await loadGrid(principalType, principalId);
    } catch (error: unknown) {
      await notifyError(error);
    }
  }, [loadGrid, notifyError, principalId, principalType, selectedAvailable]);

  const handleRemove = React.useCallback(async () => {
    if (!principalId || !selectedAssigned) return;

    try {
      const xrmWebApi = getXrmWebApi();
      if (!xrmWebApi?.execute) return;
      const principalConfig = SECURITY_PRINCIPAL_CONFIG[principalType];

      const disassociateRequest = {
        target: { entityType: principalType, id: cleanGuid(principalId) },
        relatedEntityId: cleanGuid(String(selectedAssigned.id)),
        relationship: principalConfig.roleRelationship,
        getMetadata: () => ({
          boundParameter: null,
          parameterTypes: {
            target: { typeName: principalConfig.metadataTypeName, structuralProperty: 5 },
            relatedEntityId: { typeName: "Edm.Guid", structuralProperty: 1 },
            relationship: { typeName: "Edm.String", structuralProperty: 1 },
          },
          operationType: 2,
          operationName: "Disassociate",
        }),
      };

      const response = await xrmWebApi.execute(disassociateRequest);
      if (!response?.ok) {
        throw new Error("Disassociate failed.");
      }

      setAssignedRoles((prev) => prev.filter((r) => r.roleid !== selectedAssigned.id));
      setSelectedRightId(null);
      setSelectedLeftId(selectedAssigned.id);
      await loadGrid(principalType, principalId);
    } catch (error: unknown) {
      await notifyError(error);
    }
  }, [loadGrid, notifyError, principalId, principalType, selectedAssigned]);

  return (
    <main className="page">
      <section className="bento">
        <article className="card span-2">
          <div className="inner-grid">
            <div className="block c1 topbar">
              <Input
                placeholder="Search..."
                value={query}
                onChange={(_, data) => setQuery(data.value)}
                aria-label={`Search roles for ${principalName}`}
              />
            </div>

            <div className="block c2 leftcol fluent-grid">
              <div className="role-table-header">
                <span className="grid-header-text">Available Roles</span>
              </div>
              <table className="role-table">
                <tbody>
                  {leftItems.map((item) => {
                    const isSelected = selectedLeftId === item.id;
                    return (
                      <tr
                        key={item.id}
                        className={`role-table-row${isSelected ? " is-selected" : ""}`}
                        onClick={() => {
                          setSelectedRightId(null);
                          setSelectedLeftId(item.id);
                        }}
                      >
                        <td>
                          <div className="role-row-content">
                            <span className="grid-body-text">{item.roleName}</span>
                            {item.businessUnitName ? (
                              <span className="grid-body-text role-business-unit">{item.businessUnitName}</span>
                            ) : null}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div className="block c3 midcol section3">
              <Button
                className="stack-btn"
                type="button"
                onClick={() => {
                  void handleAdd();
                }}
                disabled={!principalId || !selectedAvailable}
              >
                &gt;
              </Button>
              <Button
                className="stack-btn"
                type="button"
                onClick={() => {
                  void handleRemove();
                }}
                disabled={!principalId || !selectedAssigned}
              >
                &lt;
              </Button>
            </div>

            <div className="block c4 rightcol fluent-grid">
              <div className="role-table-header">
                <span className="grid-header-text">Assigned Roles</span>
              </div>
              <table className="role-table">
                <tbody>
                  {rightItems.map((item) => {
                    const isSelected = selectedRightId === item.id;
                    return (
                      <tr
                        key={item.id}
                        className={`role-table-row${isSelected ? " is-selected" : ""}`}
                        onClick={() => {
                          setSelectedLeftId(null);
                          setSelectedRightId(item.id);
                        }}
                      >
                        <td>
                          <div className="role-row-content">
                            <span className="grid-body-text">{item.roleName}</span>
                            {item.businessUnitName ? (
                              <span className="grid-body-text role-business-unit">{item.businessUnitName}</span>
                            ) : null}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </article>

        <article className="card span-3">
          <div className="block c5 fill matrix-grid">
            <table className="matrix-table">
              <thead>
                <tr>
                  <th />
                  <th>Create</th>
                  <th>Read</th>
                  <th>Write</th>
                  <th>Delete</th>
                  <th>Append</th>
                  <th>Append To</th>
                  <th>Assign</th>
                  <th>Share</th>
                </tr>
              </thead>
              <tbody>
                {matrixItems.map((item) => (
                  <tr key={item.id}>
                    <td className="entity-cell">{item.entity}</td>
                    <td><span className="matrix-dot" style={{ backgroundColor: item.c1 }} /></td>
                    <td><span className="matrix-dot" style={{ backgroundColor: item.c2 }} /></td>
                    <td><span className="matrix-dot" style={{ backgroundColor: item.c3 }} /></td>
                    <td><span className="matrix-dot" style={{ backgroundColor: item.c4 }} /></td>
                    <td><span className="matrix-dot" style={{ backgroundColor: item.c5 }} /></td>
                    <td><span className="matrix-dot" style={{ backgroundColor: item.c6 }} /></td>
                    <td><span className="matrix-dot" style={{ backgroundColor: item.c7 }} /></td>
                    <td><span className="matrix-dot" style={{ backgroundColor: item.c8 }} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </article>
      </section>
    </main>
  );
};

