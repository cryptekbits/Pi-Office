import type { ConnectorCatalogItem } from "@pi-office/pi-office-pack";
import { CONNECTOR_CATALOG } from "./connector-catalog-generated";

export function listConnectorCatalog(): ConnectorCatalogItem[] {
  return [...CONNECTOR_CATALOG].map((item) => ({
    ...item,
    setupProfiles: item.setupProfiles ? item.setupProfiles.map((profile) => ({ ...profile })) : undefined,
  }));
}

export function getConnectorCatalogItem(connectorId: string): ConnectorCatalogItem | undefined {
  return listConnectorCatalog().find((item) => item.id === connectorId);
}
