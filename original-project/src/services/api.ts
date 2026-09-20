import type { OfficeDocument } from "../types/office";

export interface StoredDocument {
  id: string;
  name: string;
  document: OfficeDocument;
  revision: number;
  updatedAt: string;
}
export type DocumentVersion = Pick<
  StoredDocument,
  "revision" | "name" | "updatedAt"
>;

export class ApiError extends Error {
  constructor(
    message: string,
    public status: number,
    public current?: StoredDocument | null,
  ) {
    super(message);
    this.name = "ApiError";
  }
}

async function request<T>(url: string, options?: RequestInit): Promise<T> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);
  try {
    const response = await fetch(url, {
      ...options,
      signal: controller.signal,
    });
    const text = await response.text();
    let result;
    try {
      result = JSON.parse(text);
    } catch {
      throw new ApiError(
        "Storage server is unavailable. Start the app with npm run dev.",
        response.status || 503,
      );
    }
    if (!response.ok)
      throw new ApiError(
        result.error || "Document request failed.",
        response.status,
        result.current,
      );
    return result as T;
  } catch (error) {
    if (error instanceof ApiError) throw error;
    throw new ApiError(
      error instanceof Error && error.name === "AbortError"
        ? "The storage request timed out. Your edits remain open; try saving again."
        : "Cannot reach local storage. Start the app with npm run dev, then save again.",
      0,
    );
  } finally {
    clearTimeout(timeout);
  }
}

export async function listDocuments(): Promise<StoredDocument[]> {
  return (await request<{ documents: StoredDocument[] }>("/api/documents"))
    .documents;
}

export function saveDocument(
  id: string,
  name: string,
  document: OfficeDocument,
  revision: number | null,
): Promise<StoredDocument> {
  return request(`/api/documents/${encodeURIComponent(id)}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, document, revision }),
  });
}

export async function getDocumentHistory(
  id: string,
): Promise<DocumentVersion[]> {
  return (
    await request<{ versions: DocumentVersion[] }>(
      `/api/documents/${encodeURIComponent(id)}/history`,
    )
  ).versions;
}

export function getDocumentVersion(
  id: string,
  revision: number,
): Promise<StoredDocument> {
  return request(
    `/api/documents/${encodeURIComponent(id)}/history/${revision}`,
  );
}
