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

export interface ChatMessage {
  role: "user" | "assistant" | "system";
  content: string;
  timestamp?: string;
}

export interface ChatResponse {
  message: string;
  suggestions?: string[];
  insights?: Record<string, any>;
}

export interface JobRecord {
  id: string;
  filename: string;
  file_size: number;
  status: "queued" | "in_progress" | "completed" | "failed" | "discarded";
  progress: number;
  created_at: string;
  completed_at?: string | null;
  error_message?: string | null;
  temp_dir?: string | null;
  zip_path?: string | null;
  summary?: Record<string, any>;
}

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

const getApiBase = () => {
  return "";
};

async function request<T>(url: string, options?: RequestInit): Promise<T> {
  const fullUrl = `${getApiBase()}${url}`;
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 20000);
  try {
    const response = await fetch(fullUrl, {
      ...options,
      signal: controller.signal,
    });
    const text = await response.text();
    let result: any;
    try {
      result = JSON.parse(text);
    } catch {
      throw new ApiError(
        "API server is unavailable. Start the app with run.bat or npm run dev.",
        response.status || 503,
      );
    }
    if (!response.ok) {
      throw new ApiError(
        result.error || result.detail || "Request failed.",
        response.status,
        result.current,
      );
    }
    return result as T;
  } catch (error) {
    if (error instanceof ApiError) throw error;
    throw new ApiError(
      error instanceof Error && error.name === "AbortError"
        ? "The request timed out. Please try again."
        : "Cannot reach backend server. Ensure backend is running.",
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

export function deleteDocument(id: string): Promise<{ id: string; deleted: boolean }> {
  return request(`/api/documents/${encodeURIComponent(id)}`, {
    method: "DELETE",
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

// ------------------ Chat Assistant ------------------
export async function sendChatMessage(
  messages: ChatMessage[],
  query?: string,
  document?: OfficeDocument | null,
): Promise<ChatResponse> {
  return request<ChatResponse>("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ messages, query, document }),
  });
}

// ------------------ Jobs & Ingestion ------------------
export async function convertDocument(file: File): Promise<{ job_id: string; status: string }> {
  const formData = new FormData();
  formData.append("file", file, file.name);

  const fullUrl = `${getApiBase()}/api/convert`;
  const response = await fetch(fullUrl, {
    method: "POST",
    body: formData,
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Upload failed: ${errorText}`);
  }
  return response.json();
}

export function listJobs(): Promise<JobRecord[]> {
  return request<JobRecord[]>("/api/jobs");
}

export function getJob(jobId: string): Promise<JobRecord> {
  return request<JobRecord>(`/api/jobs/${encodeURIComponent(jobId)}`);
}

export function discardJob(jobId: string): Promise<{ job_id: string; status: string }> {
  return request(`/api/jobs/${encodeURIComponent(jobId)}/discard`, {
    method: "POST",
  });
}

export function downloadZipUrl(jobId: string): string {
  return `${getApiBase()}/api/jobs/${encodeURIComponent(jobId)}/download-zip`;
}
