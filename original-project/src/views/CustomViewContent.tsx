import React from "react";
import type { OfficeDocument } from "../types/office";
import type { CustomView } from "./preferences";
import { CopilotChatView } from "../components/rightView/CopilotChatView";
import "../styles/views.css";

export function CustomViewContent({
  view,
  activeDoc,
  onChangeNotes,
  storageError,
}: {
  view: CustomView;
  activeDoc: OfficeDocument | null;
  onChangeNotes: (notes: string) => void;
  storageError?: boolean;
}) {
  if (view.kind === "context") return <CopilotChatView activeDoc={activeDoc} />;
  return (
    <div className="custom-notes-view">
      <p>
        Personal notes for this application view. Saved automatically in this
        browser, separately from your documents.
      </p>
      <textarea
        aria-label={`${view.title} notes`}
        value={view.notes}
        maxLength={100000}
        placeholder="Write a checklist, reminders, or ideas…"
        onChange={(event) => onChangeNotes(event.target.value)}
      />
      <small role={storageError ? "alert" : undefined}>
        {storageError
          ? "Browser storage is unavailable. Notes will not survive a reload."
          : "Stored on this device · not included in document exports"}
      </small>
    </div>
  );
}
