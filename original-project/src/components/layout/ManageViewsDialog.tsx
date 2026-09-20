import React, { useState } from "react";
import { ArrowDown, ArrowUp, Plus, Trash2 } from "lucide-react";
import { WorkspaceDialog } from "./WorkspaceDialog";
import type { ViewPreferencesController } from "../../views/useViewPreferences";
import { PROFILE_LABELS } from "../../views/preferences";
import type { CustomViewKind, ViewArea } from "../../views/preferences";
import "../../styles/views.css";

export function ManageViewsDialog({
  controller,
  onClose,
}: {
  controller: ViewPreferencesController;
  onClose: () => void;
}) {
  const [title, setTitle] = useState("");
  const [kind, setKind] = useState<CustomViewKind>("notes");
  const [area, setArea] = useState<ViewArea>("right");
  const [message, setMessage] = useState("");
  const [deleteId, setDeleteId] = useState<string | null>(null);
  return (
    <WorkspaceDialog title="Customize views" onClose={onClose}>
      <div className="manage-views">
        <p>
          Configure <strong>{PROFILE_LABELS[controller.profile]}</strong> views.
          Each application remembers its own arrangement on this device. Changes
          apply immediately.
        </p>
        {controller.storageError && (
          <p role="alert">
            Browser storage is unavailable; these preferences will only last for
            this session.
          </p>
        )}
        {(["left", "right"] as const).map((location) => {
          const visible = controller.preferences[location];
          const definitions = controller.availableViews(location);
          const ordered = [
            ...visible
              .map((id) => definitions.find((v) => v.id === id)!)
              .filter(Boolean),
            ...definitions.filter((v) => !visible.includes(v.id)),
          ];
          return (
            <section
              key={location}
              className="manage-views-section"
              aria-label={
                location === "left" ? "Explorer views" : "Right sidebar views"
              }
            >
              <h3>
                {location === "left"
                  ? "Explorer · left sidebar"
                  : "Tools · right sidebar"}
              </h3>
              {ordered.map((view) => {
                const index = visible.indexOf(view.id);
                return (
                  <div className="manage-view-row" key={view.id}>
                    <label>
                      <input
                        type="checkbox"
                        checked={index >= 0}
                        onChange={(event) =>
                          controller.setVisible(
                            location,
                            view.id,
                            event.target.checked,
                          )
                        }
                      />
                      <span>
                        <strong>{view.title}</strong>
                        <small>{view.description}</small>
                      </span>
                    </label>
                    <div className="manage-view-order">
                      <button
                        className="vscode-icon-btn"
                        title={`Move ${view.title} up`}
                        aria-label={`Move ${view.title} up`}
                        disabled={index <= 0}
                        onClick={() => controller.move(location, view.id, -1)}
                      >
                        <ArrowUp size={15} />
                      </button>
                      <button
                        className="vscode-icon-btn"
                        title={`Move ${view.title} down`}
                        aria-label={`Move ${view.title} down`}
                        disabled={index < 0 || index === visible.length - 1}
                        onClick={() => controller.move(location, view.id, 1)}
                      >
                        <ArrowDown size={15} />
                      </button>
                      {view.id.startsWith("custom:") && (
                        <button
                          className="vscode-icon-btn"
                          title={`Delete ${view.title}`}
                          aria-label={`Delete ${view.title}`}
                          onClick={() => setDeleteId(view.id)}
                        >
                          <Trash2 size={14} />
                        </button>
                      )}
                    </div>
                  </div>
                );
              })}
            </section>
          );
        })}
        {deleteId && (
          <div className="view-delete-confirm" role="alert">
            <p>
              Delete “
              {
                controller.preferences.custom.find(
                  (view) => view.id === deleteId,
                )?.title
              }
              ” and its saved notes? This cannot be undone. Uncheck the view
              above to hide it without losing notes.
            </p>
            <button
              className="vscode-btn-secondary"
              onClick={() => setDeleteId(null)}
            >
              Cancel deletion
            </button>{" "}
            <button
              className="vscode-btn-primary"
              onClick={() => {
                controller.removeCustomView(deleteId);
                setDeleteId(null);
                setMessage("Custom view deleted.");
              }}
            >
              Delete view and notes
            </button>
          </div>
        )}
        <form
          className="new-view-form"
          onSubmit={(event) => {
            event.preventDefault();
            const id = controller.addView({ title, kind, area });
            if (id) {
              setMessage(
                `Added ${title.trim()} to the ${area === "left" ? "left" : "right"} sidebar.`,
              );
              setTitle("");
            } else
              setMessage(
                "Enter a view name. Each application supports up to 24 custom views.",
              );
          }}
        >
          <h3>Add a new view</h3>
          <label>
            View name
            <input
              className="vscode-input"
              value={title}
              maxLength={60}
              placeholder="e.g. Editing checklist"
              onChange={(event) => setTitle(event.target.value)}
              required
            />
          </label>
          <div className="new-view-options">
            <label>
              Content
              <select
                className="vscode-input"
                value={kind}
                onChange={(event) =>
                  setKind(event.target.value as CustomViewKind)
                }
              >
                <option value="notes">Personal notes</option>
                <option value="context">Live document context</option>
              </select>
            </label>
            <label>
              Location
              <select
                className="vscode-input"
                value={area}
                onChange={(event) => setArea(event.target.value as ViewArea)}
              >
                <option value="left">Left sidebar</option>
                <option value="right">Right sidebar</option>
              </select>
            </label>
          </div>
          <p>
            Notes stay in this browser and are not part of saved or exported
            documents. Hide a view by unchecking it; its notes are retained.
          </p>
          <button
            className="vscode-btn-primary"
            type="submit"
            disabled={
              !title.trim() || controller.preferences.custom.length >= 24
            }
          >
            <Plus size={14} /> Add view
          </button>
          <span className="view-status" role="status">
            {message}
          </span>
        </form>
        <div className="dialog-actions">
          <button
            className="vscode-btn-secondary"
            onClick={() => {
              controller.resetViews();
              setMessage(
                "Default views restored. Custom views and their notes are retained and can be enabled above.",
              );
            }}
          >
            Restore default views
          </button>
          <button className="vscode-btn-primary" onClick={onClose}>
            Done
          </button>
        </div>
      </div>
    </WorkspaceDialog>
  );
}
