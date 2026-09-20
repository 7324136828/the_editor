import React, { useState } from "react";
import { ChevronDown, ChevronRight, X } from "lucide-react";

interface AccordionSectionProps {
  id: string;
  title: string;
  badge?: string | number;
  defaultExpanded?: boolean;
  onHide?: () => void;
  actions?: {
    icon: React.ReactNode;
    title: string;
    onClick: (e: React.MouseEvent) => void;
  }[];
  children: React.ReactNode;
}

export const AccordionSection: React.FC<AccordionSectionProps> = ({
  id,
  title,
  badge,
  defaultExpanded = true,
  actions = [],
  onHide,
  children,
}) => {
  const [expanded, setExpanded] = useState(defaultExpanded);

  return (
    <div className="vscode-accordion-section">
      <div className="vscode-accordion-header">
        <button
          className="vscode-accordion-toggle"
          aria-expanded={expanded}
          aria-controls={`view-body-${id}`}
          onClick={() => setExpanded(!expanded)}
        >
          <div className="vscode-accordion-left">
            {expanded ? (
              <ChevronDown size={14} color="#bbbbbb" />
            ) : (
              <ChevronRight size={14} color="#bbbbbb" />
            )}
            <span>{title}</span>
            {badge !== undefined && (
              <span className="vscode-accordion-badge">{badge}</span>
            )}
          </div>
        </button>

        {(actions.length > 0 || onHide) && (
          <div
            className="vscode-accordion-actions"
            onClick={(e) => e.stopPropagation()}
          >
            {actions.map((act, idx) => (
              <button
                key={idx}
                className="vscode-icon-btn"
                style={{ width: 20, height: 20 }}
                title={act.title}
                onClick={act.onClick}
              >
                {act.icon}
              </button>
            ))}
            {onHide && (
              <button
                className="vscode-icon-btn"
                style={{ width: 20, height: 20 }}
                title={`Hide ${title}`}
                aria-label={`Hide ${title}`}
                onClick={onHide}
              >
                <X size={12} />
              </button>
            )}
          </div>
        )}
      </div>

      {expanded && (
        <div className="vscode-accordion-body" id={`view-body-${id}`}>
          {children}
        </div>
      )}
    </div>
  );
};
