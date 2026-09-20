import React, { useState, useEffect, useRef } from "react";
import {
  Send,
  Bot,
  User,
  Sparkles,
  RotateCcw,
  Copy,
  Check,
  FileText,
  FileSpreadsheet,
  Presentation,
  Code,
  AlertCircle,
  Loader2,
} from "lucide-react";
import type { OfficeDocument } from "../../types/office";
import { sendChatMessage, ChatMessage } from "../../services/api";

interface DocumentChatProps {
  activeDoc: OfficeDocument | null;
  onApplyAction?: (action: any) => void;
}

export const DocumentChat: React.FC<DocumentChatProps> = ({ activeDoc }) => {
  const [messages, setMessages] = useState<ChatMessage[]>(() => [
    {
      role: "assistant",
      content:
        "👋 **Welcome to Office Copilot!**\n\nI can analyze your active Word documents, Excel spreadsheets, and PowerPoint presentations. Ask me to summarize text, inspect spreadsheet formulas, outline slides, or suggest improvements.",
      timestamp: new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }),
    },
  ]);
  const [inputValue, setInputValue] = useState("");
  const [loading, setLoading] = useState(false);
  const [copiedIndex, setCopiedIndex] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [suggestions, setSuggestions] = useState<string[]>([
    "Summarize this document",
    "Explain calculations",
    "Review presentation outline",
  ]);

  const messagesEndRef = useRef<HTMLDivElement>(null);

  // Auto scroll to bottom
  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages, loading]);

  // Update suggestions when active document changes
  useEffect(() => {
    if (!activeDoc) {
      setSuggestions(["Create a document", "How do I upload files?", "What formats are supported?"]);
      return;
    }
    const type = activeDoc.type;
    if (type === "word") {
      setSuggestions([
        "Summarize document",
        "Analyze readability & word count",
        "Review review comments",
        "Suggest executive summary",
      ]);
    } else if (type === "excel") {
      setSuggestions([
        "Explain sheet formulas",
        "Analyze numbers & trends",
        "Suggest formula for total",
        "List all worksheets",
      ]);
    } else if (type === "powerpoint") {
      setSuggestions([
        "Presentation outline",
        "Review speaker notes",
        "Slide layout review",
        "Tips for presenting",
      ]);
    } else if (type === "code") {
      setSuggestions(["Analyze code structure", "Find functions", "Format file"]);
    }
  }, [activeDoc?.type, activeDoc?.data?.id]);

  const handleSend = async (textToSend?: string) => {
    const text = (textToSend || inputValue).trim();
    if (!text || loading) return;

    setError(null);
    const now = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    const userMsg: ChatMessage = { role: "user", content: text, timestamp: now };

    const newMessages = [...messages, userMsg];
    setMessages(newMessages);
    if (!textToSend) setInputValue("");
    setLoading(true);

    try {
      const response = await sendChatMessage(newMessages, text, activeDoc);
      const assistantMsg: ChatMessage = {
        role: "assistant",
        content: response.message,
        timestamp: new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }),
      };
      setMessages([...newMessages, assistantMsg]);
      if (response.suggestions && response.suggestions.length > 0) {
        setSuggestions(response.suggestions);
      }
    } catch (err: any) {
      setError(err.message || "Failed to get response from assistant.");
    } finally {
      setLoading(false);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const copyToClipboard = (text: string, index: number) => {
    navigator.clipboard.writeText(text);
    setCopiedIndex(index);
    setTimeout(() => setCopiedIndex(null), 2000);
  };

  const clearChat = () => {
    setMessages([
      {
        role: "assistant",
        content: "Chat cleared. What else would you like to explore about your documents?",
        timestamp: new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }),
      },
    ]);
  };

  const getDocIcon = () => {
    if (!activeDoc) return <Sparkles size={16} className="text-blue-400" />;
    switch (activeDoc.type) {
      case "word":
        return <FileText size={16} style={{ color: "#2b579a" }} />;
      case "excel":
        return <FileSpreadsheet size={16} style={{ color: "#217346" }} />;
      case "powerpoint":
        return <Presentation size={16} style={{ color: "#d24726" }} />;
      default:
        return <Code size={16} style={{ color: "#61afef" }} />;
    }
  };

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100%",
        backgroundColor: "var(--vscode-editor-bg, #1e1e1e)",
        color: "var(--vscode-editor-fg, #d4d4d4)",
        fontFamily: "var(--vscode-font-family, system-ui, sans-serif)",
      }}
    >
      {/* Header */}
      <div
        style={{
          padding: "10px 14px",
          borderBottom: "1px solid var(--vscode-panel-border, #333)",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          backgroundColor: "var(--vscode-sideBar-bg, #252526)",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {getDocIcon()}
          <div>
            <div style={{ fontWeight: 600, fontSize: 13, display: "flex", alignItems: "center", gap: 6 }}>
              Office Copilot
              <span
                style={{
                  fontSize: 10,
                  backgroundColor: "rgba(59, 130, 246, 0.2)",
                  color: "#60a5fa",
                  padding: "1px 6px",
                  borderRadius: 10,
                  fontWeight: 500,
                }}
              >
                AI
              </span>
            </div>
            <div style={{ fontSize: 11, color: "var(--vscode-tab-inactive-fg, #888)" }}>
              {activeDoc
                ? `${activeDoc.data.title || "Untitled"} (${activeDoc.type.toUpperCase()})`
                : "No document active"}
            </div>
          </div>
        </div>
        <button
          onClick={clearChat}
          title="Clear Conversation"
          style={{
            background: "none",
            border: "none",
            color: "var(--vscode-tab-inactive-fg, #888)",
            cursor: "pointer",
            padding: 4,
            borderRadius: 4,
            display: "flex",
            alignItems: "center",
          }}
        >
          <RotateCcw size={14} />
        </button>
      </div>

      {/* Messages Feed */}
      <div
        style={{
          flex: 1,
          overflowY: "auto",
          padding: "14px",
          display: "flex",
          flexDirection: "column",
          gap: 12,
        }}
      >
        {messages.map((m, idx) => {
          const isUser = m.role === "user";
          return (
            <div
              key={idx}
              style={{
                display: "flex",
                flexDirection: "column",
                alignItems: isUser ? "flex-end" : "flex-start",
                gap: 4,
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                  fontSize: 11,
                  color: "var(--vscode-tab-inactive-fg, #888)",
                  paddingLeft: isUser ? 0 : 4,
                  paddingRight: isUser ? 4 : 0,
                }}
              >
                {isUser ? <User size={12} /> : <Bot size={12} style={{ color: "#38bdf8" }} />}
                <span>{isUser ? "You" : "Copilot"}</span>
                {m.timestamp && <span>· {m.timestamp}</span>}
              </div>

              <div
                style={{
                  maxWidth: "90%",
                  padding: "9px 13px",
                  borderRadius: 8,
                  fontSize: 13,
                  lineHeight: 1.45,
                  backgroundColor: isUser
                    ? "var(--vscode-button-bg, #0e639c)"
                    : "var(--vscode-sideBar-bg, #252526)",
                  color: isUser ? "#ffffff" : "var(--vscode-editor-fg, #cccccc)",
                  border: isUser ? "none" : "1px solid var(--vscode-panel-border, #3a3a3a)",
                  whiteSpace: "pre-wrap",
                  wordBreak: "break-word",
                  position: "relative",
                }}
              >
                {m.content}
                {!isUser && (
                  <button
                    onClick={() => copyToClipboard(m.content, idx)}
                    title="Copy answer"
                    style={{
                      position: "absolute",
                      top: 6,
                      right: 6,
                      background: "transparent",
                      border: "none",
                      color: "var(--vscode-tab-inactive-fg, #888)",
                      cursor: "pointer",
                      padding: 2,
                    }}
                  >
                    {copiedIndex === idx ? <Check size={12} style={{ color: "#4ade80" }} /> : <Copy size={12} />}
                  </button>
                )}
              </div>
            </div>
          );
        })}

        {loading && (
          <div style={{ display: "flex", alignItems: "center", gap: 8, padding: 8, color: "#9ca3af" }}>
            <Loader2 size={16} className="animate-spin" style={{ animation: "spin 1s linear infinite" }} />
            <span style={{ fontSize: 12 }}>Analyzing document...</span>
          </div>
        )}

        {error && (
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              padding: "8px 12px",
              backgroundColor: "rgba(239, 68, 68, 0.15)",
              border: "1px solid #ef4444",
              borderRadius: 6,
              color: "#fca5a5",
              fontSize: 12,
            }}
          >
            <AlertCircle size={14} />
            <span>{error}</span>
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      {/* Quick Prompts Suggestions */}
      {suggestions.length > 0 && (
        <div
          style={{
            padding: "6px 12px",
            display: "flex",
            flexWrap: "wrap",
            gap: 6,
            borderTop: "1px solid var(--vscode-panel-border, #2d2d2d)",
            backgroundColor: "var(--vscode-sideBar-bg, #222)",
          }}
        >
          {suggestions.map((s, idx) => (
            <button
              key={idx}
              onClick={() => handleSend(s)}
              disabled={loading}
              style={{
                background: "var(--vscode-editor-bg, #1e1e1e)",
                border: "1px solid var(--vscode-panel-border, #3c3c3c)",
                color: "var(--vscode-editor-fg, #d4d4d4)",
                padding: "4px 9px",
                borderRadius: 14,
                fontSize: 11,
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                gap: 4,
              }}
            >
              <Sparkles size={10} style={{ color: "#38bdf8" }} />
              {s}
            </button>
          ))}
        </div>
      )}

      {/* Input Bar */}
      <div
        style={{
          padding: "10px 12px",
          borderTop: "1px solid var(--vscode-panel-border, #333)",
          backgroundColor: "var(--vscode-sideBar-bg, #252526)",
          display: "flex",
          alignItems: "center",
          gap: 8,
        }}
      >
        <input
          type="text"
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
          onKeyDown={handleKeyDown}
          placeholder={
            activeDoc
              ? `Ask about ${activeDoc.data.title || activeDoc.type}...`
              : "Ask anything about documents..."
          }
          disabled={loading}
          style={{
            flex: 1,
            backgroundColor: "var(--vscode-input-bg, #1e1e1e)",
            color: "var(--vscode-input-fg, #cccccc)",
            border: "1px solid var(--vscode-input-border, #3c3c3c)",
            padding: "8px 12px",
            borderRadius: 6,
            fontSize: 13,
            outline: "none",
          }}
        />
        <button
          onClick={() => handleSend()}
          disabled={loading || !inputValue.trim()}
          style={{
            backgroundColor: "var(--vscode-button-bg, #0e639c)",
            color: "#ffffff",
            border: "none",
            borderRadius: 6,
            padding: "8px 12px",
            cursor: loading || !inputValue.trim() ? "not-allowed" : "pointer",
            opacity: loading || !inputValue.trim() ? 0.5 : 1,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <Send size={15} />
        </button>
      </div>
    </div>
  );
};

