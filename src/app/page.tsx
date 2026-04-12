'use client';

import { useState, useEffect, useRef, useCallback } from 'react';
import { Send, Loader2, Plus, FileText, FileSpreadsheet, Download, Trash2, Upload, Moon, Sun, Copy, Check, Sparkles, Eye, EyeOff, ChevronLeft, ChevronRight, Square, GripVertical } from 'lucide-react';
import { cn } from '@/lib/utils';

// ─── TYPES ────────────────────────────────────────────────────────────────────

interface Message {
  id: string;
  role: 'user' | 'assistant';
  content: string;
  filenames?: string[];
  createdAt: string;
}

interface Conversation {
  id: string;
  title: string;
  messages: Message[];
  files: string[];
  createdAt: string;
}

interface FileInfo {
  name: string;
  type: string;
  size: number;
  createdAt: string;
  updatedAt: string;
}

interface PreviewData {
  type: 'docx' | 'xlsx';
  html?: string;
  sheets?: Record<string, string>;
  activeSheet?: string;
  filename: string;
}

// ─── PANEL WIDTHS ─────────────────────────────────────────────────────────────

function loadWidths(): { sidebar: number; preview: number; files: number } {
  try {
    const saved = localStorage.getItem('oai-panel-widths');
    if (saved) return JSON.parse(saved);
  } catch { }
  return { sidebar: 220, preview: 320, files: 200 };
}

// ─── COMPONENT ────────────────────────────────────────────────────────────────

export default function OfficeAI() {
  const [conversations, setConversations] = useState<Conversation[]>([]);
  const [activeId, setActiveId] = useState<string | null>(null);
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState('');
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState('');
  const [allFiles, setAllFiles] = useState<FileInfo[]>([]);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [theme, setTheme] = useState<'dark' | 'light'>('dark');
  const [ready, setReady] = useState(false);
  const [copied, setCopied] = useState<string | null>(null);
  const [uploading, setUploading] = useState(false);
  const [previewFile, setPreviewFile] = useState<string | null>(null);
  const [previewData, setPreviewData] = useState<PreviewData | null>(null);
  const [activeSheet, setActiveSheet] = useState<string | null>(null);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [panelWidths, setPanelWidths] = useState(loadWidths());
  const endRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const abortRef = useRef<AbortController | null>(null);

  // Get files for current conversation
  const activeConversation = conversations.find(c => c.id === activeId);
  const conversationFiles = activeConversation?.files || [];
  const displayFiles = allFiles.filter(f => conversationFiles.includes(f.name));

  // ─── INIT ────────────────────────────────────────────────────────────────

  useEffect(() => {
    setReady(true);
    const savedTheme = localStorage.getItem('oai-theme') as 'dark' | 'light' | null;
    if (savedTheme) setTheme(savedTheme);

    const saved = localStorage.getItem('oai-conversations');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (parsed?.length > 0) {
          // Migrate old conversations without files array
          const migrated = parsed.map((c: any) => ({ ...c, files: c.files || [] }));
          setConversations(migrated);
          setActiveId(migrated[0].id);
          setMessages(migrated[0].messages || []);
        } else {
          newChat();
        }
      } catch { newChat(); }
    } else {
      newChat();
    }
    loadFiles();
  }, []);

  useEffect(() => {
    if (ready) {
      localStorage.setItem('oai-theme', theme);
      document.documentElement.classList.toggle('light', theme === 'light');
      document.documentElement.classList.toggle('dark', theme === 'dark');
    }
  }, [theme, ready]);

  useEffect(() => {
    if (ready && conversations.length > 0) {
      localStorage.setItem('oai-conversations', JSON.stringify(conversations));
    }
  }, [conversations, ready]);

  useEffect(() => {
    if (ready) {
      localStorage.setItem('oai-panel-widths', JSON.stringify(panelWidths));
    }
  }, [panelWidths, ready]);

  useEffect(() => {
    endRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, loading]);

  // ─── PREVIEW ─────────────────────────────────────────────────────────────

  useEffect(() => {
    if (previewFile) {
      loadPreview(previewFile);
    }
  }, [previewFile]);

  const loadPreview = async (filename: string) => {
    setPreviewLoading(true);
    try {
      const res = await fetch('/api/preview?filename=' + encodeURIComponent(filename));
      if (res.ok) {
        const data = await res.json();
        // Add element IDs to HTML
        if (data.html) {
          data.html = addElementIds(data.html);
        }
        if (data.sheets) {
          const newSheets: Record<string, string> = {};
          for (const [name, html] of Object.entries(data.sheets)) {
            newSheets[name] = addElementIds(html as string);
          }
          data.sheets = newSheets;
        }
        setPreviewData(data);
        if (data.activeSheet) setActiveSheet(data.activeSheet);
      }
    } catch { } finally {
      setPreviewLoading(false);
    }
  };

  // ─── ELEMENT IDs ─────────────────────────────────────────────────────────

  const addElementIds = (html: string): string => {
    let id = 0;
    // Add data-id to headings, paragraphs, tables, lists
    return html
      .replace(/<h([1-6])([^>]*)>/gi, (match, level, attrs) => {
        id++;
        return `<h${level}${attrs} data-eid="@${id}"><span class="eid-badge">@${id}</span>`;
      })
      .replace(/<p([^>]*)>/gi, (match, attrs) => {
        // Skip if inside a table
        id++;
        return `<p${attrs} data-eid="@${id}"><span class="eid-badge">@${id}</span>`;
      })
      .replace(/<table([^>]*)>/gi, (match, attrs) => {
        id++;
        return `<table${attrs} data-eid="@${id}"><caption class="eid-badge">Table @${id}</caption>`;
      });
  };

  // ─── HELPERS ─────────────────────────────────────────────────────────────

  const newChat = () => {
    const conv: Conversation = { id: Date.now().toString(), title: 'New Chat', messages: [], files: [], createdAt: new Date().toISOString() };
    setConversations(prev => [conv, ...prev]);
    setActiveId(conv.id);
    setMessages([]);
    setPreviewFile(null);
    setPreviewData(null);
  };

  const addFileToConversation = (filename: string) => {
    setConversations(prev => prev.map(c =>
      c.id === activeId && !c.files.includes(filename)
        ? { ...c, files: [...c.files, filename] }
        : c
    ));
  };

  const loadFiles = async () => {
    try {
      const res = await fetch('/api/files');
      if (res.ok) { const data = await res.json(); setAllFiles(data.files || []); }
    } catch { }
  };

  const formatSize = (b: number) => b < 1024 ? b + ' B' : b < 1048576 ? (b / 1024).toFixed(1) + ' KB' : (b / 1048576).toFixed(1) + ' MB';

  const handleUpload = async (file: File) => {
    const ext = file.name.split('.').pop()?.toLowerCase();
    if (ext !== 'docx' && ext !== 'xlsx') { alert('Only .docx and .xlsx files'); return; }
    setUploading(true);
    try {
      const fd = new FormData(); fd.append('file', file);
      const res = await fetch('/api/upload', { method: 'POST', body: fd });
      if (res.ok) {
        const data = await res.json();
        loadFiles();
        addFileToConversation(data.file.name);
        setPreviewFile(data.file.name);
        setInput('Analyze ' + data.file.name);
        inputRef.current?.focus();
      }
    } catch { } finally { setUploading(false); }
  };

  const copyMsg = (id: string, text: string) => {
    navigator.clipboard.writeText(text);
    setCopied(id);
    setTimeout(() => setCopied(null), 2000);
  };

  // ─── RESIZE HANDLERS ────────────────────────────────────────────────────

  const startResize = useCallback((panel: 'sidebar' | 'preview' | 'files') => {
    return (e: React.MouseEvent) => {
      e.preventDefault();
      const startX = e.clientX;
      const startWidth = panelWidths[panel];

      const onMouseMove = (ev: MouseEvent) => {
        const delta = ev.clientX - startX;
        if (panel === 'sidebar') {
          setPanelWidths(prev => ({ ...prev, sidebar: Math.max(160, Math.min(400, startWidth + delta)) }));
        } else if (panel === 'preview') {
          setPanelWidths(prev => ({ ...prev, preview: Math.max(200, Math.min(600, startWidth - delta)) }));
        } else {
          setPanelWidths(prev => ({ ...prev, files: Math.max(140, Math.min(350, startWidth - delta)) }));
        }
      };

      const onMouseUp = () => {
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
      };

      document.addEventListener('mousemove', onMouseMove);
      document.addEventListener('mouseup', onMouseUp);
      document.body.style.cursor = 'col-resize';
      document.body.style.userSelect = 'none';
    };
  }, [panelWidths]);

  // ─── SUBMIT ──────────────────────────────────────────────────────────────

  const stopRequest = () => {
    if (abortRef.current) {
      abortRef.current.abort();
      abortRef.current = null;
    }
    setLoading(false);
    setStatus('');
    setMessages(prev => prev.filter(m => !m.id.startsWith('t')));
  };

  const submit = async () => {
    if (!input.trim() || loading) return;
    const msg = input.trim();
    setInput('');
    setLoading(true);
    setStatus('Thinking...');

    const controller = new AbortController();
    abortRef.current = controller;

    const tempMsg: Message = { id: 't' + Date.now(), role: 'user', content: msg, createdAt: new Date().toISOString() };
    setMessages(prev => [...prev, tempMsg]);

    try {
      const history = messages.filter(m => !m.id.startsWith('t')).slice(-10).map(m => ({ role: m.role, content: m.content, filenames: m.filenames || [] }));
      const res = await fetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message: msg, history }),
        signal: controller.signal,
      });

      const reader = res.body?.getReader();
      const decoder = new TextDecoder();
      if (!reader) throw new Error('No stream');

      let buffer = '';
      while (true) {
        if (controller.signal.aborted) break;
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
          if (line.startsWith('data: ')) {
            try {
              const data = JSON.parse(line.slice(6));
              if (data.type === 'progress') setStatus(data.status);
              else if (data.type === 'complete') {
                setStatus('');
                const newMsgs = [data.userMessage, data.assistantMessage];
                setMessages(prev => prev.filter(m => !m.id.startsWith('t')).concat(newMsgs));
                setConversations(prev => prev.map(c => {
                  if (c.id !== activeId) return c;
                  // Add files from this response
                  const newFiles = [...c.files];
                  for (const f of (data.assistantMessage?.filenames || [])) {
                    if (!newFiles.includes(f)) newFiles.push(f);
                  }
                  return { ...c, messages: [...c.messages.filter(m => m.id), ...newMsgs], files: newFiles, title: msg.slice(0, 40) };
                }));
                loadFiles();
                if (data.assistantMessage?.filenames?.length > 0) {
                  const fname = data.assistantMessage.filenames[0];
                  addFileToConversation(fname);
                  setPreviewFile(fname);
                }
              } else if (data.type === 'error') {
                setStatus('');
                setMessages(prev => prev.filter(m => !m.id.startsWith('t')).concat([{ id: 'e' + Date.now(), role: 'assistant', content: 'Error: ' + (data.error || 'Unknown'), createdAt: new Date().toISOString() }]));
              }
            } catch { }
          }
        }
      }
    } catch (e: any) {
      if (e.name !== 'AbortError') {
        setStatus('');
        setMessages(prev => prev.filter(m => !m.id.startsWith('t')).concat([{ id: 'e' + Date.now(), role: 'assistant', content: 'Connection error.', createdAt: new Date().toISOString() }]));
      }
    } finally {
      abortRef.current = null;
      setLoading(false);
    }
  };

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); submit(); }
  };

  // ─── RENDER ──────────────────────────────────────────────────────────────

  if (!ready) return <div className="loading"><img src="/brand.png" alt="" style={{width: 20, height: 20, objectFit: 'contain'}} /><span>Loading...</span></div>;

  return (
    <div className={cn('app', theme)}>
      {/* SIDEBAR */}
      {sidebarOpen && (
        <aside className="sidebar" style={{ width: panelWidths.sidebar }}>
          <div className="sidebar-top">
            <div className="brand">
              <div className="brand-icon"><img src="/brand.png" alt="OfficeAI" /></div>
              <span className="brand-name">OfficeAI</span>
            </div>
            <button className="btn-new" onClick={newChat}><Plus size={14} /> New</button>
          </div>
          <div className="conv-list">
            {conversations.map(c => (
              <div key={c.id} className={cn('conv', c.id === activeId && 'active')} onClick={() => { setActiveId(c.id); setMessages(c.messages); setPreviewFile(null); setPreviewData(null); }}>
                <span className="conv-title">{c.title}</span>
                <button className="conv-del" onClick={e => {
                  e.stopPropagation();
                  setConversations(prev => prev.filter(x => x.id !== c.id));
                  if (c.id === activeId && conversations.length > 1) {
                    const next = conversations.find(x => x.id !== c.id);
                    if (next) { setActiveId(next.id); setMessages(next.messages); }
                  }
                }}><Trash2 size={11} /></button>
              </div>
            ))}
          </div>
          <div className="sidebar-bottom">
            <span className="sidebar-bottom-label">Theme</span>
            <button className="btn-icon" onClick={() => setTheme(t => t === 'dark' ? 'light' : 'dark')} title={theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode'}>
              {theme === 'dark' ? <Sun size={14} /> : <Moon size={14} />}
            </button>
          </div>
        </aside>
      )}

      {/* MAIN */}
      <main className="main">
        <header className="header">
          <button className="btn-icon" onClick={() => setSidebarOpen(v => !v)}>
            {sidebarOpen ? <ChevronLeft size={16} /> : <ChevronRight size={16} />}
          </button>
          {previewFile && <span className="header-file">{previewFile}</span>}
          <div />
        </header>

        <div className="body">
          {/* CHAT */}
          <div className="chat">
            <div className="msgs">
              {messages.length === 0 && (
                <div className="welcome">
                  <div className="welcome-icon"><img src="/brand.png" alt="OfficeAI" style={{width: 32, height: 32, objectFit: 'contain'}} /></div>
                  <h1>What can I build for you?</h1>
                  <p>Create documents, edit spreadsheets, analyze files.</p>
                  <div className="prompts">
                    {[
                      { label: 'Pie Chart Excel', prompt: 'Create an Excel file with budget data (Marketing 30%, Development 25%, Operations 20%, Sales 15%, HR 10%) and embed a pie chart' },
                      { label: 'Bar Chart Excel', prompt: 'Create a quarterly revenue spreadsheet with Q1=$50000, Q2=$75000, Q3=$90000, Q4=$120000 and add a bar chart' },
                      { label: 'Report with Charts', prompt: 'Create a Word document "Annual Sales Report" with sections for Executive Summary, Quarterly Performance table, and embed a bar chart showing the data' },
                      { label: 'Budget Spreadsheet', prompt: 'Create a monthly budget tracker with columns for Category, Budget, Actual, Difference, and formulas. Add currency formatting.' },
                    ].map((p, i) => (
                      <button key={i} className="prompt-btn" onClick={() => { setInput(p.prompt); inputRef.current?.focus(); }}>
                        {p.label}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              {messages.map(m => (
                <div key={m.id} className={cn('msg', m.role)}>
                  <div className={cn('avatar', m.role)}>{m.role === 'user' ? 'You' : 'AI'}</div>
                  <div className="bubble">
                    <div className="text">{m.content}</div>
                    {m.filenames && m.filenames.length > 0 && (
                      <div className="files">
                        {m.filenames.map((f, i) => (
                          <button key={i} className="file-chip" onClick={() => { addFileToConversation(f); setPreviewFile(f); }}>
                            {f.endsWith('.xlsx') ? <FileSpreadsheet size={12} /> : <FileText size={12} />}
                            <span>{f}</span>
                          </button>
                        ))}
                      </div>
                    )}
                    {m.role === 'assistant' && (
                      <div className="actions">
                        <button onClick={() => copyMsg(m.id, m.content)}>{copied === m.id ? <Check size={11} /> : <Copy size={11} />}</button>
                      </div>
                    )}
                  </div>
                </div>
              ))}

              {loading && (
                <div className="msg assistant">
                  <div className="avatar assistant">AI</div>
                  <div className="bubble progress">
                    <Loader2 size={13} className="spin" />
                    <span>{status || 'Processing...'}</span>
                    <div className="dots"><span /><span /><span /></div>
                  </div>
                </div>
              )}
              <div ref={endRef} />
            </div>

            {/* INPUT */}
            <div className="input-bar">
              <form onSubmit={e => { e.preventDefault(); submit(); }} className="input-form">
                <textarea
                  ref={inputRef}
                  value={input}
                  onChange={e => setInput(e.target.value)}
                  onKeyDown={onKeyDown}
                  placeholder="Ask me to create or edit documents..."
                  rows={1}
                  disabled={false}
                  className="input-textarea"
                />
                <input ref={fileInputRef} type="file" onChange={e => { const f = e.target.files?.[0]; if (f) handleUpload(f); e.target.value = ''; }} accept=".docx,.xlsx" style={{ display: 'none' }} />
                <button type="button" className="btn-input" onClick={() => fileInputRef.current?.click()} disabled={uploading} title="Upload">
                  {uploading ? <Loader2 size={14} className="spin" /> : <Upload size={14} />}
                </button>
                {loading ? (
                  <button type="button" className="btn-stop" onClick={stopRequest} title="Stop">
                    <Square size={12} fill="currentColor" />
                  </button>
                ) : (
                  <button type="submit" disabled={!input.trim()} className="btn-send">
                    <Send size={14} />
                  </button>
                )}
              </form>
            </div>
          </div>

          {/* RESIZE HANDLE */}
          <div className="resize-handle" onMouseDown={startResize('preview')}>
            <GripVertical size={12} />
          </div>

          {/* PREVIEW PANEL */}
          <aside className="preview-panel" style={{ width: panelWidths.preview }}>
            <div className="preview-header">
              <span className="preview-title">
                <Eye size={13} /> Preview
              </span>
              {previewFile && (
                <button className="btn-icon" onClick={() => { setPreviewFile(null); setPreviewData(null); }}>
                  <EyeOff size={13} />
                </button>
              )}
            </div>

            <div className="preview-body">
              {!previewFile && (
                <div className="preview-empty">
                  <FileText size={24} />
                  <span>Select a file to preview</span>
                </div>
              )}

              {previewLoading && (
                <div className="preview-loading">
                  <Loader2 size={20} className="spin" />
                  <span>Loading preview...</span>
                </div>
              )}

              {!previewLoading && previewData?.type === 'docx' && previewData.html && (
                <div className="preview-content" dangerouslySetInnerHTML={{ __html: previewData.html }} />
              )}

              {!previewLoading && previewData?.type === 'xlsx' && previewData.sheets && (
                <div className="preview-excel">
                  <div className="sheet-tabs">
                    {Object.keys(previewData.sheets).map(name => (
                      <button key={name} className={cn('sheet-tab', name === activeSheet && 'active')} onClick={() => setActiveSheet(name)}>
                        {name}
                      </button>
                    ))}
                  </div>
                  <div className="preview-content" dangerouslySetInnerHTML={{ __html: previewData.sheets[activeSheet || Object.keys(previewData.sheets)[0]] || '' }} />
                </div>
              )}
            </div>
          </aside>

          {/* RESIZE HANDLE */}
          <div className="resize-handle" onMouseDown={startResize('files')}>
            <GripVertical size={12} />
          </div>

          {/* FILES PANEL - SHOWS ONLY CONVERSATION FILES */}
          <aside className="files-panel" style={{ width: panelWidths.files }}>
            <div className="files-header">
              <span>Files ({displayFiles.length})</span>
            </div>
            <div className="files-body">
              <button className="btn-upload" onClick={() => fileInputRef.current?.click()} disabled={uploading}>
                {uploading ? <><Loader2 size={12} className="spin" /> Uploading...</> : <><Upload size={12} /> Upload</>}
              </button>
              {displayFiles.length === 0 ? (
                <div className="empty">
                  {conversationFiles.length === 0 ? 'No files in this conversation' : 'Loading...'}
                </div>
              ) : (
                displayFiles.map(f => (
                  <div key={f.name} className={cn('file-row', f.name === previewFile && 'active')} onClick={() => setPreviewFile(f.name)}>
                    <div className={cn('file-type', f.type)}>{f.type === 'xlsx' ? <FileSpreadsheet size={12} /> : <FileText size={12} />}</div>
                    <div className="file-info">
                      <div className="file-name">{f.name}</div>
                      <div className="file-size">{formatSize(f.size)}</div>
                    </div>
                    <div className="file-actions">
                      <button onClick={e => { e.stopPropagation(); window.open('/api/files?action=download&filename=' + f.name, '_blank'); }} title="Download"><Download size={11} /></button>
                      <button onClick={e => { e.stopPropagation(); if (confirm('Delete ' + f.name + '?')) { fetch('/api/files?filename=' + f.name, { method: 'DELETE' }).then(() => { loadFiles(); setConversations(prev => prev.map(c => ({ ...c, files: c.files.filter(x => x !== f.name) }))); if (previewFile === f.name) { setPreviewFile(null); setPreviewData(null); } }); } }} title="Delete" className="danger"><Trash2 size={11} /></button>
                    </div>
                  </div>
                ))
              )}
            </div>
          </aside>
        </div>
      </main>
    </div>
  );
}
