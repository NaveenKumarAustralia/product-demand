// Vision Board V2 — UI components shared by the in-portal page render.
// All intents posted by these components are prefixed `vb_` and handled by
// the action in app/routes/portal._index.tsx.

import { memo, useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { useFetcher, useSearchParams } from "react-router";

// ─── Types ───────────────────────────────────────────────────────────────────

export type VisionBoardListItem = { id: number; name: string; sortOrder: number };
export type VisionItemListItem = {
  id: number;
  name: string;
  sortOrder: number;
  imageCount: number;
  hasThumbnail: boolean;
  updatedAt: string | Date;
};
type VisionField = { id: string; text: string };

const CARD_THUMB_BASE = "/portal/thumbnail/visionV2";
const ITEM_IMAGE_BASE = "/portal/image/visionV2";

// Higher quality than the legacy page: 1200 px max dim, JPEG q=0.85,
// 800 KB target. Visibly sharper on retina screens, still small enough
// for fast cards and drawer previews.
const FULL_MAX_DIM = 1200;
const FULL_QUALITY = 0.85;
const FULL_TARGET_BYTES = 800 * 1024;
const THUMB_MAX_DIM = 256;
const THUMB_QUALITY = 0.7;

// ─── Image helpers ───────────────────────────────────────────────────────────

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result));
    r.onerror = () => reject(r.error);
    r.readAsDataURL(file);
  });
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("decode failed"));
    img.src = src;
  });
}

function dataUrlBytes(url: string): number {
  const i = url.indexOf(",");
  if (i < 0) return url.length;
  const b64 = url.slice(i + 1);
  const pad = b64.endsWith("==") ? 2 : b64.endsWith("=") ? 1 : 0;
  return Math.max(0, Math.floor((b64.length * 3) / 4) - pad);
}

function renderToJpeg(img: HTMLImageElement, maxDim: number, quality: number): string | null {
  const longEdge = Math.max(img.width, img.height);
  const scale = longEdge > maxDim ? maxDim / longEdge : 1;
  const w = Math.max(1, Math.round(img.width * scale));
  const h = Math.max(1, Math.round(img.height * scale));
  const canvas = document.createElement("canvas");
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  ctx.drawImage(img, 0, 0, w, h);
  try { return canvas.toDataURL("image/jpeg", quality); } catch { return null; }
}

async function compressUpload(file: File): Promise<string> {
  if (!file.type.startsWith("image/")) return readFileAsDataUrl(file);
  const original = await readFileAsDataUrl(file);
  if (file.size <= FULL_TARGET_BYTES) return original;
  try {
    const img = await loadImage(original);
    let out = renderToJpeg(img, FULL_MAX_DIM, FULL_QUALITY);
    if (out && dataUrlBytes(out) > FULL_TARGET_BYTES) out = renderToJpeg(img, 1024, 0.78) ?? out;
    if (out && dataUrlBytes(out) > FULL_TARGET_BYTES) out = renderToJpeg(img, 900, 0.72) ?? out;
    return out ?? original;
  } catch {
    return original;
  }
}

async function makeThumb(dataUrl: string): Promise<string | null> {
  if (!dataUrl.startsWith("data:image/")) return null;
  try {
    const img = await loadImage(dataUrl);
    return renderToJpeg(img, THUMB_MAX_DIM, THUMB_QUALITY);
  } catch {
    return null;
  }
}

function fieldsFromUnknown(value: unknown): VisionField[] {
  if (!Array.isArray(value)) return [];
  return (value as Array<Record<string, unknown>>).map((f) => {
    const id = typeof f?.id === "string" ? f.id : `f_${Math.random().toString(36).slice(2, 9)}`;
    if (typeof f?.text === "string") return { id, text: f.text };
    const label = typeof f?.label === "string" ? f.label : "";
    const value = typeof f?.value === "string" ? f.value : "";
    const text = label && value ? `${label}: ${value}` : (label || value || "");
    return { id, text };
  });
}

// ─── Panel ───────────────────────────────────────────────────────────────────

export function VisionBoardV2Panel({
  boards,
  activeBoardId,
  items,
}: {
  boards: VisionBoardListItem[];
  activeBoardId: number | null;
  items: VisionItemListItem[];
}) {
  const [searchParams, setSearchParams] = useSearchParams();
  const fetcher = useFetcher();

  const [renamingBoardId, setRenamingBoardId] = useState<number | null>(null);
  const [renameDraft, setRenameDraft] = useState("");
  const [confirmDeleteBoardId, setConfirmDeleteBoardId] = useState<number | null>(null);
  const [addBoardOpen, setAddBoardOpen] = useState(false);
  const [addBoardName, setAddBoardName] = useState("");
  const [addItemOpen, setAddItemOpen] = useState(false);
  const [addItemName, setAddItemName] = useState("");
  const [selectedItemId, setSelectedItemId] = useState<number | null>(null);
  const [dragItemId, setDragItemId] = useState<number | null>(null);
  const [dragOverItemId, setDragOverItemId] = useState<number | null>(null);

  const setActiveBoard = (id: number) => {
    const next = new URLSearchParams(searchParams);
    next.set("boardId", String(id));
    setSearchParams(next, { replace: false });
  };

  const submitAddBoard = () => {
    const name = addBoardName.trim() || "New Board";
    fetcher.submit({ intent: "vb_add_board", name }, { method: "post" });
    setAddBoardOpen(false);
    setAddBoardName("");
  };
  const submitRenameBoard = (boardId: number) => {
    const name = renameDraft.trim();
    if (!name) { setRenamingBoardId(null); return; }
    fetcher.submit({ intent: "vb_rename_board", boardId: String(boardId), name }, { method: "post" });
    setRenamingBoardId(null);
  };
  const submitDeleteBoard = (boardId: number) => {
    fetcher.submit({ intent: "vb_delete_board", boardId: String(boardId) }, { method: "post" });
    setConfirmDeleteBoardId(null);
    if (activeBoardId === boardId) {
      const next = boards.find((b) => b.id !== boardId);
      const params = new URLSearchParams(searchParams);
      if (next) params.set("boardId", String(next.id));
      else params.delete("boardId");
      setSearchParams(params, { replace: true });
    }
  };

  const submitAddItem = () => {
    if (!activeBoardId) return;
    const name = addItemName.trim() || "Untitled";
    fetcher.submit({ intent: "vb_add_item", boardId: String(activeBoardId), name }, { method: "post" });
    setAddItemOpen(false);
    setAddItemName("");
  };

  const handleDeleteItem = (itemId: number) => {
    fetcher.submit({ intent: "vb_delete_item", itemId: String(itemId) }, { method: "post" });
    if (selectedItemId === itemId) setSelectedItemId(null);
  };

  const reorder = (targetId: number) => {
    if (!dragItemId || dragItemId === targetId) return;
    const fromIdx = items.findIndex((it) => it.id === dragItemId);
    const toIdx = items.findIndex((it) => it.id === targetId);
    if (fromIdx < 0 || toIdx < 0) return;
    const next = [...items];
    const [moved] = next.splice(fromIdx, 1);
    next.splice(toIdx, 0, moved);
    fetcher.submit({ intent: "vb_reorder_items", ids: JSON.stringify(next.map((it) => it.id)) }, { method: "post" });
  };

  const handleDropImageOnEmptyCard = async (file: File) => {
    if (!activeBoardId) return;
    const dataUrl = await compressUpload(file);
    const thumb = await makeThumb(dataUrl);
    const payload: Record<string, string> = {
      intent: "vb_add_item",
      boardId: String(activeBoardId),
      name: "Untitled",
      image: dataUrl,
    };
    if (thumb) payload.thumbnail = thumb;
    fetcher.submit(payload, { method: "post" });
  };

  const selectedItem = useMemo(
    () => items.find((it) => it.id === selectedItemId) ?? null,
    [items, selectedItemId],
  );

  return (
    <div style={S.page}>
      <div style={S.tabBar}>
        {boards.map((board) => {
          const isActive = board.id === activeBoardId;
          return (
            <div
              key={board.id}
              onClick={() => setActiveBoard(board.id)}
              style={{ ...S.tab, ...(isActive ? S.tabActive : {}) }}
            >
              {renamingBoardId === board.id ? (
                <input
                  autoFocus
                  value={renameDraft}
                  onChange={(e) => setRenameDraft(e.target.value)}
                  onBlur={() => submitRenameBoard(board.id)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") submitRenameBoard(board.id);
                    if (e.key === "Escape") setRenamingBoardId(null);
                  }}
                  onClick={(e) => e.stopPropagation()}
                  style={S.tabRenameInput}
                />
              ) : (
                <>
                  <span onDoubleClick={(e) => { e.stopPropagation(); setRenamingBoardId(board.id); setRenameDraft(board.name); }}>
                    {board.name}
                  </span>
                  <button
                    type="button"
                    title="Rename"
                    onClick={(e) => { e.stopPropagation(); setRenamingBoardId(board.id); setRenameDraft(board.name); }}
                    style={S.tabIconButton}
                  >
                    <svg width="11" height="11" viewBox="0 0 20 20" fill="currentColor"><path d="M13.586 3.586a2 2 0 112.828 2.828l-.793.793-2.828-2.828.793-.793zM11.379 5.793L3 14.172V17h2.828l8.38-8.379-2.83-2.828z" /></svg>
                  </button>
                </>
              )}
              {confirmDeleteBoardId === board.id ? (
                <span style={S.tabConfirm} onClick={(e) => e.stopPropagation()}>
                  <button onClick={() => submitDeleteBoard(board.id)} style={S.tabConfirmDelete}>Delete</button>
                  <button onClick={() => setConfirmDeleteBoardId(null)} style={S.tabConfirmCancel}>Cancel</button>
                </span>
              ) : (
                <span onClick={(e) => { e.stopPropagation(); setConfirmDeleteBoardId(board.id); }} style={S.tabClose}>×</span>
              )}
            </div>
          );
        })}
        <button onClick={() => { setAddBoardOpen(true); setAddBoardName(""); }} style={S.addTabButton}>+ Add board</button>
      </div>

      <div style={S.toolbar}>
        <div>
          <h2 style={S.heading}>{boards.find((b) => b.id === activeBoardId)?.name ?? "Vision Board"}</h2>
          <div style={S.meta}>{items.length} item{items.length !== 1 ? "s" : ""}</div>
        </div>
        <button
          type="button"
          disabled={!activeBoardId}
          onClick={() => { setAddItemName(""); setAddItemOpen(true); }}
          style={{ ...S.primary, opacity: activeBoardId ? 1 : 0.4, cursor: activeBoardId ? "pointer" : "not-allowed" }}
        >
          Add Item
        </button>
      </div>

      <div style={S.grid}>
        {items.map((item) => (
          <VisionCard
            key={item.id}
            item={item}
            isDragging={dragItemId === item.id}
            isDragOver={dragOverItemId === item.id && dragItemId !== item.id}
            onOpen={() => setSelectedItemId(item.id)}
            onDragStart={() => setDragItemId(item.id)}
            onDragOver={(e) => {
              if (!dragItemId) return;
              e.preventDefault();
              setDragOverItemId(item.id);
            }}
            onDragLeave={() => setDragOverItemId((c) => c === item.id ? null : c)}
            onDrop={(e) => {
              e.preventDefault();
              reorder(item.id);
              setDragItemId(null);
              setDragOverItemId(null);
            }}
            onDragEnd={() => { setDragItemId(null); setDragOverItemId(null); }}
            onDelete={() => handleDeleteItem(item.id)}
          />
        ))}
        {activeBoardId && <EmptyDropCard onAddImage={handleDropImageOnEmptyCard} />}
        {!activeBoardId && (
          <div style={S.emptyState}>Add a board using the tab bar above.</div>
        )}
      </div>

      {selectedItem && typeof document !== "undefined" && createPortal(
        <ItemDrawer itemList={selectedItem} onClose={() => setSelectedItemId(null)} />,
        document.body,
      )}

      {addBoardOpen && typeof document !== "undefined" && createPortal(
        <Modal title="Add board" placeholder="Board name" value={addBoardName} setValue={setAddBoardName} onSubmit={submitAddBoard} onCancel={() => setAddBoardOpen(false)} />,
        document.body,
      )}

      {addItemOpen && typeof document !== "undefined" && createPortal(
        <Modal title="Add item" placeholder="Item name" value={addItemName} setValue={setAddItemName} onSubmit={submitAddItem} onCancel={() => setAddItemOpen(false)} />,
        document.body,
      )}
    </div>
  );
}

// ─── Cards ───────────────────────────────────────────────────────────────────

const VisionCard = memo(function VisionCard({
  item,
  isDragging,
  isDragOver,
  onOpen,
  onDragStart,
  onDragOver,
  onDragLeave,
  onDrop,
  onDragEnd,
  onDelete,
}: {
  item: VisionItemListItem;
  isDragging: boolean;
  isDragOver: boolean;
  onOpen: () => void;
  onDragStart: (e: React.DragEvent<HTMLDivElement>) => void;
  onDragOver: (e: React.DragEvent<HTMLDivElement>) => void;
  onDragLeave: () => void;
  onDrop: (e: React.DragEvent<HTMLDivElement>) => void;
  onDragEnd: () => void;
  onDelete: () => void;
}) {
  const [hover, setHover] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const updatedAtVersion = typeof item.updatedAt === "string"
    ? new Date(item.updatedAt).getTime()
    : item.updatedAt.getTime();

  return (
    <div
      style={{
        ...S.card,
        ...(isDragging ? S.cardDragging : {}),
        ...(isDragOver ? S.cardDropTarget : {}),
      }}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      onClick={onOpen}
      onDragOver={onDragOver}
      onDragLeave={onDragLeave}
      onDrop={onDrop}
    >
      <span
        draggable
        style={{ ...S.dragHandle, opacity: hover ? 0.6 : 0.25 }}
        title="Drag to reorder"
        onDragStart={onDragStart}
        onDragEnd={onDragEnd}
        onClick={(e) => e.stopPropagation()}
      >::</span>

      {(hover || confirmDelete) && (
        <button
          type="button"
          title="Delete item"
          onClick={(e) => { e.stopPropagation(); setConfirmDelete(true); }}
          style={S.cardDelete}
        >×</button>
      )}
      {confirmDelete && (
        <div style={S.confirmOverlay} onClick={(e) => e.stopPropagation()}>
          <div style={S.confirmText}>Delete this item?</div>
          <div style={{ display: "flex", gap: 8 }}>
            <button type="button" onClick={(e) => { e.stopPropagation(); onDelete(); }} style={S.confirmDelete}>Delete</button>
            <button type="button" onClick={(e) => { e.stopPropagation(); setConfirmDelete(false); }} style={S.confirmCancel}>Cancel</button>
          </div>
        </div>
      )}

      <div style={S.cardImageWrap}>
        {item.imageCount > 0 ? (
          <img
            src={`${CARD_THUMB_BASE}/${item.id}?v=${updatedAtVersion}`}
            alt={item.name}
            style={S.cardImage}
            loading="lazy"
            decoding="async"
          />
        ) : (
          <div style={S.cardImageEmpty}>No image yet</div>
        )}
      </div>

      <div style={S.cardBody}>
        <span style={S.cardTitle}>{item.name || "Untitled"}</span>
      </div>
    </div>
  );
});

function EmptyDropCard({ onAddImage }: { onAddImage: (file: File) => void }) {
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  return (
    <div
      style={{ ...S.card, ...S.emptyDropCard, ...(dragOver ? S.emptyDropCardActive : {}) }}
      onClick={() => inputRef.current?.click()}
      onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
      onDragEnter={(e) => { e.preventDefault(); setDragOver(true); }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        const file = Array.from(e.dataTransfer.files).find((f) => f.type.startsWith("image/"));
        if (file) onAddImage(file);
      }}
    >
      <div style={S.cardImageWrap}>
        <div style={S.cardImageEmpty}>{dragOver ? "Drop image" : "+ Drop or click"}</div>
      </div>
      <input
        ref={inputRef}
        type="file"
        accept="image/*"
        style={{ display: "none" }}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onAddImage(f);
          e.target.value = "";
        }}
      />
    </div>
  );
}

// ─── Drawer ──────────────────────────────────────────────────────────────────

function ItemDrawer({
  itemList,
  onClose,
}: {
  itemList: VisionItemListItem;
  onClose: () => void;
}) {
  const fetcher = useFetcher();
  const updateFetcher = useFetcher();
  const [loaded, setLoaded] = useState(false);
  const [name, setName] = useState(itemList.name);
  const [notes, setNotes] = useState("");
  const [fields, setFields] = useState<VisionField[]>([]);
  const [savedCount, setSavedCount] = useState(itemList.imageCount);
  const [version, setVersion] = useState(() =>
    typeof itemList.updatedAt === "string"
      ? new Date(itemList.updatedAt).getTime()
      : itemList.updatedAt.getTime(),
  );
  const [pending, setPending] = useState<Array<{ id: string; blobUrl: string; dataUrl: string }>>([]);
  const [lightboxIndex, setLightboxIndex] = useState<number | null>(null);
  const [imageDragOver, setImageDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Fetch full item once on open. Slim list payload doesn't carry fields/notes.
  useEffect(() => {
    fetcher.submit({ intent: "vb_get_item", itemId: String(itemList.id) }, { method: "post" });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [itemList.id]);

  useEffect(() => {
    const item = (fetcher.data as { item?: { name?: string; notes?: string | null; fields?: unknown; images?: unknown[]; updatedAt?: string | Date } | null } | undefined)?.item;
    if (item && !loaded) {
      setName(item.name ?? itemList.name);
      setNotes(item.notes ?? "");
      setFields(fieldsFromUnknown(item.fields));
      const len = Array.isArray(item.images) ? item.images.length : 0;
      setSavedCount(len);
      if (item.updatedAt) {
        const t = typeof item.updatedAt === "string" ? new Date(item.updatedAt).getTime() : item.updatedAt.getTime();
        setVersion(t);
      }
      setLoaded(true);
    }
  }, [fetcher.data, itemList.name, loaded]);

  useEffect(() => () => {
    setPending((cur) => {
      cur.forEach((p) => URL.revokeObjectURL(p.blobUrl));
      return [];
    });
  }, []);

  const saveName = (next: string) => {
    if (next === itemList.name) return;
    updateFetcher.submit({ intent: "vb_update_item", itemId: String(itemList.id), name: next }, { method: "post" });
  };
  const saveNotes = (next: string) => {
    updateFetcher.submit({ intent: "vb_update_item", itemId: String(itemList.id), notes: next }, { method: "post" });
  };
  const saveFields = (next: VisionField[]) => {
    updateFetcher.submit({ intent: "vb_update_item", itemId: String(itemList.id), fields: JSON.stringify(next) }, { method: "post" });
  };

  const handleAddImage = async (file: File) => {
    const localId = `p_${Math.random().toString(36).slice(2, 9)}`;
    const blobUrl = URL.createObjectURL(file);
    setPending((cur) => [...cur, { id: localId, blobUrl, dataUrl: "" }]);
    const dataUrl = await compressUpload(file);
    let thumb: string | null = null;
    if (savedCount === 0 && pending.length === 0) thumb = await makeThumb(dataUrl);
    const payload: Record<string, string> = {
      intent: "vb_append_item_image",
      itemId: String(itemList.id),
      image: dataUrl,
    };
    if (thumb) payload.thumbnail = thumb;
    fetcher.submit(payload, { method: "post" });
    setPending((cur) => cur.filter((p) => p.id !== localId));
    URL.revokeObjectURL(blobUrl);
    setSavedCount((c) => c + 1);
    setVersion(Date.now());
  };

  const handleRemoveImage = (idx: number) => {
    if (!window.confirm("Remove this image?")) return;
    updateFetcher.submit(
      { intent: "vb_remove_item_image", itemId: String(itemList.id), index: String(idx) },
      { method: "post" },
    );
    setSavedCount((c) => Math.max(0, c - 1));
    setVersion(Date.now());
  };

  const allImageUrls = useMemo(() => [
    ...Array.from({ length: savedCount }, (_, i) => `${ITEM_IMAGE_BASE}/${itemList.id}/${i}?v=${version}`),
    ...pending.map((p) => p.blobUrl),
  ], [savedCount, version, pending, itemList.id]);

  return (
    <div style={S.drawerScrim} onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div style={S.drawer} onClick={(e) => e.stopPropagation()}>
        <div style={S.drawerHeader}>
          <input
            value={name}
            onChange={(e) => setName(e.target.value)}
            onBlur={() => saveName(name)}
            style={S.drawerNameInput}
            placeholder="Untitled"
          />
          <button type="button" onClick={onClose} style={S.drawerClose}>×</button>
        </div>

        <div style={S.drawerBody}>
          <section style={S.drawerSection}>
            <h3 style={S.drawerSectionTitle}>Images</h3>
            <div
              style={{ ...S.imageDropZone, ...(imageDragOver ? S.imageDropZoneActive : {}) }}
              onDragOver={(e) => { e.preventDefault(); setImageDragOver(true); }}
              onDragLeave={() => setImageDragOver(false)}
              onDrop={(e) => {
                e.preventDefault();
                setImageDragOver(false);
                Array.from(e.dataTransfer.files)
                  .filter((f) => f.type.startsWith("image/"))
                  .forEach((f) => void handleAddImage(f));
              }}
              onPaste={(e) => {
                Array.from(e.clipboardData.files)
                  .filter((f) => f.type.startsWith("image/"))
                  .forEach((f) => void handleAddImage(f));
              }}
              tabIndex={0}
            >
              <span>Drop, paste, or click to add image</span>
              <button type="button" onClick={() => fileInputRef.current?.click()} style={S.drawerSmallButton}>Browse</button>
              <input
                ref={fileInputRef}
                type="file"
                accept="image/*"
                multiple
                style={{ display: "none" }}
                onChange={(e) => {
                  Array.from(e.target.files ?? []).forEach((f) => void handleAddImage(f));
                  e.target.value = "";
                }}
              />
            </div>

            <div style={S.imageGrid}>
              {allImageUrls.map((url, idx) => {
                const isPending = idx >= savedCount;
                return (
                  <div key={url} style={S.imageTile} onClick={() => !isPending && setLightboxIndex(idx)}>
                    <img src={url} alt="" style={S.imageTileImg} loading="lazy" decoding="async" />
                    {isPending && <div style={S.pendingBadge}>Uploading…</div>}
                    {!isPending && (
                      <button
                        type="button"
                        title="Remove"
                        onClick={(e) => { e.stopPropagation(); void handleRemoveImage(idx); }}
                        style={S.imageTileRemove}
                      >×</button>
                    )}
                  </div>
                );
              })}
            </div>
          </section>

          <section style={S.drawerSection}>
            <h3 style={S.drawerSectionTitle}>Fields</h3>
            {fields.map((f, i) => (
              <div key={f.id} style={S.fieldRow}>
                <textarea
                  value={f.text}
                  rows={2}
                  onChange={(e) => setFields((c) => c.map((x, j) => j === i ? { ...x, text: e.target.value } : x))}
                  onBlur={() => saveFields(fields)}
                  style={S.fieldInput}
                />
                <button type="button" onClick={() => { const next = fields.filter((_, j) => j !== i); setFields(next); saveFields(next); }} style={S.fieldRemove}>×</button>
              </div>
            ))}
            <button type="button" onClick={() => { const next = [...fields, { id: `f_${Math.random().toString(36).slice(2, 9)}`, text: "" }]; setFields(next); saveFields(next); }} style={S.drawerSmallButton}>+ Add field</button>
          </section>

          <section style={S.drawerSection}>
            <h3 style={S.drawerSectionTitle}>Notes</h3>
            <textarea
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
              onBlur={() => saveNotes(notes)}
              rows={4}
              style={S.notesInput}
              placeholder="Free-form notes…"
            />
          </section>
        </div>

        {lightboxIndex !== null && createPortal(
          <div style={S.lightboxScrim} onClick={() => setLightboxIndex(null)}>
            <img src={allImageUrls[lightboxIndex]} alt="" style={S.lightboxImg} />
          </div>,
          document.body,
        )}
      </div>
    </div>
  );
}

// ─── Modal helper ────────────────────────────────────────────────────────────

function Modal({
  title,
  placeholder,
  value,
  setValue,
  onSubmit,
  onCancel,
}: {
  title: string;
  placeholder: string;
  value: string;
  setValue: (v: string) => void;
  onSubmit: () => void;
  onCancel: () => void;
}) {
  return (
    <div
      style={S.modalScrim}
      onClick={(e) => { if (e.target === e.currentTarget) onCancel(); }}
    >
      <div style={S.modal}>
        <div style={S.modalTitle}>{title}</div>
        <input
          autoFocus
          style={S.modalInput}
          placeholder={placeholder}
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") onSubmit();
            if (e.key === "Escape") onCancel();
          }}
        />
        <div style={S.modalButtons}>
          <button type="button" onClick={onCancel} style={S.modalCancel}>Cancel</button>
          <button type="button" onClick={onSubmit} style={S.modalConfirm}>Add</button>
        </div>
      </div>
    </div>
  );
}

// ─── Styles ──────────────────────────────────────────────────────────────────

const S: Record<string, React.CSSProperties> = {
  page: { display: "flex", flexDirection: "column", height: "100%", minHeight: 0, gap: 12 },
  tabBar: { display: "flex", alignItems: "center", gap: 4, padding: "0 4px", borderBottom: "1px solid #e5e7eb", flexWrap: "wrap" },
  tab: { display: "flex", alignItems: "center", gap: 4, background: "#f3f4f6", color: "#374151", borderRadius: "6px 6px 0 0", padding: "6px 12px", cursor: "pointer", fontSize: 13, fontWeight: 500, border: "1px solid #e5e7eb", position: "relative", top: 1 },
  tabActive: { background: "var(--portal-primary-button-bg, #111827)", color: "var(--portal-primary-button-color, #fff)", borderColor: "var(--portal-primary-button-bg, #111827)" },
  tabRenameInput: { background: "transparent", border: "none", outline: "none", color: "inherit", fontSize: 13, fontWeight: 500, minWidth: 60 },
  tabIconButton: { background: "none", border: "none", padding: "2px 4px", cursor: "pointer", color: "inherit", opacity: 0.6, lineHeight: 1, borderRadius: 4, display: "flex", alignItems: "center" },
  tabConfirm: { display: "flex", gap: 3, marginLeft: 4 },
  tabConfirmDelete: { background: "#ef4444", color: "#fff", border: "none", borderRadius: 3, padding: "1px 6px", fontSize: 11, cursor: "pointer" },
  tabConfirmCancel: { background: "#e5e7eb", color: "#374151", border: "none", borderRadius: 3, padding: "1px 6px", fontSize: 11, cursor: "pointer" },
  tabClose: { marginLeft: 4, opacity: 0.6, fontSize: 13, cursor: "pointer", lineHeight: 1 },
  addTabButton: { background: "none", border: "1px dashed #d1d5db", borderRadius: "6px 6px 0 0", padding: "6px 12px", fontSize: 13, cursor: "pointer", color: "#6b7280", position: "relative", top: 1 },

  toolbar: { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 16px", background: "#fff", border: "1px solid #dbe3ee", borderRadius: 10 },
  heading: { margin: 0, fontSize: 20, color: "#111827", fontWeight: 700 },
  meta: { fontSize: 12, color: "#6b7280", marginTop: 2 },
  primary: { padding: "9px 18px", borderRadius: 7, border: "none", background: "var(--portal-primary-button-bg, #111827)", color: "var(--portal-primary-button-color, #fff)", fontSize: 13, fontWeight: 700, cursor: "pointer" },

  grid: { display: "grid", gridTemplateColumns: "repeat(8, minmax(0, 1fr))", gap: 12, padding: "4px 0 14px", contentVisibility: "auto" as React.CSSProperties["contentVisibility"], containIntrinsicSize: "auto 280px" as unknown as React.CSSProperties["containIntrinsicSize"] },
  card: { position: "relative", background: "#fff", border: "1px solid #e5e7eb", borderRadius: 8, overflow: "hidden", cursor: "pointer", transition: "box-shadow 0.15s, transform 0.15s" },
  cardDragging: { opacity: 0.5 },
  cardDropTarget: { boxShadow: "0 0 0 2px #2563eb" },
  cardImageWrap: { width: "100%", aspectRatio: "1.3 / 1.8", background: "#f8fafc", overflow: "hidden", position: "relative" },
  cardImage: { width: "100%", height: "100%", objectFit: "cover", display: "block" },
  cardImageEmpty: { width: "100%", height: "100%", display: "flex", alignItems: "center", justifyContent: "center", color: "#9ca3af", fontSize: 12 },
  cardBody: { padding: "8px 10px 12px" },
  cardTitle: { fontSize: 13, fontWeight: 600, color: "#111827", display: "block", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  cardDelete: { position: "absolute", top: 8, right: 8, width: 26, height: 26, borderRadius: "50%", border: "none", background: "#ef4444", color: "#fff", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, fontWeight: 700, lineHeight: 1, padding: 0, zIndex: 2, boxShadow: "0 2px 6px rgba(0,0,0,0.18)" },
  dragHandle: { position: "absolute", top: 4, left: 6, fontSize: 16, color: "#374151", cursor: "grab", transition: "opacity 0.15s", zIndex: 2, userSelect: "none" },
  confirmOverlay: { position: "absolute", inset: 0, background: "rgba(0,0,0,0.6)", zIndex: 10, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 10, borderRadius: 8 },
  confirmText: { color: "#fff", fontSize: 13, fontWeight: 600 },
  confirmDelete: { padding: "7px 16px", borderRadius: 7, border: "none", background: "#ef4444", color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" },
  confirmCancel: { padding: "7px 16px", borderRadius: 7, border: "1px solid rgba(255,255,255,0.3)", background: "transparent", color: "#fff", fontSize: 13, cursor: "pointer" },

  emptyDropCard: { borderStyle: "dashed", background: "#f8fafc" },
  emptyDropCardActive: { background: "#e0f2fe", borderColor: "#2563eb" },

  emptyState: { gridColumn: "1 / -1", padding: "48px 0", textAlign: "center", color: "#9ca3af", fontSize: 14 },

  drawerScrim: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 1500, display: "flex", justifyContent: "flex-end" },
  drawer: { width: "min(720px, 100vw)", height: "100%", background: "#fff", display: "flex", flexDirection: "column", overflow: "hidden" },
  drawerHeader: { padding: "14px 18px", borderBottom: "1px solid #e5e7eb", display: "flex", alignItems: "center", gap: 10 },
  drawerNameInput: { flex: 1, fontSize: 18, fontWeight: 700, color: "#111827", border: "1px solid transparent", outline: "none", padding: "6px 8px", borderRadius: 6, background: "#fff" },
  drawerClose: { background: "transparent", border: "none", fontSize: 24, color: "#9ca3af", cursor: "pointer", lineHeight: 1, padding: "0 8px" },
  drawerBody: { padding: "14px 18px", overflow: "auto", flex: 1, display: "flex", flexDirection: "column", gap: 18 },
  drawerSection: {},
  drawerSectionTitle: { fontSize: 13, fontWeight: 700, color: "#6b7280", textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 8, marginTop: 0 },
  drawerSmallButton: { background: "#f3f4f6", border: "1px solid #d1d5db", padding: "6px 10px", borderRadius: 6, fontSize: 12, fontWeight: 600, color: "#374151", cursor: "pointer" },

  imageDropZone: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "10px 14px", border: "1px dashed #cbd5e1", borderRadius: 8, color: "#6b7280", fontSize: 13, cursor: "pointer", marginBottom: 10 },
  imageDropZoneActive: { background: "#eff6ff", borderColor: "#2563eb", color: "#2563eb" },
  imageGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(120px, 1fr))", gap: 8 },
  imageTile: { position: "relative", aspectRatio: "1 / 1.3", border: "1px solid #e5e7eb", borderRadius: 6, overflow: "hidden", background: "#f8fafc", cursor: "zoom-in" },
  imageTileImg: { width: "100%", height: "100%", objectFit: "cover", display: "block" },
  imageTileRemove: { position: "absolute", top: 4, right: 4, width: 22, height: 22, borderRadius: "50%", border: "none", background: "rgba(0,0,0,0.6)", color: "#fff", cursor: "pointer", fontSize: 14, lineHeight: 1, padding: 0 },
  pendingBadge: { position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", background: "rgba(255,255,255,0.7)", color: "#374151", fontSize: 11, fontWeight: 700 },

  fieldRow: { display: "flex", alignItems: "flex-start", gap: 6, marginBottom: 6 },
  fieldInput: { flex: 1, fontSize: 13, padding: "6px 8px", border: "1px solid #d1d5db", borderRadius: 6, fontFamily: "inherit", resize: "vertical", outline: "none", boxSizing: "border-box" },
  fieldRemove: { background: "transparent", border: "1px solid #fecaca", color: "#ef4444", borderRadius: 6, width: 26, height: 26, cursor: "pointer", lineHeight: 1, padding: 0, fontSize: 16 },

  notesInput: { width: "100%", fontSize: 13, padding: "8px 10px", border: "1px solid #d1d5db", borderRadius: 6, fontFamily: "inherit", resize: "vertical", outline: "none", boxSizing: "border-box" },

  lightboxScrim: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.85)", zIndex: 1600, display: "flex", alignItems: "center", justifyContent: "center", cursor: "zoom-out" },
  lightboxImg: { maxWidth: "92vw", maxHeight: "92vh", objectFit: "contain", borderRadius: 6 },

  modalScrim: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.35)", zIndex: 1700, display: "flex", alignItems: "center", justifyContent: "center" },
  modal: { background: "#fff", borderRadius: 12, padding: "28px 28px 24px", width: 380, boxShadow: "0 20px 60px rgba(0,0,0,0.18)" },
  modalTitle: { fontSize: 16, fontWeight: 700, color: "#111827", marginBottom: 16 },
  modalInput: { width: "100%", border: "1px solid #d1d5db", borderRadius: 7, padding: "9px 12px", fontSize: 14, outline: "none", boxSizing: "border-box", fontFamily: "inherit" },
  modalButtons: { display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 },
  modalCancel: { padding: "8px 18px", borderRadius: 7, border: "1px solid #d1d5db", background: "#fff", fontSize: 14, cursor: "pointer", color: "#374151" },
  modalConfirm: { padding: "8px 18px", borderRadius: 7, border: "none", background: "var(--portal-primary-button-bg, #111827)", color: "var(--portal-primary-button-color, #fff)", fontSize: 14, fontWeight: 600, cursor: "pointer" },
};
