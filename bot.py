#!/usr/bin/env python3
"""
📚 EduSave Bot — O'quv materiallarini kategoriyalarda saqlash boti
  + 🗜 Fayl siqish (rasm va hujjatlar)
  + 📦 Kategoriyani ZIP arxiv qilish
"""

import logging
import sqlite3
import os
import io
import zipfile
from datetime import datetime

from PIL import Image

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, filters, ContextTypes
)

logging.basicConfig(
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

DB_PATH        = "edusave.db"
ITEMS_PER_PAGE = 5

# ─────────────────────────── DATABASE ────────────────────────────

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    with get_conn() as conn:
        conn.executescript('''
            CREATE TABLE IF NOT EXISTS categories (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id    INTEGER NOT NULL,
                name       TEXT    NOT NULL,
                emoji      TEXT    DEFAULT '📁',
                created_at TEXT    DEFAULT (datetime('now')),
                UNIQUE(user_id, name)
            );
            CREATE TABLE IF NOT EXISTS items (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id      INTEGER NOT NULL,
                category_id  INTEGER NOT NULL,
                msg_type     TEXT    NOT NULL,
                text_content TEXT,
                file_id      TEXT,
                file_name    TEXT,
                caption      TEXT,
                tags         TEXT    DEFAULT '',
                is_favorite  INTEGER DEFAULT 0,
                created_at   TEXT    DEFAULT (datetime('now')),
                FOREIGN KEY (category_id) REFERENCES categories(id) ON DELETE CASCADE
            );
        ''')

# ─────────────────────────── DB HELPERS ──────────────────────────

def get_categories(user_id):
    with get_conn() as conn:
        return conn.execute(
            "SELECT id, name, emoji FROM categories WHERE user_id=? ORDER BY name",
            (user_id,)
        ).fetchall()

def add_category(user_id, name, emoji='📁'):
    try:
        with get_conn() as conn:
            conn.execute(
                "INSERT INTO categories (user_id, name, emoji) VALUES (?,?,?)",
                (user_id, name, emoji)
            )
        return True
    except sqlite3.IntegrityError:
        return False

def delete_category(user_id, cat_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM categories WHERE id=? AND user_id=?", (cat_id, user_id))

def get_category(user_id, cat_id):
    with get_conn() as conn:
        return conn.execute(
            "SELECT id, name, emoji FROM categories WHERE id=? AND user_id=?",
            (cat_id, user_id)
        ).fetchone()

def save_item(user_id, cat_id, msg_type, text=None, file_id=None,
              file_name=None, caption=None, tags=''):
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO items "
            "(user_id,category_id,msg_type,text_content,file_id,file_name,caption,tags)"
            " VALUES (?,?,?,?,?,?,?,?)",
            (user_id, cat_id, msg_type, text, file_id, file_name, caption, tags)
        )

def get_items(user_id, cat_id, page=0, search=None):
    offset = page * ITEMS_PER_PAGE
    with get_conn() as conn:
        if search:
            q = f'%{search}%'
            rows = conn.execute(
                "SELECT id,msg_type,text_content,file_name,caption,tags,is_favorite,created_at"
                " FROM items WHERE user_id=? AND category_id=?"
                " AND (text_content LIKE ? OR caption LIKE ? OR tags LIKE ? OR file_name LIKE ?)"
                " ORDER BY is_favorite DESC, created_at DESC LIMIT ? OFFSET ?",
                (user_id, cat_id, q, q, q, q, ITEMS_PER_PAGE, offset)
            ).fetchall()
            total = conn.execute(
                "SELECT COUNT(*) FROM items WHERE user_id=? AND category_id=?"
                " AND (text_content LIKE ? OR caption LIKE ? OR tags LIKE ? OR file_name LIKE ?)",
                (user_id, cat_id, q, q, q, q)
            ).fetchone()[0]
        else:
            rows = conn.execute(
                "SELECT id,msg_type,text_content,file_name,caption,tags,is_favorite,created_at"
                " FROM items WHERE user_id=? AND category_id=?"
                " ORDER BY is_favorite DESC, created_at DESC LIMIT ? OFFSET ?",
                (user_id, cat_id, ITEMS_PER_PAGE, offset)
            ).fetchall()
            total = conn.execute(
                "SELECT COUNT(*) FROM items WHERE user_id=? AND category_id=?",
                (user_id, cat_id)
            ).fetchone()[0]
    return rows, total

def get_item(user_id, item_id):
    with get_conn() as conn:
        return conn.execute(
            "SELECT id,msg_type,text_content,file_id,file_name,caption,tags,is_favorite"
            " FROM items WHERE id=? AND user_id=?",
            (item_id, user_id)
        ).fetchone()

def get_all_file_items(user_id, cat_id):
    with get_conn() as conn:
        return conn.execute(
            "SELECT id,msg_type,file_id,file_name,caption"
            " FROM items WHERE user_id=? AND category_id=? AND file_id IS NOT NULL"
            " ORDER BY created_at",
            (user_id, cat_id)
        ).fetchall()

def toggle_favorite(user_id, item_id):
    with get_conn() as conn:
        conn.execute(
            "UPDATE items SET is_favorite = 1 - is_favorite WHERE id=? AND user_id=?",
            (item_id, user_id)
        )

def delete_item(user_id, item_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM items WHERE id=? AND user_id=?", (item_id, user_id))

def search_all_items(user_id, query):
    q = f'%{query}%'
    with get_conn() as conn:
        return conn.execute(
            "SELECT i.id,i.msg_type,i.text_content,i.file_name,i.caption,"
            "       i.is_favorite,i.created_at,c.name,c.emoji,i.category_id"
            " FROM items i JOIN categories c ON i.category_id=c.id"
            " WHERE i.user_id=?"
            " AND (i.text_content LIKE ? OR i.caption LIKE ? OR i.tags LIKE ? OR i.file_name LIKE ?)"
            " ORDER BY i.is_favorite DESC, i.created_at DESC LIMIT 15",
            (user_id, q, q, q, q)
        ).fetchall()

def get_favorites(user_id):
    with get_conn() as conn:
        return conn.execute(
            "SELECT i.id,i.msg_type,i.text_content,i.file_name,i.caption,"
            "       i.created_at,c.name,c.emoji,i.category_id"
            " FROM items i JOIN categories c ON i.category_id=c.id"
            " WHERE i.user_id=? AND i.is_favorite=1"
            " ORDER BY i.created_at DESC LIMIT 20",
            (user_id,)
        ).fetchall()

def get_stats(user_id):
    with get_conn() as conn:
        total_cats  = conn.execute("SELECT COUNT(*) FROM categories WHERE user_id=?", (user_id,)).fetchone()[0]
        total_items = conn.execute("SELECT COUNT(*) FROM items WHERE user_id=?", (user_id,)).fetchone()[0]
        favorites   = conn.execute("SELECT COUNT(*) FROM items WHERE user_id=? AND is_favorite=1", (user_id,)).fetchone()[0]
        top_cats    = conn.execute(
            "SELECT c.name,c.emoji,COUNT(i.id) cnt FROM categories c"
            " LEFT JOIN items i ON c.id=i.category_id WHERE c.user_id=?"
            " GROUP BY c.id ORDER BY cnt DESC LIMIT 5",
            (user_id,)
        ).fetchall()
    return total_cats, total_items, favorites, top_cats

def cat_item_count(user_id, cat_id):
    with get_conn() as conn:
        return conn.execute(
            "SELECT COUNT(*) FROM items WHERE user_id=? AND category_id=?",
            (user_id, cat_id)
        ).fetchone()[0]

# ─────────────────────────── ICONS ───────────────────────────────

TYPE_ICON = {
    "text": "📝", "photo": "🖼", "document": "📄",
    "video": "🎬", "audio": "🎵", "voice": "🎤", "sticker": "😊"
}
TYPE_NAME = {
    "text": "Matn", "photo": "Rasm", "document": "Hujjat",
    "video": "Video", "audio": "Audio", "voice": "Ovozli xabar"
}
COMPRESSIBLE = {"photo", "document", "audio", "video", "voice"}

def item_preview(row):
    _, msg_type, text, fname, cap, *_ = row
    raw = text or cap or fname or "📎 Fayl"
    return TYPE_ICON.get(msg_type, "📎") + " " + (raw[:35] + "…" if len(raw) > 35 else raw)

def human_size(nbytes: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if nbytes < 1024:
            return f"{nbytes:.1f} {unit}"
        nbytes /= 1024
    return f"{nbytes:.1f} TB"

# ─────────────────────────── COMPRESSION ─────────────────────────

QUALITY_LEVELS = {
    "yuqori":  85,   # engil siqish  ~30% kichrayadi
    "urtacha": 60,   # o'rtacha      ~55% kichrayadi
    "kuchli":  35,   # kuchli siqish ~75% kichrayadi
}

async def compress_image_bytes(data: bytes, quality: int) -> bytes:
    img = Image.open(io.BytesIO(data))
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=quality, optimize=True)
    return out.getvalue()

async def compress_generic_bytes(data: bytes, filename: str) -> tuple:
    out = io.BytesIO()
    zname = (filename or "fayl") + ".zip"
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        zf.writestr(filename or "fayl", data)
    return out.getvalue(), zname

async def download_file(bot, file_id: str) -> bytes:
    tg_file = await bot.get_file(file_id)
    buf = io.BytesIO()
    await tg_file.download_to_memory(buf)
    return buf.getvalue()

async def build_zip_for_category(bot, user_id: int, cat_id: int) -> io.BytesIO | None:
    file_items = get_all_file_items(user_id, cat_id)
    if not file_items:
        return None
    zip_buf = io.BytesIO()
    counters = {}
    with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for item_id, msg_type, file_id, file_name, caption in file_items:
            try:
                data = await download_file(bot, file_id)
            except Exception as e:
                logger.warning(f"ZIP skip {file_id}: {e}")
                continue
            if file_name:
                arc_name = file_name
            else:
                ext = {"photo": "jpg", "video": "mp4", "audio": "mp3", "voice": "ogg"}.get(msg_type, "bin")
                arc_name = f"{msg_type}_{item_id}.{ext}"
            if arc_name in counters:
                counters[arc_name] += 1
                base, _, ext2 = arc_name.rpartition(".")
                arc_name = f"{base}_{counters[arc_name]}.{ext2}" if ext2 else f"{arc_name}_{counters[arc_name]}"
            else:
                counters[arc_name] = 0
            zf.writestr(arc_name, data)
    zip_buf.seek(0)
    return zip_buf

# ─────────────────────────── KEYBOARDS ───────────────────────────

def main_menu_kb():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📚 Kategoriyalar", callback_data="cat_list"),
         InlineKeyboardButton("🔍 Qidirish",      callback_data="search_start")],
        [InlineKeyboardButton("➕ Kategoriya",    callback_data="cat_add"),
         InlineKeyboardButton("⭐ Sevimlilar",    callback_data="favorites")],
        [InlineKeyboardButton("📊 Statistika",    callback_data="stats")],
    ])

def back_btn(target="main_menu"):
    return InlineKeyboardButton("🔙 Orqaga", callback_data=target)

def categories_kb(user_id):
    cats = get_categories(user_id)
    rows = []
    for cat_id, name, emoji in cats:
        cnt = cat_item_count(user_id, cat_id)
        rows.append([InlineKeyboardButton(
            f"{emoji} {name}  ({cnt})",
            callback_data=f"cat_open:{cat_id}:0"
        )])
    rows.append([back_btn("main_menu")])
    return InlineKeyboardMarkup(rows)

def items_list_kb(user_id, cat_id, page, total, items):
    buttons = []
    for row in items:
        item_id = row[0]
        star = "⭐ " if row[6] else ""
        buttons.append([InlineKeyboardButton(
            star + item_preview(row),
            callback_data=f"item_view:{item_id}:{cat_id}:{page}"
        )])
    nav = []
    total_pages = max(1, (total + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE)
    if page > 0:
        nav.append(InlineKeyboardButton("◀️", callback_data=f"cat_open:{cat_id}:{page-1}"))
    nav.append(InlineKeyboardButton(f"{page+1}/{total_pages}", callback_data="noop"))
    if (page + 1) * ITEMS_PER_PAGE < total:
        nav.append(InlineKeyboardButton("▶️", callback_data=f"cat_open:{cat_id}:{page+1}"))
    if nav:
        buttons.append(nav)
    buttons.append([
        InlineKeyboardButton("🔍 Izlash",     callback_data=f"cat_search:{cat_id}"),
        InlineKeyboardButton("📦 ZIP arxiv",  callback_data=f"zip_cat:{cat_id}"),
    ])
    buttons.append([
        InlineKeyboardButton("🗑 Kategoriyani o'chir", callback_data=f"cat_del_confirm:{cat_id}"),
        back_btn("cat_list"),
    ])
    return InlineKeyboardMarkup(buttons)

def item_actions_kb(item_id, cat_id, page, is_fav, msg_type):
    star_label = "💛 Sevimlilardan chiqar" if is_fav else "⭐ Sevimliga qo'sh"
    rows = [
        [InlineKeyboardButton(star_label,  callback_data=f"item_fav:{item_id}:{cat_id}:{page}"),
         InlineKeyboardButton("🗑 O'chir", callback_data=f"item_del:{item_id}:{cat_id}:{page}")],
    ]
    if msg_type in COMPRESSIBLE:
        rows.append([
            InlineKeyboardButton("🗜 Engil",   callback_data=f"compress:{item_id}:{cat_id}:{page}:yuqori"),
            InlineKeyboardButton("🗜 O'rtacha", callback_data=f"compress:{item_id}:{cat_id}:{page}:urtacha"),
            InlineKeyboardButton("🗜 Kuchli",   callback_data=f"compress:{item_id}:{cat_id}:{page}:kuchli"),
        ])
    rows.append([back_btn(f"cat_open:{cat_id}:{page}")])
    return InlineKeyboardMarkup(rows)

# ──────────────────────────── HANDLERS ───────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    name = update.effective_user.first_name
    text = (
        f"👋 Salom, <b>{name}</b>!\n\n"
        "📚 <b>EduSave Bot</b> — o'quv materiallaringizni tartibli saqlang!\n\n"
        "<b>Imkoniyatlar:</b>\n"
        "• Cheksiz kategoriya (📐 Matematika, 📖 Ona tili…)\n"
        "• Matn, rasm, hujjat, video, audio saqlash\n"
        "• 🗜 Faylni siqish — 3 daraja: engil / o'rtacha / kuchli\n"
        "• 📦 Kategoriyani ZIP arxivga yuklash\n"
        "• 🔍 Global va kategoriyada qidirish\n"
        "• ⭐ Sevimlilar va 📊 Statistika\n\n"
        "<b>Qanday saqlash:</b>\n"
        "Istalgan xabar yuboring → kategoriyani tanlang → tayyor! ✅"
    )
    await update.message.reply_text(text, reply_markup=main_menu_kb(), parse_mode="HTML")

async def cb_main_menu(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    ctx.user_data.clear()
    await q.edit_message_text("🏠 <b>Asosiy menyu</b>", reply_markup=main_menu_kb(), parse_mode="HTML")

async def cb_cat_list(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    cats = get_categories(uid)
    text = (
        f"📚 <b>Kategoriyalar</b> ({len(cats)} ta)"
        if cats else
        "📂 Hali kategoriya yo'q.\n\n➕ Kategoriya qo'shing!"
    )
    await q.edit_message_text(text, reply_markup=categories_kb(uid), parse_mode="HTML")

async def cb_cat_add(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    ctx.user_data['state'] = 'adding_category'
    kb = InlineKeyboardMarkup([[InlineKeyboardButton("❌ Bekor", callback_data="cat_list")]])
    await q.edit_message_text(
        "➕ <b>Yangi kategoriya</b>\n\nKategoriya nomini yuboring. Emoji ham qo'shsangiz bo'ladi:\n\n"
        "<code>📐 Matematika</code>\n<code>📖 Ona tili</code>\n"
        "<code>🧪 Kimyo</code>\n<code>🌍 Geografiya</code>\n<code>💻 Informatika</code>",
        reply_markup=kb, parse_mode="HTML"
    )

async def cb_cat_open(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    _, cat_id, page = q.data.split(":")
    cat_id, page = int(cat_id), int(page)
    cat = get_category(uid, cat_id)
    if not cat:
        await q.edit_message_text("❌ Kategoriya topilmadi", reply_markup=main_menu_kb())
        return
    _, name, emoji = cat
    items, total = get_items(uid, cat_id, page=page)
    text = (
        f"{emoji} <b>{name}</b>\n📦 Jami: {total} ta material"
        if total else
        f"{emoji} <b>{name}</b>\n\n📭 Hali material yo'q.\nXabar yuboring va shu kategoriyani tanlang!"
    )
    await q.edit_message_text(text, reply_markup=items_list_kb(uid, cat_id, page, total, items), parse_mode="HTML")

async def cb_cat_del_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    cat_id = int(q.data.split(":")[1])
    cat = get_category(uid, cat_id)
    if not cat:
        return
    _, name, emoji = cat
    cnt = cat_item_count(uid, cat_id)
    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Ha, o'chir", callback_data=f"cat_del:{cat_id}"),
        InlineKeyboardButton("❌ Bekor",      callback_data=f"cat_open:{cat_id}:0"),
    ]])
    await q.edit_message_text(
        f"⚠️ <b>{emoji} {name}</b> kategoriyasini o'chirmoqchimisiz?\n"
        f"Ichidagi <b>{cnt}</b> ta material ham o'chib ketadi!",
        reply_markup=kb, parse_mode="HTML"
    )

async def cb_cat_del(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    uid = q.from_user.id
    cat_id = int(q.data.split(":")[1])
    delete_category(uid, cat_id)
    await q.answer("🗑 Kategoriya o'chirildi", show_alert=True)
    cats = get_categories(uid)
    await q.edit_message_text(
        f"📚 <b>Kategoriyalar</b> ({len(cats)} ta)",
        reply_markup=categories_kb(uid), parse_mode="HTML"
    )

async def cb_cat_search(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    cat_id = int(q.data.split(":")[1])
    ctx.user_data['state']      = 'searching_cat'
    ctx.user_data['search_cat'] = cat_id
    kb = InlineKeyboardMarkup([[InlineKeyboardButton("❌ Bekor", callback_data=f"cat_open:{cat_id}:0")]])
    await q.edit_message_text("🔍 <b>Kategoriyada qidirish</b>\n\nSo'zni yuboring:", reply_markup=kb, parse_mode="HTML")

async def cb_search_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    ctx.user_data['state'] = 'searching_all'
    kb = InlineKeyboardMarkup([[InlineKeyboardButton("❌ Bekor", callback_data="main_menu")]])
    await q.edit_message_text(
        "🔍 <b>Global qidirish</b>\n\nBarcha kategoriyalarda qidiradi.\nSo'zni yuboring:",
        reply_markup=kb, parse_mode="HTML"
    )

async def cb_item_view(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid  = q.from_user.id
    _, item_id, cat_id, page = q.data.split(":")
    item_id, cat_id, page = int(item_id), int(cat_id), int(page)
    row = get_item(uid, item_id)
    if not row:
        await q.answer("❌ Topilmadi", show_alert=True)
        return
    _, msg_type, text, file_id, file_name, caption, tags, is_fav = row
    kb = item_actions_kb(item_id, cat_id, page, is_fav, msg_type)
    tag_line = f"\n\n🏷 <i>{tags}</i>" if tags else ""
    try:
        if msg_type == "text":
            await q.edit_message_text(f"📝 <b>Matn</b>{tag_line}\n\n{text}", reply_markup=kb, parse_mode="HTML")
        elif msg_type == "photo":
            await q.message.reply_photo(file_id, caption=f"{caption or ''}{tag_line}", reply_markup=kb, parse_mode="HTML")
        elif msg_type == "document":
            await q.message.reply_document(file_id, caption=f"📄 {file_name or ''}\n{caption or ''}{tag_line}", reply_markup=kb, parse_mode="HTML")
        elif msg_type == "video":
            await q.message.reply_video(file_id, caption=f"{caption or ''}{tag_line}", reply_markup=kb, parse_mode="HTML")
        elif msg_type == "audio":
            await q.message.reply_audio(file_id, caption=f"{caption or ''}{tag_line}", reply_markup=kb, parse_mode="HTML")
        elif msg_type == "voice":
            await q.message.reply_voice(file_id, caption=f"{caption or ''}{tag_line}", reply_markup=kb, parse_mode="HTML")
    except Exception as e:
        logger.warning(f"item_view error: {e}")
        await q.answer("❌ Faylni ko'rsatishda xato", show_alert=True)

async def cb_item_fav(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    uid = q.from_user.id
    _, item_id, cat_id, page = q.data.split(":")
    item_id, cat_id, page = int(item_id), int(cat_id), int(page)
    toggle_favorite(uid, item_id)
    row = get_item(uid, item_id)
    is_fav = row[7] if row else 0
    await q.answer("⭐ Sevimliga qo'shildi!" if is_fav else "💛 Sevimlilardan olib tashlandi", show_alert=True)
    update.callback_query.data = f"item_view:{item_id}:{cat_id}:{page}"
    await cb_item_view(update, ctx)

async def cb_item_del(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    uid = q.from_user.id
    _, item_id, cat_id, page = q.data.split(":")
    item_id, cat_id, page = int(item_id), int(cat_id), int(page)
    delete_item(uid, item_id)
    await q.answer("🗑 O'chirildi", show_alert=True)
    items, total = get_items(uid, cat_id, page=page)
    real_page = min(page, max(0, (total - 1) // ITEMS_PER_PAGE))
    items, total = get_items(uid, cat_id, page=real_page)
    cat = get_category(uid, cat_id)
    if cat:
        _, name, emoji = cat
        await q.edit_message_text(
            f"{emoji} <b>{name}</b>\n📦 Jami: {total} ta material",
            reply_markup=items_list_kb(uid, cat_id, real_page, total, items),
            parse_mode="HTML"
        )

async def cb_favorites(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    favs = get_favorites(uid)
    kb = InlineKeyboardMarkup([[back_btn("main_menu")]])
    if not favs:
        await q.edit_message_text("⭐ Hali sevimli material yo'q", reply_markup=kb)
        return
    lines = [f"⭐ <b>Sevimlilar</b> ({len(favs)} ta)\n"]
    for row in favs:
        _, msg_type, text, fname, cap, created, cat_name, cat_emoji, cat_id = row
        lines.append(f"• {item_preview(row)}\n   └ {cat_emoji} {cat_name}\n")
    await q.edit_message_text("\n".join(lines), reply_markup=kb, parse_mode="HTML")

async def cb_stats(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    total_cats, total_items, favorites, top_cats = get_stats(uid)
    lines = [
        "📊 <b>Statistika</b>\n",
        f"📁 Kategoriyalar: <b>{total_cats}</b>",
        f"📦 Jami materiallar: <b>{total_items}</b>",
        f"⭐ Sevimlilar: <b>{favorites}</b>",
    ]
    if top_cats:
        lines.append("\n🏆 <b>Top kategoriyalar:</b>")
        for i, (name, emoji, cnt) in enumerate(top_cats, 1):
            lines.append(f"  {i}. {emoji} {name} — {cnt} ta")
    kb = InlineKeyboardMarkup([[back_btn("main_menu")]])
    await q.edit_message_text("\n".join(lines), reply_markup=kb, parse_mode="HTML")

# ─────────────────── 🗜 COMPRESS HANDLER ─────────────────────────

async def cb_compress(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer("⏳ Siqilmoqda…")
    uid = q.from_user.id
    parts = q.data.split(":")
    item_id, cat_id, page, level = int(parts[1]), int(parts[2]), int(parts[3]), parts[4]

    row = get_item(uid, item_id)
    if not row:
        await q.answer("❌ Topilmadi", show_alert=True)
        return
    _, msg_type, text, file_id, file_name, caption, tags, is_fav = row

    level_label = {"yuqori": "Engil (~30%)", "urtacha": "O'rtacha (~55%)", "kuchli": "Kuchli (~75%)"}
    status_msg = await q.message.reply_text(
        f"⏳ Fayl yuklanmoqda…\n🗜 Siqish darajasi: {level_label[level]}"
    )
    try:
        original_data = await download_file(ctx.bot, file_id)
        original_size = len(original_data)

        if msg_type == "photo":
            quality = QUALITY_LEVELS[level]
            compressed = await compress_image_bytes(original_data, quality)
            compressed_size = len(compressed)
            ratio = 100 - (compressed_size / original_size * 100)
            await status_msg.delete()
            await q.message.reply_photo(
                io.BytesIO(compressed),
                caption=(
                    f"🗜 <b>Siqilgan rasm</b> [{level_label[level]}]\n"
                    f"📏 Avval: <b>{human_size(original_size)}</b>\n"
                    f"📉 Keyin: <b>{human_size(compressed_size)}</b> "
                    f"(<b>-{ratio:.0f}%</b> tejaldi)"
                ),
                parse_mode="HTML"
            )

        elif msg_type in ("document", "audio", "voice", "video"):
            fname = file_name or f"fayl_{item_id}.bin"
            zip_data, zip_name = await compress_generic_bytes(original_data, fname)
            zip_size = len(zip_data)
            ratio = 100 - (zip_size / original_size * 100) if original_size else 0
            await status_msg.delete()
            await q.message.reply_document(
                io.BytesIO(zip_data),
                filename=zip_name,
                caption=(
                    f"🗜 <b>Siqilgan fayl (ZIP)</b> [{level_label[level]}]\n"
                    f"📏 Avval: <b>{human_size(original_size)}</b>\n"
                    f"📉 Keyin: <b>{human_size(zip_size)}</b> "
                    f"(<b>-{ratio:.0f}%</b> tejaldi)"
                ),
                parse_mode="HTML"
            )
        else:
            await status_msg.edit_text("❌ Bu turdagi fayl siqishni qo'llab-quvvatlamaydi")

    except Exception as e:
        logger.error(f"compress error: {e}")
        await status_msg.edit_text(f"❌ Siqishda xato: {e}")

# ─────────────────── 📦 ZIP CATEGORY HANDLER ─────────────────────

async def cb_zip_cat(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer("⏳ ZIP yaratilmoqda…")
    uid = q.from_user.id
    cat_id = int(q.data.split(":")[1])

    cat = get_category(uid, cat_id)
    if not cat:
        await q.answer("❌ Kategoriya topilmadi", show_alert=True)
        return
    _, cat_name, cat_emoji = cat

    file_items = get_all_file_items(uid, cat_id)
    if not file_items:
        await q.answer("📭 Bu kategoriyada fayl yo'q!", show_alert=True)
        return

    status_msg = await q.message.reply_text(
        f"⏳ <b>{cat_emoji} {cat_name}</b> uchun ZIP yaratilmoqda…\n"
        f"📎 {len(file_items)} ta fayl yuklanmoqda, biroz kuting…",
        parse_mode="HTML"
    )
    try:
        zip_buf = await build_zip_for_category(ctx.bot, uid, cat_id)
        if zip_buf is None:
            await status_msg.edit_text("📭 Fayllar topilmadi")
            return
        zip_data = zip_buf.read()
        zip_size = len(zip_data)
        safe_name = "".join(c if c.isalnum() or c in (" ", "_", "-") else "_" for c in cat_name)
        zip_filename = f"{safe_name}_{datetime.now().strftime('%Y%m%d')}.zip"
        await status_msg.delete()
        await q.message.reply_document(
            io.BytesIO(zip_data),
            filename=zip_filename,
            caption=(
                f"📦 <b>{cat_emoji} {cat_name}</b> — ZIP arxiv\n"
                f"📎 {len(file_items)} ta fayl\n"
                f"📏 Hajm: <b>{human_size(zip_size)}</b>\n"
                f"📅 {datetime.now().strftime('%d.%m.%Y %H:%M')}"
            ),
            parse_mode="HTML"
        )
    except Exception as e:
        logger.error(f"zip_cat error: {e}")
        await status_msg.edit_text(f"❌ ZIP yaratishda xato: {e}")

# ─────────────────── SAVE / STATE HANDLERS ───────────────────────

async def cb_save_to(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    cat_id = int(q.data.split(":")[1])
    pending = ctx.user_data.pop("pending_msg", None)
    ctx.user_data.clear()
    if not pending:
        await q.edit_message_text("❌ Saqlash uchun xabar topilmadi", reply_markup=main_menu_kb())
        return
    save_item(uid, cat_id, **pending)
    cat = get_category(uid, cat_id)
    _, name, emoji = cat if cat else (None, "?", "📁")
    await q.edit_message_text(
        f"✅ <b>{emoji} {name}</b> ga saqlandi!",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("📂 Kategoriyani ochish", callback_data=f"cat_open:{cat_id}:0"),
            InlineKeyboardButton("🏠 Menyu",               callback_data="main_menu"),
        ]]),
        parse_mode="HTML"
    )

async def cb_save_cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer("❌ Bekor qilindi")
    ctx.user_data.clear()
    await q.edit_message_text("🏠 <b>Asosiy menyu</b>", reply_markup=main_menu_kb(), parse_mode="HTML")

async def cb_noop(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()

# ── Message handler ───────────────────────────────────────────────

async def handle_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid   = update.effective_user.id
    msg   = update.message
    state = ctx.user_data.get("state")

    if state == "adding_category":
        text = msg.text.strip() if msg.text else ""
        if not text:
            await msg.reply_text("❌ Matn yuboring")
            return
        emoji, name = "📁", text
        if text and ord(text[0]) > 127:
            parts = text.split(None, 1)
            if len(parts) == 2:
                emoji, name = parts
        if len(name) > 50:
            await msg.reply_text("❌ Nom juda uzun (max 50 belgi)")
            return
        ctx.user_data.clear()
        ok = add_category(uid, name, emoji)
        if ok:
            await msg.reply_text(f"✅ <b>{emoji} {name}</b> yaratildi!", reply_markup=main_menu_kb(), parse_mode="HTML")
        else:
            await msg.reply_text("⚠️ Bu nom allaqachon mavjud!", reply_markup=main_menu_kb())
        return

    if state in ("searching_all", "searching_cat"):
        query_text = msg.text.strip() if msg.text else ""
        if not query_text:
            await msg.reply_text("❌ Matn yuboring")
            return
        if state == "searching_cat":
            cat_id = ctx.user_data.get("search_cat")
            ctx.user_data.clear()
            items, total = get_items(uid, cat_id, page=0, search=query_text)
            cat = get_category(uid, cat_id)
            _, name, emoji = cat if cat else (None, "?", "📁")
            if not items:
                await msg.reply_text(
                    f"🔍 <b>«{query_text}»</b> bo'yicha {emoji} {name} da hech narsa topilmadi",
                    reply_markup=main_menu_kb(), parse_mode="HTML"
                )
            else:
                await msg.reply_text(
                    f"🔍 <b>«{query_text}»</b> — {total} ta natija ({emoji} {name})",
                    reply_markup=items_list_kb(uid, cat_id, 0, total, items), parse_mode="HTML"
                )
        else:
            ctx.user_data.clear()
            results = search_all_items(uid, query_text)
            if not results:
                await msg.reply_text(
                    f"🔍 <b>«{query_text}»</b> bo'yicha hech narsa topilmadi",
                    reply_markup=main_menu_kb(), parse_mode="HTML"
                )
            else:
                lines = [f"🔍 <b>«{query_text}»</b> — {len(results)} ta natija:\n"]
                for row in results:
                    _, msg_type, text, fname, cap, is_fav, created, cat_name, cat_emoji, cat_id = row
                    star = "⭐ " if is_fav else ""
                    lines.append(f"{star}{item_preview(row)}\n   └ {cat_emoji} {cat_name}\n")
                await msg.reply_text("\n".join(lines), reply_markup=main_menu_kb(), parse_mode="HTML")
        return

    # ── Yangi material saqlash ─────────────────────────────────────
    m = msg
    msg_type = file_id = file_name = caption = text_content = None

    if m.text and not m.text.startswith("/"):
        msg_type, text_content = "text", m.text
    elif m.photo:
        msg_type = "photo";    file_id = m.photo[-1].file_id;   caption = m.caption
    elif m.document:
        msg_type = "document"; file_id = m.document.file_id
        file_name = m.document.file_name;                        caption = m.caption
    elif m.video:
        msg_type = "video";    file_id = m.video.file_id;        caption = m.caption
    elif m.audio:
        msg_type = "audio";    file_id = m.audio.file_id;        caption = m.caption
    elif m.voice:
        msg_type = "voice";    file_id = m.voice.file_id;        caption = m.caption
    else:
        await m.reply_text("❓ Bu turdagi xabarni saqlab bo'lmaydi")
        return

    ctx.user_data['pending_msg'] = dict(
        msg_type=msg_type, text=text_content, file_id=file_id,
        file_name=file_name, caption=caption
    )
    ctx.user_data['state'] = 'saving'

    cats = get_categories(uid)
    if not cats:
        await m.reply_text(
            "📂 Hali kategoriya yo'q!\n\nAvval kategoriya yarating:",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("➕ Kategoriya yaratish", callback_data="cat_add")
            ]])
        )
        return

    buttons = [
        [InlineKeyboardButton(f"{emoji} {name}", callback_data=f"save_to:{cat_id}")]
        for cat_id, name, emoji in cats
    ]
    buttons.append([InlineKeyboardButton("❌ Bekor qilish", callback_data="save_cancel")])
    icon  = TYPE_ICON.get(msg_type, "📎")
    tname = TYPE_NAME.get(msg_type, "Fayl")
    await m.reply_text(
        f"{icon} <b>{tname}</b> — qaysi kategoriyaga saqlash?",
        reply_markup=InlineKeyboardMarkup(buttons), parse_mode="HTML"
    )

# ─────────────────────────── MAIN ────────────────────────────────

def main():
    token = os.getenv("BOT_TOKEN")
    if not token:
        print("❌ BOT_TOKEN muhit o'zgaruvchisi topilmadi!")
        print("   export BOT_TOKEN=your_token_here")
        return
    init_db()
    logger.info("DB initialized ✅")
    app = Application.builder().token(token).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("menu",  cmd_start))
    cb_map = {
        "^main_menu$":         cb_main_menu,
        "^cat_list$":          cb_cat_list,
        "^cat_add$":           cb_cat_add,
        "^cat_open:":          cb_cat_open,
        "^cat_del_confirm:":   cb_cat_del_confirm,
        "^cat_del:":           cb_cat_del,
        "^cat_search:":        cb_cat_search,
        "^search_start$":      cb_search_start,
        "^item_view:":         cb_item_view,
        "^item_fav:":          cb_item_fav,
        "^item_del:":          cb_item_del,
        "^favorites$":         cb_favorites,
        "^stats$":             cb_stats,
        "^save_to:":           cb_save_to,
        "^save_cancel$":       cb_save_cancel,
        "^compress:":          cb_compress,
        "^zip_cat:":           cb_zip_cat,
        "^noop$":              cb_noop,
    }
    for pattern, handler in cb_map.items():
        app.add_handler(CallbackQueryHandler(handler, pattern=pattern))
    app.add_handler(MessageHandler(
        filters.TEXT | filters.PHOTO | filters.Document.ALL |
        filters.VIDEO | filters.AUDIO | filters.VOICE,
        handle_message
    ))
    logger.info("🤖 Bot ishga tushdi!")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
