import discord
from discord.ext import commands, tasks
import asyncio
import os
import csv
import io
from datetime import datetime, timedelta, timezone
from collections import defaultdict
import sqlite3

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==================== CONFIG ====================
DISCORD_TOKEN         = os.environ.get("DISCORD_TOKEN")
APPROVAL_CHANNELS     = ["ssrp-approval", "outrider-ssrp-approval"]  # Channel submit SSRP
REPORT_CHANNEL_NAME   = "leaderboard-ssrp-command"   # Channel laporan otomatis
INACTIVE_CHANNEL_NAME = "inactive-permission"         # Channel notif tidak aktif
SUBMIT_COOLDOWN       = 30                            # Detik jeda antar submit (anti-spam)
INACTIVE_CHECK_HOUR   = 9                             # Jam berapa cek inaktif (UTC)
INACTIVE_CHECK_WEEKDAY= 0                             # 0=Senin, 6=Minggu
ADMIN_ROLE_ID         = 1358301859663839374           # Role ID yang bisa pakai command admin
LOG_CHANNEL_NAME      = "ssrp-check"                 # Channel log addpoint/removepoint
# ================================================

intents = discord.Intents.default()
intents.message_content = True
intents.members = True

bot = commands.Bot(command_prefix="!", intents=intents)
last_submit: dict[str, datetime] = defaultdict(lambda: datetime.min.replace(tzinfo=timezone.utc))
pending_removals: dict[str, dict] = {}

# ─── Database ─────────────────────────────────────────────────────────────────
DB = "ssrp_points.db"

def init_db():
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS points (
        user_id TEXT PRIMARY KEY, username TEXT,
        total_points INTEGER DEFAULT 0, week_points INTEGER DEFAULT 0,
        last_reset TEXT, last_submit TEXT
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS submissions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id TEXT, username TEXT, submitted_at TEXT,
        photo_count INTEGER, is_valid INTEGER, reason TEXT
    )""")
    conn.commit(); conn.close()

def upsert_user(user_id, username):
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("SELECT user_id FROM points WHERE user_id=?", (user_id,))
    if not c.fetchone():
        c.execute("INSERT INTO points (user_id,username,total_points,week_points,last_reset,last_submit) VALUES (?,?,0,0,?,NULL)",
                  (user_id, username, datetime.now(timezone.utc).isoformat()))
    conn.commit(); conn.close()

def add_point(user_id, username):
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    now = datetime.now(timezone.utc)
    c.execute("SELECT week_points, last_reset FROM points WHERE user_id=?", (user_id,))
    row = c.fetchone()
    if row:
        last_reset = datetime.fromisoformat(row[1])
        if last_reset.tzinfo is None:
            last_reset = last_reset.replace(tzinfo=timezone.utc)
        if now - last_reset > timedelta(days=7):
            c.execute("UPDATE points SET week_points=0, last_reset=? WHERE user_id=?",
                      (now.isoformat(), user_id))
    c.execute("""UPDATE points SET
        total_points=total_points+1, week_points=week_points+1,
        username=?, last_submit=? WHERE user_id=?""",
        (username, now.isoformat(), user_id))
    conn.commit()
    c.execute("SELECT total_points, week_points FROM points WHERE user_id=?", (user_id,))
    result = c.fetchone()
    conn.close()
    return result

def log_submission(user_id, username, photo_count, is_valid, reason):
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("INSERT INTO submissions (user_id,username,submitted_at,photo_count,is_valid,reason) VALUES (?,?,?,?,?,?)",
              (user_id, username, datetime.now(timezone.utc).isoformat(), photo_count, int(is_valid), reason))
    conn.commit(); conn.close()

def get_all_points():
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("SELECT user_id, username, total_points, week_points, last_submit FROM points ORDER BY total_points DESC")
    rows = c.fetchall()
    conn.close()
    return rows

def get_user_point(user_id):
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("SELECT total_points, week_points, last_submit FROM points WHERE user_id=?", (user_id,))
    row = c.fetchone()
    conn.close()
    return row

def get_inactive_users(days=7):
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    threshold = (datetime.now(timezone.utc) - timedelta(days=days)).isoformat()
    c.execute("""SELECT user_id, username, last_submit FROM points
                 WHERE last_submit IS NULL OR last_submit < ?""", (threshold,))
    rows = c.fetchall()
    conn.close()
    return rows

def manual_add_point(user_id, username, amount=1):
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("SELECT user_id FROM points WHERE user_id=?", (user_id,))
    if not c.fetchone():
        c.execute("INSERT INTO points (user_id,username,total_points,week_points,last_reset,last_submit) VALUES (?,?,0,0,?,NULL)",
                  (user_id, username, datetime.now(timezone.utc).isoformat()))
    c.execute("UPDATE points SET total_points=total_points+?, week_points=week_points+?, username=? WHERE user_id=?",
              (amount, amount, username, user_id))
    conn.commit()
    c.execute("SELECT total_points, week_points FROM points WHERE user_id=?", (user_id,))
    result = c.fetchone()
    conn.close()
    return result

def manual_remove_point(user_id, amount=1):
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("SELECT total_points, week_points FROM points WHERE user_id=?", (user_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return None
    new_total = max(0, row[0] - amount)
    new_week  = max(0, row[1] - amount)
    c.execute("UPDATE points SET total_points=?, week_points=? WHERE user_id=?",
              (new_total, new_week, user_id))
    conn.commit()
    conn.close()
    return (new_total, new_week)

# ─── Export ───────────────────────────────────────────────────────────────────
def get_user_photo_count(user_id):
    """Ambil total foto yang pernah disubmit user."""
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("SELECT COALESCE(SUM(photo_count), 0) FROM submissions WHERE user_id=? AND is_valid=1", (user_id,))
    result = c.fetchone()[0]
    conn.close()
    return result

def format_timestamp(ts_str):
    """Format timestamp ISO ke WIB (UTC+7)."""
    if not ts_str:
        return "-"
    try:
        dt = datetime.fromisoformat(ts_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        # Convert ke WIB (UTC+7)
        wib = dt + timedelta(hours=7)
        return wib.strftime("%d/%m/%Y %H:%M WIB")
    except:
        return ts_str

def generate_excel(rows: list) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SSRP Points"
    header_fill = PatternFill("solid", fgColor="2C3E50")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    center = Alignment(horizontal="center", vertical="center")
    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"), bottom=Side(style="thin"))
    headers = ["No", "Username", "User ID", "Total Poin", "Poin Minggu Ini", "Total Foto", "Terakhir Submit"]
    col_widths = [5, 25, 22, 14, 18, 12, 25]
    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = center; cell.border = thin
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 28
    fill_even = PatternFill("solid", fgColor="ECF0F1")
    fill_odd  = PatternFill("solid", fgColor="FFFFFF")
    data_font = Font(name="Arial", size=10)
    for i, (uid, uname, total, week, last_sub) in enumerate(rows, 1):
        row = i + 1
        fill = fill_even if i % 2 == 0 else fill_odd
        photo_count = get_user_photo_count(uid)
        last_str = format_timestamp(last_sub)
        for col, val in enumerate([i, uname, uid, total, week, photo_count, last_str], 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.fill = fill; cell.font = data_font; cell.border = thin
            cell.alignment = Alignment(horizontal="center" if col in (1,4,5,6) else "left", vertical="center")
        ws.row_dimensions[row].height = 22
    summary_row = len(rows) + 2
    ws.cell(row=summary_row, column=1, value="TOTAL").font = Font(bold=True, name="Arial")
    ws.cell(row=summary_row, column=4, value=f"=SUM(D2:D{len(rows)+1})").font = Font(bold=True, name="Arial")
    ws.cell(row=summary_row, column=5, value=f"=SUM(E2:E{len(rows)+1})").font = Font(bold=True, name="Arial")
    ws.cell(row=summary_row, column=6, value=f"=SUM(F2:F{len(rows)+1})").font = Font(bold=True, name="Arial")
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

def generate_csv(rows: list) -> bytes:
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["No", "Username", "User ID", "Total Poin", "Poin Minggu Ini", "Total Foto", "Terakhir Submit"])
    for i, (uid, uname, total, week, last_sub) in enumerate(rows, 1):
        photo_count = get_user_photo_count(uid)
        last_str = format_timestamp(last_sub)
        writer.writerow([i, uname, uid, total, week, photo_count, last_str])
    return buf.getvalue().encode("utf-8-sig")

# ─── Inactive Task ────────────────────────────────────────────────────────────
@tasks.loop(hours=24)
async def check_inactive():
    now = datetime.now(timezone.utc)
    if now.weekday() != INACTIVE_CHECK_WEEKDAY or now.hour != INACTIVE_CHECK_HOUR:
        return
    for guild in bot.guilds:
        inactive_channel = discord.utils.get(guild.text_channels, name=INACTIVE_CHANNEL_NAME)
        if not inactive_channel:
            continue
        inactive = get_inactive_users(days=7)
        if not inactive:
            await inactive_channel.send("✅ **Semua member sudah submit SSRP minggu ini!**")
            continue
        chunks = [inactive[i:i+20] for i in range(0, len(inactive), 20)]
        for chunk_idx, chunk in enumerate(chunks):
            desc = ""
            for uid, uname, last_sub in chunk:
                member = guild.get_member(int(uid))
                mention = member.mention if member else f"**{uname}**"
                last_str = "Belum pernah" if not last_sub else datetime.fromisoformat(last_sub).strftime("%d/%m/%Y")
                desc += f"• {mention} — terakhir submit: {last_str}\n"
            embed = discord.Embed(title=f"⚠️ Member Tidak Aktif SSRP (7 Hari) — Part {chunk_idx+1}",
                                  description=desc, color=0xE74C3C, timestamp=now)
            embed.set_footer(text=f"Total tidak aktif: {len(inactive)} member")
            await inactive_channel.send(embed=embed)
        for uid, uname, last_sub in inactive:
            member = guild.get_member(int(uid))
            if member:
                try:
                    await member.send(
                        f"⚠️ **Hei {member.display_name}!**\n"
                        f"Kamu belum submit SSRP dalam **7 hari terakhir**.\n"
                        f"Jangan lupa kirim screenshot SSRP kamu di server ya! 🎭"
                    )
                except Exception:
                    pass

# ─── Events ───────────────────────────────────────────────────────────────────
@bot.event
async def on_ready():
    init_db()
    check_inactive.start()
    print(f"✅ Bot online: {bot.user}")
    print(f"📋 Monitoring: {', '.join(['#'+c for c in APPROVAL_CHANNELS])}")

@bot.event
async def on_message(message: discord.Message):
    if message.author.bot:
        return
    if message.channel.name not in APPROVAL_CHANNELS:
        await bot.process_commands(message)
        return

    images = [a for a in message.attachments if a.content_type and a.content_type.startswith("image/")]
    if not images:
        await bot.process_commands(message)
        return

    user = message.author
    user_id = str(user.id)
    now = datetime.now(timezone.utc)

    elapsed = (now - last_submit[user_id]).total_seconds()
    if elapsed < SUBMIT_COOLDOWN:
        remaining = int(SUBMIT_COOLDOWN - elapsed)
        await message.add_reaction("⏳")
        await message.reply(
            f"⚠️ **{user.display_name}**, tunggu **{remaining} detik** lagi ya!\n_(Submit ini tidak dihitung)_",
            delete_after=10
        )
        return

    last_submit[user_id] = now
    upsert_user(user_id, user.display_name)
    log_submission(user_id, user.display_name, len(images), True, "Auto approved")
    total_pts, week_pts = add_point(user_id, user.display_name)
    await message.add_reaction("✅")

    report_channel = discord.utils.get(message.guild.text_channels, name=REPORT_CHANNEL_NAME) or message.channel
    msg_link = f"https://discord.com/channels/{message.guild.id}/{message.channel.id}/{message.id}"
    channel_label = "Outrider SSRP" if message.channel.name == "outrider-ssrp-approval" else "SSRP"

    embed = discord.Embed(title=f"🎭 Laporan {channel_label} Baru!", color=0x57F287, timestamp=now)
    embed.set_author(name=user.display_name, icon_url=user.display_avatar.url)
    embed.add_field(name="👤 Player",           value=user.mention,                    inline=True)
    embed.add_field(name="🖼️ Jumlah Foto",      value=f"{len(images)} foto",           inline=True)
    embed.add_field(name="📅 Waktu",            value=f"<t:{int(now.timestamp())}:F>", inline=True)
    embed.add_field(name="📈 Point Minggu Ini", value=f"**{week_pts} Point**",         inline=True)
    embed.add_field(name="🏆 Total Point",      value=f"**{total_pts} Point**",        inline=True)
    embed.add_field(name="🔗 Link Bukti",       value=f"[Klik di sini]({msg_link})",   inline=True)
    embed.set_image(url=images[0].url)
    embed.set_footer(text=f"SSRP Checker Bot • #{message.channel.name}")
    await report_channel.send(embed=embed)
    await bot.process_commands(message)

# ─── Commands ─────────────────────────────────────────────────────────────────
@bot.command(name="point")
async def check_point(ctx, member: discord.Member = None):
    target = member or ctx.author
    row = get_user_point(str(target.id))
    if not row:
        await ctx.reply(f"**{target.display_name}** belum pernah submit SSRP."); return
    total, week, last_sub = row
    last_str = "-"
    if last_sub:
        try: last_str = f"<t:{int(datetime.fromisoformat(last_sub).timestamp())}:R>"
        except: last_str = last_sub
    embed = discord.Embed(title="📊 SSRP Points", color=0x5865F2)
    embed.set_author(name=target.display_name, icon_url=target.display_avatar.url)
    embed.add_field(name="📈 Poin Minggu Ini", value=f"**{week}**",  inline=True)
    embed.add_field(name="🏆 Total Poin",      value=f"**{total}**", inline=True)
    embed.add_field(name="🕐 Terakhir Submit", value=last_str,       inline=True)
    await ctx.reply(embed=embed)

@bot.command(name="leaderboard", aliases=["lb"])
async def leaderboard(ctx):
    rows = get_all_points()
    if not rows:
        await ctx.reply("Belum ada data SSRP."); return
    medals = ["🥇","🥈","🥉"] + ["🏅"]*7
    desc = ""
    for i, (uid, uname, total, week, _) in enumerate(rows[:10]):
        desc += f"{medals[i]} **{uname}** — {week} poin minggu ini _(total: {total})_\n"
    embed = discord.Embed(title="🏆 SSRP Leaderboard — Minggu Ini", description=desc, color=0xFFD700)
    embed.set_footer(text="Reset setiap 7 hari")
    await ctx.reply(embed=embed)

@bot.command(name="export")
@commands.has_any_role(ADMIN_ROLE_ID)
async def export_data(ctx, fmt: str = "excel"):
    rows = get_all_points()
    if not rows:
        await ctx.reply("Belum ada data untuk di-export."); return
    fmt = fmt.lower()
    files = []
    if fmt in ("excel", "xlsx", "all"):
        files.append(discord.File(io.BytesIO(generate_excel(rows)), filename="ssrp_data.xlsx"))
    if fmt in ("csv", "all"):
        files.append(discord.File(io.BytesIO(generate_csv(rows)), filename="ssrp_data.csv"))
    if not files:
        await ctx.reply("Format tidak valid. Gunakan: `excel`, `csv`, atau `all`"); return
    await ctx.reply(f"📊 **Export SSRP Data** — {len(rows)} member", files=files)

@bot.command(name="inactive")
@commands.has_any_role(ADMIN_ROLE_ID)
async def check_inactive_cmd(ctx, days: int = 7):
    inactive = get_inactive_users(days=days)
    if not inactive:
        await ctx.reply(f"✅ Semua member sudah submit SSRP dalam {days} hari terakhir!"); return
    desc = ""
    for uid, uname, last_sub in inactive[:30]:
        member = ctx.guild.get_member(int(uid))
        mention = member.mention if member else f"**{uname}**"
        last_str = "Belum pernah" if not last_sub else datetime.fromisoformat(last_sub).strftime("%d/%m/%Y")
        desc += f"• {mention} — terakhir: {last_str}\n"
    if len(inactive) > 30:
        desc += f"\n_...dan {len(inactive)-30} lainnya. Gunakan `!export` untuk list lengkap._"
    embed = discord.Embed(title=f"⚠️ Tidak Aktif SSRP ({days} Hari Terakhir)", description=desc, color=0xE74C3C)
    embed.set_footer(text=f"Total: {len(inactive)} member tidak aktif")
    await ctx.reply(embed=embed)

@bot.command(name="resetweek")
@commands.has_any_role(ADMIN_ROLE_ID)
async def reset_week(ctx):
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("UPDATE points SET week_points=0, last_reset=?", (datetime.now(timezone.utc).isoformat(),))
    conn.commit(); conn.close()
    await ctx.reply("✅ **Poin mingguan semua member sudah di-reset!**")

async def send_log(guild, action, admin, target, amount, before, after):
    log_channel = discord.utils.get(guild.text_channels, name=LOG_CHANNEL_NAME)
    if not log_channel: return
    color = 0x2ECC71 if action == "ADD" else 0xE74C3C
    icon  = "➕" if action == "ADD" else "➖"
    embed = discord.Embed(title=f"{icon} Poin {action} — Log Admin", color=color, timestamp=datetime.now(timezone.utc))
    embed.add_field(name="👮 Admin",   value=admin.mention,    inline=True)
    embed.add_field(name="👤 Target",  value=target.mention,   inline=True)
    embed.add_field(name="🔢 Jumlah",  value=f"{amount} poin", inline=True)
    embed.add_field(name="📉 Sebelum", value=f"{before} poin", inline=True)
    embed.add_field(name="📈 Sesudah", value=f"{after} poin",  inline=True)
    embed.set_footer(text="SSRP Admin Log")
    await log_channel.send(embed=embed)

@bot.command(name="addpoint")
@commands.has_any_role(ADMIN_ROLE_ID)
async def add_point_cmd(ctx, member: discord.Member = None, amount: int = 1):
    if not member:
        await ctx.reply("❌ Gunakan: `!addpoint @user [jumlah]`"); return
    if amount < 1:
        await ctx.reply("❌ Jumlah harus minimal 1."); return
    before_row = get_user_point(str(member.id))
    before = before_row[0] if before_row else 0
    result = manual_add_point(str(member.id), member.display_name, amount)
    after = result[0]
    embed = discord.Embed(title="➕ Poin Ditambahkan!", color=0x2ECC71)
    embed.set_author(name=member.display_name, icon_url=member.display_avatar.url)
    embed.add_field(name="👤 Player",     value=member.mention,   inline=True)
    embed.add_field(name="➕ Ditambah",   value=f"{amount} poin", inline=True)
    embed.add_field(name="🏆 Total Baru", value=f"{after} poin",  inline=True)
    await ctx.reply(embed=embed)
    await send_log(ctx.guild, "ADD", ctx.author, member, amount, before, after)

@bot.command(name="removepoint")
@commands.has_any_role(ADMIN_ROLE_ID)
async def remove_point_cmd(ctx, member: discord.Member = None, amount: int = 1):
    if not member:
        await ctx.reply("❌ Gunakan: `!removepoint @user [jumlah]`"); return
    if amount < 1:
        await ctx.reply("❌ Jumlah harus minimal 1."); return
    row = get_user_point(str(member.id))
    if not row:
        await ctx.reply(f"❌ **{member.display_name}** belum punya data poin."); return
    before = row[0]
    pending_removals[str(ctx.author.id)] = {
        "target_id": str(member.id), "target_name": member.display_name,
        "target_obj": member, "amount": amount, "before": before,
        "expires": datetime.now(timezone.utc).timestamp() + 30
    }
    embed = discord.Embed(
        title="⚠️ Konfirmasi Hapus Poin",
        description=(
            f"Kamu akan menghapus **{amount} poin** dari {member.mention}\n"
            f"Poin sekarang: **{before}** → setelah: **{max(0, before - amount)}**\n\n"
            f"Ketik **`!confirm`** dalam **30 detik** untuk lanjutkan.\nKetik **`!cancel`** untuk batalkan."
        ),
        color=0xE67E22
    )
    await ctx.reply(embed=embed)

@bot.command(name="confirm")
@commands.has_any_role(ADMIN_ROLE_ID)
async def confirm_remove(ctx):
    admin_id = str(ctx.author.id)
    pending = pending_removals.get(admin_id)
    if not pending:
        await ctx.reply("❌ Tidak ada pending removepoint untuk kamu."); return
    if datetime.now(timezone.utc).timestamp() > pending["expires"]:
        del pending_removals[admin_id]
        await ctx.reply("❌ Waktu konfirmasi habis (30 detik). Ulangi `!removepoint` lagi."); return
    result = manual_remove_point(pending["target_id"], pending["amount"])
    if not result:
        await ctx.reply("❌ Gagal, member tidak ditemukan di database."); return
    after = result[0]; member = pending["target_obj"]
    before = pending["before"]; amount = pending["amount"]
    del pending_removals[admin_id]
    embed = discord.Embed(title="➖ Poin Dihapus!", color=0xE74C3C)
    embed.set_author(name=member.display_name, icon_url=member.display_avatar.url)
    embed.add_field(name="👤 Player",     value=member.mention,   inline=True)
    embed.add_field(name="➖ Dihapus",    value=f"{amount} poin", inline=True)
    embed.add_field(name="🏆 Total Baru", value=f"{after} poin",  inline=True)
    await ctx.reply(embed=embed)
    await send_log(ctx.guild, "REMOVE", ctx.author, member, amount, before, after)

@bot.command(name="cancel")
@commands.has_any_role(ADMIN_ROLE_ID)
async def cancel_remove(ctx):
    admin_id = str(ctx.author.id)
    if admin_id in pending_removals:
        del pending_removals[admin_id]
        await ctx.reply("✅ Removepoint dibatalkan.")
    else:
        await ctx.reply("❌ Tidak ada pending removepoint.")

@bot.command(name="ssrphelp")
async def ssrp_help(ctx):
    embed = discord.Embed(title="🤖 SSRP Checker Bot", color=0x5865F2)
    embed.add_field(name="📋 Cara Submit", value=(
        "Kirim screenshot di:\n"
        "• `#ssrp-approval`\n"
        "• `#outrider-ssrp-approval`\n"
        "Bot langsung kasih poin otomatis!"
    ), inline=False)
    embed.add_field(name="👤 Commands Member", value=(
        "`!point` — Cek poin sendiri\n"
        "`!point @user` — Cek poin member lain\n"
        "`!lb` — Leaderboard minggu ini\n"
        "`!ssrphelp` — Bantuan"
    ), inline=False)
    embed.add_field(name="🔧 Commands Admin", value=(
        "`(bug only)!addpoint @user [n]` — Tambah poin manual\n"
        "`(bug only)!removepoint @user [n]` — Hapus poin (butuh !confirm)\n"
        "`(bug only)!confirm` — Konfirmasi removepoint\n"
        "`(bug only)!cancel` — Batalkan removepoint\n"
        "`!export excel` — Export ke Excel\n"
        "`!export csv` — Export ke CSV\n"
        "`!export all` — Export keduanya\n"
        "`!inactive` — Cek tidak aktif 7 hari\n"
        "`!inactive 14` — Cek tidak aktif 14 hari\n"
        "`!resetweek` — Reset poin mingguan"
    ), inline=False)
    embed.add_field(name="⚙️ Sistem", value=(
        f"• 1 submit = 1 poin (berapapun jumlah foto)\n"
        f"• Cooldown: **{SUBMIT_COOLDOWN} detik** antar submit\n"
        f"• Cooldown berlaku di semua channel submit\n"
        f"• Cek inaktif otomatis setiap Senin pagi"
    ), inline=False)
    await ctx.reply(embed=embed)

# ─── Run ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    bot.run(DISCORD_TOKEN)
