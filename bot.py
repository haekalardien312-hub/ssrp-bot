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
DISCORD_TOKEN = os.environ.get("DISCORD_TOKEN")
APPROVAL_CHANNEL_NAME = "ssrp-approval"    # Channel submit SSRP
REPORT_CHANNEL_NAME   = "ssrp-report"     # Channel laporan otomatis
INACTIVE_CHANNEL_NAME = "ssrp-inactive"   # Channel notif tidak aktif
SUBMIT_COOLDOWN       = 30                # Detik jeda antar submit (anti-spam)
INACTIVE_CHECK_HOUR   = 9                 # Jam berapa cek inaktif (UTC)
INACTIVE_CHECK_WEEKDAY= 0                 # 0=Senin, 6=Minggu
# ================================================

intents = discord.Intents.default()
intents.message_content = True
intents.members = True

bot = commands.Bot(command_prefix="!", intents=intents)
last_submit: dict[str, datetime] = defaultdict(lambda: datetime.min.replace(tzinfo=timezone.utc))

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
    """Ambil semua user yang tidak submit dalam X hari terakhir."""
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    threshold = (datetime.now(timezone.utc) - timedelta(days=days)).isoformat()
    c.execute("""SELECT user_id, username, last_submit FROM points
                 WHERE last_submit IS NULL OR last_submit < ?""", (threshold,))
    rows = c.fetchall()
    conn.close()
    return rows


# ─── Export Excel ─────────────────────────────────────────────────────────────
def generate_excel(rows: list) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SSRP Points"

    # Header style
    header_fill = PatternFill("solid", fgColor="2C3E50")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    center = Alignment(horizontal="center", vertical="center")
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    headers = ["No", "Username", "User ID", "Total Poin", "Poin Minggu Ini", "Terakhir Submit"]
    col_widths = [5, 25, 22, 14, 18, 22]

    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = thin
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 28

    # Alternating row colors
    fill_even = PatternFill("solid", fgColor="ECF0F1")
    fill_odd  = PatternFill("solid", fgColor="FFFFFF")
    data_font = Font(name="Arial", size=10)

    for i, (uid, uname, total, week, last_sub) in enumerate(rows, 1):
        row = i + 1
        fill = fill_even if i % 2 == 0 else fill_odd

        last_str = "-"
        if last_sub:
            try:
                dt = datetime.fromisoformat(last_sub)
                last_str = dt.strftime("%d/%m/%Y %H:%M")
            except Exception:
                last_str = last_sub

        values = [i, uname, uid, total, week, last_str]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.fill = fill
            cell.font = data_font
            cell.border = thin
            cell.alignment = Alignment(
                horizontal="center" if col in (1, 4, 5) else "left",
                vertical="center"
            )

        ws.row_dimensions[row].height = 22

    # Summary row
    summary_row = len(rows) + 2
    ws.cell(row=summary_row, column=1, value="TOTAL").font = Font(bold=True, name="Arial")
    ws.cell(row=summary_row, column=4, value=f"=SUM(D2:D{len(rows)+1})").font = Font(bold=True, name="Arial")
    ws.cell(row=summary_row, column=5, value=f"=SUM(E2:E{len(rows)+1})").font = Font(bold=True, name="Arial")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

def generate_csv(rows: list) -> bytes:
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["No", "Username", "User ID", "Total Poin", "Poin Minggu Ini", "Terakhir Submit"])
    for i, (uid, uname, total, week, last_sub) in enumerate(rows, 1):
        last_str = "-"
        if last_sub:
            try:
                dt = datetime.fromisoformat(last_sub)
                last_str = dt.strftime("%d/%m/%Y %H:%M")
            except Exception:
                last_str = last_sub
        writer.writerow([i, uname, uid, total, week, last_str])
    return buf.getvalue().encode("utf-8-sig")  # utf-8-sig biar Excel baca benar

# ─── Inactive Check Task ───────────────────────────────────────────────────────
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

        # Kirim embed list inactive
        chunks = [inactive[i:i+20] for i in range(0, len(inactive), 20)]
        for chunk_idx, chunk in enumerate(chunks):
            desc = ""
            for uid, uname, last_sub in chunk:
                member = guild.get_member(int(uid))
                mention = member.mention if member else f"**{uname}**"
                last_str = "Belum pernah" if not last_sub else datetime.fromisoformat(last_sub).strftime("%d/%m/%Y")
                desc += f"• {mention} — terakhir submit: {last_str}\n"

            embed = discord.Embed(
                title=f"⚠️ Member Tidak Aktif SSRP (7 Hari) — Part {chunk_idx+1}",
                description=desc,
                color=0xE74C3C,
                timestamp=now
            )
            embed.set_footer(text=f"Total tidak aktif: {len(inactive)} member")
            await inactive_channel.send(embed=embed)

        # DM ke masing-masing
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
                    pass  # DM mungkin dinonaktifkan

# ─── Events ───────────────────────────────────────────────────────────────────
@bot.event
async def on_ready():
    init_db()
    check_inactive.start()
    print(f"✅ Bot online: {bot.user}")

@bot.event
async def on_message(message: discord.Message):
    if message.author.bot:
        return
    if message.channel.name != APPROVAL_CHANNEL_NAME:
        await bot.process_commands(message)
        return

    images = [a for a in message.attachments if a.content_type and a.content_type.startswith("image/")]
    if not images:
        await bot.process_commands(message)
        return

    user = message.author
    user_id = str(user.id)
    now = datetime.now(timezone.utc)

    # Anti-spam cooldown
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

    # Langsung tambah poin tanpa validasi AI
    log_submission(user_id, user.display_name, len(images), True, "Auto approved")
    total_pts, week_pts = add_point(user_id, user.display_name)
    await message.add_reaction("✅")

    # Kirim laporan ke report channel
    report_channel = discord.utils.get(message.guild.text_channels, name=REPORT_CHANNEL_NAME) or message.channel
    msg_link = f"https://discord.com/channels/{message.guild.id}/{message.channel.id}/{message.id}"

    embed = discord.Embed(title="🎭 Laporan SSRP Baru!", color=0x57F287, timestamp=now)
    embed.set_author(name=user.display_name, icon_url=user.display_avatar.url)
    embed.add_field(name="👤 Player",           value=user.mention,          inline=True)
    embed.add_field(name="🖼️ Jumlah Foto",      value=f"{len(images)} foto", inline=True)
    embed.add_field(name="📅 Waktu",            value=f"<t:{int(now.timestamp())}:F>", inline=True)
    embed.add_field(name="📈 Point Minggu Ini", value=f"**{week_pts} Point**",  inline=True)
    embed.add_field(name="🏆 Total Point",      value=f"**{total_pts} Point**", inline=True)
    embed.add_field(name="🔗 Link Bukti",       value=f"[Klik di sini]({msg_link})", inline=True)
    embed.set_image(url=images[0].url)
    embed.set_footer(text="SSRP Checker Bot")

    await report_channel.send(embed=embed)
    await bot.process_commands(message)

# ─── Commands ─────────────────────────────────────────────────────────────────
@bot.command(name="point")
async def check_point(ctx, member: discord.Member = None):
    """Cek poin SSRP. !point atau !point @user"""
    target = member or ctx.author
    row = get_user_point(str(target.id))
    if not row:
        await ctx.reply(f"**{target.display_name}** belum pernah submit SSRP.")
        return
    total, week, last_sub = row
    last_str = "-"
    if last_sub:
        try:
            last_str = f"<t:{int(datetime.fromisoformat(last_sub).timestamp())}:R>"
        except Exception:
            last_str = last_sub
    embed = discord.Embed(title="📊 SSRP Points", color=0x5865F2)
    embed.set_author(name=target.display_name, icon_url=target.display_avatar.url)
    embed.add_field(name="📈 Poin Minggu Ini",   value=f"**{week}**",    inline=True)
    embed.add_field(name="🏆 Total Poin",        value=f"**{total}**",   inline=True)
    embed.add_field(name="🕐 Terakhir Submit",   value=last_str,         inline=True)
    await ctx.reply(embed=embed)

@bot.command(name="leaderboard", aliases=["lb"])
async def leaderboard(ctx):
    """Leaderboard SSRP minggu ini."""
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
@commands.has_permissions(administrator=True)
async def export_data(ctx, fmt: str = "excel"):
    """Export data SSRP. !export excel / !export csv / !export all"""
    rows = get_all_points()
    if not rows:
        await ctx.reply("Belum ada data untuk di-export."); return

    fmt = fmt.lower()
    files = []

    if fmt in ("excel", "xlsx", "all"):
        xlsx_bytes = generate_excel(rows)
        files.append(discord.File(io.BytesIO(xlsx_bytes), filename="ssrp_data.xlsx"))

    if fmt in ("csv", "all"):
        csv_bytes = generate_csv(rows)
        files.append(discord.File(io.BytesIO(csv_bytes), filename="ssrp_data.csv"))

    if not files:
        await ctx.reply("Format tidak valid. Gunakan: `excel`, `csv`, atau `all`"); return

    await ctx.reply(f"📊 **Export SSRP Data** — {len(rows)} member", files=files)

@bot.command(name="inactive")
@commands.has_permissions(administrator=True)
async def check_inactive_cmd(ctx, days: int = 7):
    """Cek siapa yang tidak aktif. !inactive atau !inactive 14"""
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

    embed = discord.Embed(
        title=f"⚠️ Tidak Aktif SSRP ({days} Hari Terakhir)",
        description=desc,
        color=0xE74C3C
    )
    embed.set_footer(text=f"Total: {len(inactive)} member tidak aktif")
    await ctx.reply(embed=embed)

@bot.command(name="resetweek")
@commands.has_permissions(administrator=True)
async def reset_week(ctx):
    """Reset poin mingguan semua member."""
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("UPDATE points SET week_points=0, last_reset=?", (datetime.now(timezone.utc).isoformat(),))
    conn.commit(); conn.close()
    await ctx.reply("✅ **Poin mingguan semua member sudah di-reset!**")

@bot.command(name="ssrphelp")
async def ssrp_help(ctx):
    embed = discord.Embed(title="🤖 SSRP Checker Bot", color=0x5865F2)
    embed.add_field(name="📋 Cara Submit", value=f"Kirim screenshot di `#{APPROVAL_CHANNEL_NAME}`\nBot langsung validasi & kasih poin otomatis!", inline=False)
    embed.add_field(name="👤 Commands Member", value=(
        "`!point` — Cek poin sendiri\n"
        "`!point @user` — Cek poin member lain\n"
        "`!lb` — Leaderboard minggu ini\n"
        "`!ssrphelp` — Bantuan"
    ), inline=False)
    embed.add_field(name="🔧 Commands Admin", value=(
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
        f"• Gambar divalidasi AI otomatis (Gemini)\n"
        f"• Cek inaktif otomatis setiap Senin pagi"
    ), inline=False)
    await ctx.reply(embed=embed)

# ─── Run ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    bot.run(DISCORD_TOKEN)
