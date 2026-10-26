"""
Gmail contact -> separate PDF files (GUI)

Searches ALL Gmail folders/labels (Inbox, Sent, custom labels, All Mail and
optionally Spam/Trash) for emails where a given address appears in From, To,
Cc or Bcc, and saves each unique email as its own PDF.

PDFs are rendered with Google Chrome / Microsoft Edge (headless), so Arabic,
other right-to-left scripts, emoji and complex HTML emails print correctly.

Requirements:
    - Python 3 (no extra pip packages needed)
    - Google Chrome, Microsoft Edge, Brave or Chromium installed
    - Gmail App Password (https://myaccount.google.com/apppasswords)
      Use that 16-character password here, NOT your normal Google password.

Run:
    python gmail_sender_to_pdf.py
"""

import base64
import datetime
import email
import html
import imaplib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import urllib.request
from email import policy
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from urllib.parse import unquote

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# ------------------------------------------------------------ expiry date
# The program stops working after this date. Change it, then rebuild the app.
EXPIRY_DATE = datetime.date(2026, 11, 15)      # year, month, day


def current_date():
    """Today's date. Uses internet time when reachable (it cannot be changed
    by editing the PC clock); falls back to the computer's own date."""
    try:
        req = urllib.request.Request("https://www.google.com", method="HEAD")
        with urllib.request.urlopen(req, timeout=5) as resp:
            return parsedate_to_datetime(resp.headers["Date"]).date()
    except Exception:
        return datetime.date.today()


def is_expired(today=None):
    return (today or current_date()) > EXPIRY_DATE


# ----------------------------------------------------------------- helpers
def imap_date(text):
    """Convert YYYY-MM-DD to the DD-Mon-YYYY format IMAP expects."""
    y, m, d = text.strip().split("-")
    return f"{int(d):02d}-{MONTHS[int(m) - 1]}-{y}"


def safe_name(text, limit=80):
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", text).strip(" ._")
    return (text or "no_subject")[:limit]


def find_browser():
    cands = []
    if sys.platform.startswith("win"):
        for base in filter(None, [os.environ.get("PROGRAMFILES(X86)"),
                                  os.environ.get("PROGRAMFILES"),
                                  os.environ.get("LOCALAPPDATA")]):
            cands += [
                base + r"\Microsoft\Edge\Application\msedge.exe",
                base + r"\Google\Chrome\Application\chrome.exe",
                base + r"\BraveSoftware\Brave-Browser\Application\brave.exe",
            ]
    elif sys.platform == "darwin":
        cands = [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
        ]
    else:
        for name in ("google-chrome", "chromium", "chromium-browser",
                     "microsoft-edge", "brave-browser"):
            p = shutil.which(name)
            if p:
                cands.append(p)
    for c in cands:
        if os.path.exists(c):
            return c
    return ""


def safe_content(part):
    try:
        return part.get_content()
    except Exception:
        payload = part.get_payload(decode=True) or b""
        return payload.decode("utf-8", "replace")


def cid_map(msg):
    """Map Content-ID -> data: URI so inline images show up in the PDF."""
    out = {}
    for part in msg.walk():
        cid = part.get("Content-ID")
        if cid and part.get_content_maintype() == "image":
            data = part.get_payload(decode=True)
            if data:
                out[str(cid).strip().strip("<>")] = (
                    f"data:{part.get_content_type()};base64,"
                    f"{base64.b64encode(data).decode()}")
    return out


def addr_set(msg, header):
    values = [str(h) for h in msg.get_all(header, [])]
    return {a.lower() for _, a in getaddresses(values) if a}


def find_roles(msg, target):
    """Return which headers (From/To/Cc/Bcc) contain the target address."""
    t = target.lower()
    exact = re.fullmatch(r"[^@\s]+@[^@\s]+", t) is not None
    roles = []
    for label, hdr in (("From", "from"), ("To", "to"),
                       ("Cc", "cc"), ("Bcc", "bcc")):
        addrs = addr_set(msg, hdr)
        hit = (t in addrs) if exact else any(t in a for a in addrs)
        if hit:
            roles.append(label)
    return roles


def build_html(msg, load_images, roles):
    def row(label, value):
        if not value:
            return ""
        return (f'<div dir="auto"><b>{label}:</b> '
                f"{html.escape(str(value))}</div>")

    header = (
        '<div style="font-family:Arial,Tahoma,sans-serif;font-size:12px;'
        'border-bottom:1px solid #888;padding-bottom:8px;margin-bottom:14px">'
        f'<div dir="auto" style="font-size:16px;font-weight:bold">'
        f"{html.escape(str(msg['subject'] or '(no subject)'))}</div>"
        + row("From", msg["from"]) + row("To", msg["to"])
        + row("Cc", msg["cc"]) + row("Bcc", msg["bcc"])
        + row("Date", msg["date"])
        + (f'<div style="color:#555">Address found in: '
           f'{", ".join(roles)}</div>' if roles else "")
        + "</div>"
    )

    img_src = "data: http: https:" if load_images else "data:"
    csp = (f"default-src 'none'; img-src {img_src}; "
           "style-src 'unsafe-inline'; font-src data:")
    head = (
        '<meta charset="utf-8">'
        f'<meta http-equiv="Content-Security-Policy" content="{csp}">'
        "<style>@page{size:A4;margin:1.5cm}"
        "img{max-width:100%;height:auto}"
        "body{word-wrap:break-word}</style>"
    )

    part = msg.get_body(preferencelist=("html", "plain"))
    if part is None:
        content = "<html><body><p><i>(no body)</i></p></body></html>"
    elif part.get_content_type() == "text/plain":
        lines = safe_content(part).splitlines() or [""]
        body = "".join(
            f'<div dir="auto" style="white-space:pre-wrap">'
            f"{html.escape(l) or '&nbsp;'}</div>" for l in lines)
        content = ("<html><body style=\"font-family:Arial,Tahoma,"
                   f"sans-serif;font-size:13px\">{body}</body></html>")
    else:
        content = safe_content(part)
        cmap = cid_map(msg)
        if cmap:
            content = re.sub(
                r"cid:([^\"'\s>)]+)",
                lambda mo: cmap.get(unquote(mo.group(1)), mo.group(0)),
                content)

    m = re.match(r"\s*(<!doctype[^>]*>)", content, re.I)
    if m:
        doctype, content = m.group(1), content[m.end():]
    else:
        doctype = "<!DOCTYPE html>"

    m = re.search(r"<body[^>]*>", content, re.I)
    if m:
        content = content[:m.end()] + header + content[m.end():]
    else:
        content = header + content
    return doctype + head + content


def render_pdf(browser, html_text, pdf_path, workdir, load_images):
    html_file = os.path.join(workdir, "msg.html")
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(html_text)
    # Chrome writes into the temp work folder (always allowed); this program
    # then moves the file to the chosen folder. This avoids macOS blocking
    # Chrome from writing into Documents / Desktop / Downloads.
    tmp_pdf = os.path.join(workdir, "out.pdf")
    if os.path.exists(tmp_pdf):
        os.remove(tmp_pdf)
    cmd = [
        browser, "--headless=new", "--disable-gpu", "--no-first-run",
        "--disable-extensions",
        f"--user-data-dir={os.path.join(workdir, 'profile')}",
        "--no-pdf-header-footer", "--print-to-pdf-no-header",
        f"--print-to-pdf={tmp_pdf}",
    ]
    if load_images:
        cmd.append("--virtual-time-budget=15000")
    cmd.append(Path(html_file).as_uri())
    flags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
    subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                   timeout=120, creationflags=flags)
    if not (os.path.exists(tmp_pdf) and os.path.getsize(tmp_pdf) > 0):
        return False
    if os.path.exists(pdf_path):
        os.remove(pdf_path)
    shutil.move(tmp_pdf, pdf_path)
    return True


def list_folders(imap, include_junk):
    """Return the names of every selectable folder/label."""
    status, data = imap.list()
    folders = []
    for raw in data or []:
        if not raw:
            continue
        line = raw.decode("utf-8", "ignore") if isinstance(raw, bytes) else raw
        m = re.match(r'\((?P<flags>[^)]*)\)\s+"(?P<d>[^"]*)"\s+(?P<name>.+)$',
                     line)
        if not m:
            continue
        flags = m.group("flags").lower()
        name = m.group("name").strip()
        if name.startswith('"') and name.endswith('"'):
            name = name[1:-1]
        if "\\noselect" in flags:
            continue
        if not include_junk and ("\\junk" in flags or "\\trash" in flags):
            continue
        folders.append(name)
    return folders


def build_criteria(v):
    q = f'"{v["sender"]}"'
    fields = {"from": ["FROM"], "to": ["TO", "CC", "BCC"],
              "both": ["FROM", "TO", "CC", "BCC"]}[v["mode"]]
    criteria = ["OR"] * (len(fields) - 1)
    for f in fields:
        criteria += [f, q]
    if v["since"]:
        criteria += ["SINCE", imap_date(v["since"])]
    if v["before"]:
        criteria += ["BEFORE", imap_date(v["before"])]
    return criteria


ALLOWED_ROLES = {"from": {"From"}, "to": {"To", "Cc", "Bcc"},
                 "both": {"From", "To", "Cc", "Bcc"}}


# --------------------------------------------------------------------- GUI
class Tooltip:
    """Small hover tooltip for any Tk widget."""

    def __init__(self, widget, text):
        self.widget, self.text, self.tip = widget, text, None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _event=None):
        if self.tip:
            return
        x = self.widget.winfo_rootx() + 16
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.text, justify="left",
                 background="#ffffe0", relief="solid", borderwidth=1,
                 padx=8, pady=6, wraplength=380).pack()

    def hide(self, _event=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


APP_PASSWORD_HELP = (
    "How to generate a Gmail App Password:\n"
    "1. Go to myaccount.google.com > Security and turn on "
    "2-Step Verification.\n"
    "2. Open myaccount.google.com/apppasswords (or search "
    "\"App passwords\" in your Google Account).\n"
    "3. Enter a name such as \"Gmail PDF exporter\" and click Create.\n"
    "4. Copy the 16-character password shown and paste it here.\n"
    "Use this instead of your normal Gmail password. Spaces are ignored."
)


class App:
    def __init__(self, root):
        self.root = root
        root.title("Extract Gmail to PDF with Attachments")
        root.geometry("700x900")

        self.pause_ev = threading.Event()
        self.pause_ev.set()          # set = running, cleared = paused
        self.stop_ev = threading.Event()

        ttk.Label(root, text="Extract Gmail to PDF with Attachments",
                  font=("Segoe UI", 20, "bold"), anchor="center",
                  foreground="blue").pack(fill="x", pady=(14, 0))

        ttk.Label(root, text="Developed by Sanjeev",
                  font=("Segoe UI", 10, "italic", "bold"), anchor="center",
                  foreground="darkred").pack(fill="x", pady=(0, 0))
#        ttk.Label(root, text=f"Valid until {EXPIRY_DATE:%d %b %Y}",
#                  font=("Segoe UI", 9), anchor="center",
#                  foreground="gray30").pack(fill="x", pady=(0, 4))

        frm = ttk.Frame(root, padding=12)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        self.vars = {k: tk.StringVar() for k in
                     ("user", "password", "sender", "since", "before")}
        self.vars["outdir"] = tk.StringVar(
            value=os.path.join(os.path.expanduser("~"), "email_pdfs"))
        self.vars["browser"] = tk.StringVar(value=find_browser())
        self.mode = tk.StringVar(value="both")
        self.attach = tk.BooleanVar(value=True)
        self.images = tk.BooleanVar(value=False)
        self.junk = tk.BooleanVar(value=False)
        self.latest = tk.BooleanVar(value=True)

        rows = [
            ("Your Gmail address", "user", False),
            ("App password", "password", True),
            ("Email address to export", "sender", False),
            ("From date (YYYY-MM-DD, optional)", "since", False),
            ("Until date (YYYY-MM-DD, optional)", "before", False),
        ]
        for i, (label, key, secret) in enumerate(rows):
            if key == "password":
                lf = ttk.Frame(frm)
                lf.grid(row=i, column=0, sticky="w", pady=4)
                ttk.Label(lf, text=label).pack(side="left")
                star = tk.Label(lf, text="*", fg="red", cursor="question_arrow",
                                font=("Segoe UI", 11, "bold"))
                star.pack(side="left", padx=(2, 0))
                Tooltip(star, APP_PASSWORD_HELP)
            else:
                ttk.Label(frm, text=label).grid(row=i, column=0, sticky="w",
                                                pady=4)
            ttk.Entry(frm, textvariable=self.vars[key],
                      show="*" if secret else "").grid(
                row=i, column=1, columnspan=2, sticky="ew", pady=4)

        r = len(rows)
        ttk.Label(frm, text="Save PDFs to").grid(row=r, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.vars["outdir"]).grid(
            row=r, column=1, sticky="ew", pady=4)
        ttk.Button(frm, text="Browse…", command=self.browse_dir).grid(
            row=r, column=2, padx=(6, 0))

        ttk.Label(frm, text="Chrome/Edge location").grid(
            row=r + 1, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.vars["browser"]).grid(
            row=r + 1, column=1, sticky="ew", pady=4)
        ttk.Button(frm, text="Browse…", command=self.browse_exe).grid(
            row=r + 1, column=2, padx=(6, 0))

        box = ttk.LabelFrame(frm, text="Which emails?", padding=8)
        box.grid(row=r + 2, column=0, columnspan=3, sticky="ew", pady=8)
        ttk.Radiobutton(box, text="Received from this address (From)",
                        variable=self.mode, value="from").pack(anchor="w")
        ttk.Radiobutton(box, text="Sent to this address (To / Cc / Bcc)",
                        variable=self.mode, value="to").pack(anchor="w")
        ttk.Radiobutton(box, text="All: address in From, To, Cc or Bcc",
                        variable=self.mode, value="both").pack(anchor="w")

        ttk.Checkbutton(frm, text="Also save attachments",
                        variable=self.attach).grid(
            row=r + 3, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(frm, text="Include Spam and Trash folders",
                        variable=self.junk).grid(
            row=r + 4, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(frm, text="Load remote images (slower; embedded "
                        "images always show)", variable=self.images).grid(
            row=r + 5, column=0, columnspan=3, sticky="w")

        ttk.Checkbutton(frm, text="Only export the latest email of each "
                        "conversation (skip earlier emails in the trail)",
                        variable=self.latest).grid(
            row=r + 6, column=0, columnspan=3, sticky="w")

        bf = ttk.Frame(frm)
        bf.grid(row=r + 7, column=0, columnspan=3, pady=10)
        self.btn = ttk.Button(bf, text="Export PDFs", command=self.start)
        self.btn.pack(side="left", padx=4)
        self.btn_pause = ttk.Button(bf, text="Pause", command=self.toggle_pause,
                                    state="disabled")
        self.btn_pause.pack(side="left", padx=4)
        self.btn_stop = ttk.Button(bf, text="Stop", command=self.stop,
                                   state="disabled")
        self.btn_stop.pack(side="left", padx=4)

        self.bar = ttk.Progressbar(frm, mode="determinate")
        self.bar.grid(row=r + 8, column=0, columnspan=3, sticky="ew")

        self.log_box = scrolledtext.ScrolledText(frm, height=14,
                                                 state="disabled")
        self.log_box.grid(row=r + 9, column=0, columnspan=3, sticky="nsew",
                          pady=(8, 0))
        frm.rowconfigure(r + 9, weight=1)

    # ---- UI helpers (safe to call from worker thread) ----
    def log(self, text):
        def _do():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", text + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.root.after(0, _do)

    def set_progress(self, value, maximum):
        self.root.after(0, lambda: self.bar.configure(maximum=maximum,
                                                      value=value))

    def set_running(self, running):
        def _do():
            self.btn.configure(state="disabled" if running else "normal")
            self.btn_pause.configure(state="normal" if running else "disabled",
                                     text="Pause")
            self.btn_stop.configure(state="normal" if running else "disabled")
        self.root.after(0, _do)

    def browse_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.vars["outdir"].set(d)

    def browse_exe(self):
        f = filedialog.askopenfilename(title="Select Chrome / Edge program")
        if f:
            self.vars["browser"].set(f)

    # ---- pause / stop ----
    def toggle_pause(self):
        if self.pause_ev.is_set():
            self.pause_ev.clear()
            self.btn_pause.configure(text="Resume")
            self.log("Paused… click Resume to continue.")
        else:
            self.pause_ev.set()
            self.btn_pause.configure(text="Pause")
            self.log("Resumed.")

    def stop(self):
        self.stop_ev.set()
        self.pause_ev.set()          # release a paused worker
        self.log("Stopping after the current email…")

    def checkpoint(self):
        """Block while paused; return True if the user asked to stop."""
        self.pause_ev.wait()
        return self.stop_ev.is_set()

    # ---- main flow ----
    def start(self):
        if is_expired():
            messagebox.showerror(
                "Program expired",
                f"This program expired on {EXPIRY_DATE:%d %B %Y}.\n"
                "Please contact the developer (Sanjeev) for a new version.")
            return
        v = {k: x.get().strip() for k, x in self.vars.items()}
        if not (v["user"] and v["password"] and v["sender"]):
            messagebox.showwarning("Missing details",
                                   "Gmail address, app password and email "
                                   "address to export are required.")
            return
        if not (v["browser"] and os.path.exists(v["browser"])):
            messagebox.showwarning(
                "Browser not found",
                "Install Google Chrome or Microsoft Edge, or use Browse to "
                "select its program file.")
            return
        for key in ("since", "before"):
            if v[key]:
                try:
                    imap_date(v[key])
                except Exception:
                    messagebox.showwarning("Bad date",
                                           "Use the format YYYY-MM-DD.")
                    return
        v["mode"] = self.mode.get()
        v["junk"] = self.junk.get()
        v["attach"] = self.attach.get()
        v["images"] = self.images.get()
        v["latest"] = self.latest.get()

        self.stop_ev.clear()
        self.pause_ev.set()
        self.set_running(True)
        threading.Thread(target=self.run, args=(v,), daemon=True).start()

    def collect(self, imap, folders, criteria, latest_only=False):
        """Search every folder; return [(folder, [msg numbers])].

        Emails are deduplicated across folders using Gmail's unique message
        id (X-GM-MSGID). With latest_only, only the newest matching email of
        each Gmail conversation (X-GM-THRID) is kept."""
        seen = set()
        plan = []
        best = {}        # thread id -> (timestamp, folder, msg number)
        total_matches = 0
        for folder in folders:
            if self.checkpoint():
                return None
            status, _ = imap.select(f'"{folder}"', readonly=True)
            if status != "OK":
                self.log(f"Skipped folder (cannot open): {folder}")
                continue
            status, data = imap.search(None, *criteria)
            nums = data[0].split() if status == "OK" and data[0] else []
            if not nums:
                continue

            new = []
            for i in range(0, len(nums), 500):
                chunk = nums[i:i + 500]
                st, resp = imap.fetch(
                    b",".join(chunk), "(X-GM-MSGID X-GM-THRID INTERNALDATE)")
                for item in resp or []:
                    text = item[0] if isinstance(item, tuple) else item
                    if not isinstance(text, bytes):
                        continue
                    m = re.match(rb"(\d+) ", text)
                    g = re.search(rb"X-GM-MSGID (\d+)", text)
                    if not (m and g):
                        continue
                    gid = g.group(1)
                    if gid in seen:
                        continue
                    seen.add(gid)
                    total_matches += 1
                    if latest_only:
                        t = re.search(rb"X-GM-THRID (\d+)", text)
                        thr = t.group(1) if t else gid
                        try:
                            ts = time.mktime(imaplib.Internaldate2tuple(text))
                        except Exception:
                            ts = 0
                        cur = best.get(thr)
                        if cur is None or ts > cur[0]:
                            best[thr] = (ts, folder, m.group(1))
                    else:
                        new.append(m.group(1))
            if latest_only:
                self.log(f"{folder}: {len(nums)} match(es)")
            else:
                if new:
                    plan.append((folder, new))
                self.log(f"{folder}: {len(nums)} match(es), {len(new)} new")

        if latest_only:
            by_folder = {}
            for ts, folder, num in best.values():
                by_folder.setdefault(folder, []).append(num)
            plan = list(by_folder.items())
            self.log(f"\n{total_matches} unique email(s) collapsed into "
                     f"{len(best)} conversation(s); keeping the latest "
                     "email of each.")
        return plan

    def run(self, v):
        workdir = tempfile.mkdtemp(prefix="gmail_pdf_")
        try:
            try:
                os.makedirs(v["outdir"], exist_ok=True)
                probe = os.path.join(v["outdir"], ".write_test")
                with open(probe, "w") as f:
                    f.write("ok")
                os.remove(probe)
            except OSError as e:
                self.log(
                    f"Cannot write to the folder:\n  {v['outdir']}\n({e})\n"
                    "On a Mac: choose a different folder with Browse (for "
                    "example a new folder inside Documents) and click Allow "
                    "if macOS asks for access. If it keeps failing, open "
                    "System Settings > Privacy & Security > Files and "
                    "Folders (or Full Disk Access) and allow this app.")
                return
            self.log("Connecting to Gmail…")
            imap = imaplib.IMAP4_SSL("imap.gmail.com")
            imap.login(v["user"], v["password"].replace(" ", ""))

            folders = list_folders(imap, v["junk"])
            self.log(f"Searching {len(folders)} folders…")
            plan = self.collect(imap, folders, build_criteria(v), v["latest"])
            if plan is None:
                self.log("Stopped.")
                imap.logout()
                return
            total = sum(len(n) for _, n in plan)
            self.log(f"\n{total} candidate email(s) found.\n")
            if not total:
                imap.logout()
                return

            user_l = v["user"].lower()
            allowed = ALLOWED_ROLES[v["mode"]]
            ok = failed = skipped = done = 0
            stopped = False

            for folder, nums in plan:
                imap.select(f'"{folder}"', readonly=True)
                for num in nums:
                    if self.checkpoint():
                        stopped = True
                        break
                    done += 1
                    try:
                        status, parts = imap.fetch(num, "(BODY.PEEK[])")
                        raw = parts[0][1]
                        msg = email.message_from_bytes(raw,
                                                       policy=policy.default)

                        roles = find_roles(msg, v["sender"])
                        if not set(roles) & allowed:
                            skipped += 1
                            self.log(f"[{done}/{total}] skipped (address "
                                     "only a partial match)")
                            self.set_progress(done, total)
                            continue

                        try:
                            dt = parsedate_to_datetime(str(msg["date"]))
                            stamp = dt.strftime("%Y-%m-%d_%H%M%S")
                        except Exception:
                            stamp = "undated"
                        direction = ("SENT" if user_l in
                                     str(msg["from"] or "").lower()
                                     else "RECEIVED")
                        base = (f"{ok + 1:04d}_{stamp}_{direction}_"
                                f"{safe_name(str(msg['subject'] or ''))}")
                        pdf_path = os.path.join(v["outdir"], base + ".pdf")

                        good = render_pdf(
                            v["browser"],
                            build_html(msg, v["images"], roles),
                            pdf_path, workdir, v["images"])

                        if v["attach"]:
                            self.save_attachments(msg, v["outdir"], base)

                        if good:
                            ok += 1
                            self.log(f"[{done}/{total}] saved {base}.pdf  "
                                     f"(in: {', '.join(roles)})")
                        else:
                            failed += 1
                            self.log(f"[{done}/{total}] FAILED {base}")
                    except Exception as e:
                        failed += 1
                        self.log(f"[{done}/{total}] error: {e}")
                    self.set_progress(done, total)
                if stopped:
                    break

            imap.logout()
            self.log(("\nStopped by user. " if stopped else "\nDone. ")
                     + f"{ok} saved, {failed} failed, {skipped} skipped.\n"
                     f"Folder: {v['outdir']}")
        except imaplib.IMAP4.error as e:
            self.log(f"Login/IMAP error: {e}\nCheck that you are using an "
                     "App Password and that IMAP is enabled in Gmail "
                     "settings.")
        except Exception as e:
            self.log(f"Error: {e}")
        finally:
            shutil.rmtree(workdir, ignore_errors=True)
            self.set_running(False)

    @staticmethod
    def save_attachments(msg, outdir, base):
        folder = None
        for att in msg.iter_attachments():
            name = att.get_filename()
            if not name:
                continue
            if folder is None:
                folder = os.path.join(outdir, base + "_attachments")
                os.makedirs(folder, exist_ok=True)
            with open(os.path.join(folder, safe_name(name, 120)), "wb") as f:
                f.write(att.get_payload(decode=True) or b"")


if __name__ == "__main__":
    root = tk.Tk()
    if is_expired():
        root.withdraw()
        messagebox.showerror(
            "Program expired",
            f"This program expired on {EXPIRY_DATE:%d %B %Y}.\n"
            "Please contact the developer (Sanjeev) for a new version.")
        root.destroy()
        sys.exit(0)
    App(root)
    root.mainloop()
