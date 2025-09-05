import os, sys, time, uuid, threading, ctypes, subprocess, tempfile
from datetime import datetime, timedelta
from configparser import ConfigParser


def base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_settings():
    p = ConfigParser()
    path = os.path.join(base_dir(), "settings.ini")
    p.read(path, encoding="utf-8")
    mp3 = "alarm.mp3"
    test_mode = False
    test_interval = 20
    play_limit = 0
    if p.has_section("app"):
        mp3 = p.get("app", "mp3", fallback=mp3)
        test_mode = p.get(
            "app", "test_mode", fallback=str(test_mode)
        ).strip().lower() in ("1", "true", "yes", "on")
        test_interval = max(
            1, min(59, p.getint("app", "test_interval_sec", fallback=test_interval))
        )
        play_limit = max(0, p.getint("app", "play_limit_sec", fallback=play_limit))
    return (
        os.path.join(base_dir(), mp3) if not os.path.isabs(mp3) else mp3,
        test_mode,
        test_interval,
        play_limit,
    )


def mci(cmd):
    buf = ctypes.create_unicode_buffer(512)
    err = ctypes.windll.winmm.mciSendStringW(cmd, buf, 512, 0)
    return err, buf.value


def ensure_ps_script():
    path = os.path.join(tempfile.gettempdir(), "chime_play.ps1")
    content = r"""param([string]$Path,[int]$Limit)
Add-Type -AssemblyName presentationCore
$u = [System.Uri]::new((Resolve-Path $Path))
$p = New-Object System.Windows.Media.MediaPlayer
$p.Open($u)
$p.Volume = 1.0
$e = New-Object System.Threading.AutoResetEvent($false)
$p.Add_MediaEnded({$e.Set()|Out-Null})
$p.Play()
if ($Limit -gt 0) { Start-Sleep -Seconds $Limit } else { $e.WaitOne()|Out-Null }
$p.Stop()"""
    try:
        need_write = True
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                if f.read() == content:
                    need_write = False
        if need_write:
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
    except:
        pass
    return path


def play_mp3_ps(path, limit_sec):
    ps1 = ensure_ps_script()
    creationflags = 0x08000000
    try:
        subprocess.Popen(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-WindowStyle",
                "Hidden",
                "-File",
                ps1,
                "-Path",
                path,
                "-Limit",
                str(int(limit_sec)),
            ],
            creationflags=creationflags,
        )
    except:
        pass


def play_mp3(path, limit_sec):
    if not os.path.isfile(path):
        return
    alias = f"a{uuid.uuid4().hex[:8]}"
    try:
        mci(f"close {alias}")
    except:
        pass
    started = False
    try:
        e, _ = mci(f'open "{path}" type mpegvideo alias {alias}')
        if e != 0:
            e, _ = mci(f'open "{path}" alias {alias}')
            if e != 0:
                play_mp3_ps(path, limit_sec)
                return
        e, _ = mci(f"play {alias}")
        if e != 0:
            mci(f"close {alias}")
            play_mp3_ps(path, limit_sec)
            return
        started = True
        t0 = time.time()
        time.sleep(0.4)
        _, pos = mci(f"status {alias} position")
        try:
            pos_ok = int((pos or "0").strip()) > 0
        except:
            pos_ok = False
        if not pos_ok:
            mci(f"stop {alias}")
            mci(f"close {alias}")
            e, _ = mci(f'open "{path}" alias {alias}')
            if e == 0:
                e, _ = mci(f"play {alias}")
                if e != 0:
                    mci(f"close {alias}")
                    play_mp3_ps(path, limit_sec)
                    return
                t0 = time.time()
                time.sleep(0.4)
                _, pos = mci(f"status {alias} position")
                try:
                    pos_ok = int((pos or "0").strip()) > 0
                except:
                    pos_ok = False
                if not pos_ok:
                    mci(f"stop {alias}")
                    mci(f"close {alias}")
                    play_mp3_ps(path, limit_sec)
                    return
        while True:
            if limit_sec > 0 and time.time() - t0 >= limit_sec:
                break
            _, mode = mci(f"status {alias} mode")
            mode = (mode or "").strip().lower()
            if mode in ("stopped", "not ready", ""):
                break
            time.sleep(0.2)
    finally:
        if started:
            try:
                mci(f"stop {alias}")
                mci(f"close {alias}")
            except:
                pass


def play_async(path, limit_sec):
    threading.Thread(target=play_mp3, args=(path, limit_sec), daemon=True).start()


def next_mark(now, test_mode, interval_sec):
    if test_mode:
        k = (now.second // interval_sec + 1) * interval_sec
        if k >= 60:
            return now.replace(second=0, microsecond=0) + timedelta(minutes=1)
        return now.replace(second=k, microsecond=0)
    return now.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)


def single_instance():
    h = ctypes.windll.kernel32.CreateMutexW(None, False, "hourly_chime_mutex_v6")
    if ctypes.windll.kernel32.GetLastError() == 183:
        return None
    return h


def main():
    if os.name != "nt":
        sys.exit(0)
    h = single_instance()
    if not h:
        sys.exit(0)
    mp3, test_mode, interval_sec, play_limit = load_settings()
    while True:
        now = datetime.now()
        target = next_mark(now, test_mode, interval_sec)
        wait = (target - now).total_seconds()
        if wait > 0:
            time.sleep(wait)
        play_async(mp3, play_limit)
        time.sleep(0.5)


if __name__ == "__main__":
    main()
