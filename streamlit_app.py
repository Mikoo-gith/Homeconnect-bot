import asyncio
import json
import os
import threading
import time
import streamlit as st
from telegram.ext import Application

st.set_page_config(page_title="Bot Home Survey", page_icon="🤖")

for key, val in st.secrets.items():
    if key == "GOOGLE_CREDS_JSON":
        creds_data = json.loads(val) if isinstance(val, str) else dict(val)
        with open("credentials.json", "w", encoding="utf-8") as f:
            json.dump(creds_data, f)
    else:
        os.environ[key] = str(val)

_orig_run_polling = Application.run_polling


def _thread_safe_run_polling(self, *args, **kwargs):
    kwargs.setdefault("stop_signals", None)
    return _orig_run_polling(self, *args, **kwargs)


Application.run_polling = _thread_safe_run_polling

import homeconnect  # noqa: E402


@st.cache_resource
def start_bot_once():
    def _runner():
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        while True:
            try:
                homeconnect.main()
            except Exception as e:
                print(f"[BOT RESTART] Error: {e}")
                time.sleep(5)

    t = threading.Thread(target=_runner, daemon=True, name="TelegramBotThread")
    t.start()
    return t


bot_thread = start_bot_once()

st.title("🤖 Bot Tiket Home Survey")
st.success(f"Status Bot: {'Aktif (Running)' if bot_thread.is_alive() else 'Mati'}")