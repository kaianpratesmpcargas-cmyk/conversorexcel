import threading
import time
import webbrowser

from web_app import app


def run_server():
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)


if __name__ == "__main__":
    thread = threading.Thread(target=run_server, daemon=True)
    thread.start()
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:5000")
    try:
        while thread.is_alive():
            time.sleep(1)
    except KeyboardInterrupt:
        pass
