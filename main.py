# main.py
import logging
import ttkbootstrap as ttk
from gui import ProformaApp
from logger_config import setup_logging

if __name__ == "__main__":
    setup_logging()
    logger = logging.getLogger(__name__)
    logger.info("Запуск приложения")

    root = ttk.Window(themename='aqua')
    app = ProformaApp(root)
    root.mainloop()
