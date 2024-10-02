# main.py
import logging
import ttkbootstrap as ttk
from gui import ProformaApp

if __name__ == "__main__":
    # Настройка логирования
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        handlers=[
            logging.FileHandler("app.log", encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    logger = logging.getLogger(__name__)
    logger.info("Запуск приложения")

    root = ttk.Window(themename='cosmo')
    app = ProformaApp(root)
    root.mainloop()
