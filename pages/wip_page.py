import os

import customtkinter as ctk
from PIL import Image

_ASSETS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")
_IMG_PATH = os.path.join(_ASSETS, "em_construcao.png")


class WIPPage(ctk.CTkFrame):
    page_key = ""

    def __init__(self, parent):
        super().__init__(parent, corner_radius=0)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._ctk_img = None
        self._pil_img = Image.open(_IMG_PATH) if os.path.exists(_IMG_PATH) else None
        self._label = ctk.CTkLabel(self, text="")
        self._label.grid(row=0, column=0)

        self.bind("<Configure>", self._on_resize)

    def on_show(self):
        self.after_idle(self._renderizar)

    def _on_resize(self, _event):
        self._renderizar()

    def _renderizar(self):
        if self._pil_img is None:
            self._label.configure(text="🚧  Em construção!", font=("Arial", 24, "bold"), text_color="gray55")
            return

        w = max(self.winfo_width() - 80, 100)
        h = max(self.winfo_height() - 80, 60)

        pil = self._pil_img
        ratio = min(w / pil.width, h / pil.height)
        size = (int(pil.width * ratio), int(pil.height * ratio))

        self._ctk_img = ctk.CTkImage(pil, size=size)
        self._label.configure(image=self._ctk_img, text="")
