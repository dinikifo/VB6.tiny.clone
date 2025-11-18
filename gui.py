from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton,
    QMessageBox, QListWidget, QComboBox
)
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtCore import QUrl

from runtime import VBJsonRuntime
from interpreter import VBInterpreter
from runtime import VBContext


class VBTextBox:
    def __init__(self, widget: QLineEdit):
        self._widget = widget

    @property
    def Text(self):
        return self._widget.text()

    @Text.setter
    def Text(self, value):
        self._widget.setText(str(value))


class VBListBox:
    def __init__(self, widget: QListWidget):
        self._widget = widget

    @property
    def SelectedIndex(self):
        return self._widget.currentRow()

    @SelectedIndex.setter
    def SelectedIndex(self, value):
        self._widget.setCurrentRow(int(value))

    @property
    def SelectedText(self):
        item = self._widget.currentItem()
        return item.text() if item else ""


class VBComboBox:
    def __init__(self, widget: QComboBox):
        self._widget = widget

    @property
    def Text(self):
        return self._widget.currentText()

    @Text.setter
    def Text(self, value):
        self._widget.setCurrentText(str(value))

    @property
    def SelectedIndex(self):
        return self._widget.currentIndex()

    @SelectedIndex.setter
    def SelectedIndex(self, value):
        self._widget.setCurrentIndex(int(value))

    @property
    def SelectedText(self):
        return self._widget.currentText()


class VBWebBrowser:
    """Wrapper for QWebEngineView exposing Url and simple methods."""
    def __init__(self, widget: QWebEngineView):
        self._widget = widget

    @property
    def Url(self):
        url = self._widget.url()
        return url.toString() if url.isValid() else ""

    @Url.setter
    def Url(self, value):
        if value:
            self._widget.setUrl(QUrl(str(value)))

    def Navigate(self, url):
        self.Url = url

    def SetHtml(self, html):
        self._widget.setHtml(str(html))

    def EvalJs(self, script):
        # fire-and-forget JavaScript execution
        self._widget.page().runJavaScript(str(script))


class VBForm(QWidget):
    def __init__(self, form_def: dict, interpreter: VBInterpreter, ctx: VBContext):
        super().__init__()
        self.form_def = form_def
        self.interpreter = interpreter
        self.ctx = ctx
        self.control_bindings = {}  # name -> (widget, bind_path_rel, type)
        self.data_context = form_def.get("dataContext")

        self._build_ui()
        self.interpreter._msgbox = self._msgbox

        load_name = f"{form_def.get('name')}_Load"
        self.interpreter.call_sub(load_name)

        self._load_initial_data()

    def _build_ui(self):
        self.setWindowTitle(self.form_def.get("title", "VB Form"))

        size = self.form_def.get("size", [400, 200])
        if isinstance(size, (list, tuple)) and len(size) == 2:
            self.resize(size[0], size[1])

        layout = QVBoxLayout()
        self.setLayout(layout)

        controls = self.form_def.get("controls", [])
        for ctrl in controls:
            ctype = ctrl.get("type", "").lower()

            if ctype == "label":
                text = ctrl.get("text", "")
                w = QLabel(text)
                layout.addWidget(w)

            elif ctype == "textbox":
                name = ctrl.get("name")
                w = QLineEdit()
                if "placeholder" in ctrl:
                    w.setPlaceholderText(ctrl["placeholder"])
                layout.addWidget(w)

                if name:
                    vb_obj = VBTextBox(w)
                    self.ctx.set_var(name, vb_obj)

                bind = ctrl.get("bind")
                if bind and name:
                    self.control_bindings[name] = (w, bind, "textbox")

            elif ctype == "button":
                name = ctrl.get("name")
                text = ctrl.get("text", "Button")
                w = QPushButton(text)
                layout.addWidget(w)

                if name:
                    self.ctx.set_var(name, w)

                events = ctrl.get("events", {})
                for event_name, handler_name in events.items():
                    if event_name.lower() == "click":
                        w.clicked.connect(self._make_event_handler(handler_name))

            elif ctype == "listbox":
                name = ctrl.get("name")
                w = QListWidget()
                layout.addWidget(w)

                if name:
                    vb_obj = VBListBox(w)
                    self.ctx.set_var(name, vb_obj)

                bind = ctrl.get("bind")
                if bind and name:
                    self.control_bindings[name] = (w, bind, "listbox")

                events = ctrl.get("events", {})
                for event_name, handler_name in events.items():
                    if event_name.lower() in ("change", "selectedindexchanged", "selectionchanged"):
                        w.currentRowChanged.connect(self._make_event_handler(handler_name))

            elif ctype == "combobox":
                name = ctrl.get("name")
                w = QComboBox()
                layout.addWidget(w)

                if name:
                    vb_obj = VBComboBox(w)
                    self.ctx.set_var(name, vb_obj)

                bind = ctrl.get("bind")
                if bind and name:
                    self.control_bindings[name] = (w, bind, "combobox")

                events = ctrl.get("events", {})
                for event_name, handler_name in events.items():
                    if event_name.lower() in ("change", "selectedindexchanged", "selectionchanged"):
                        w.currentIndexChanged.connect(self._make_event_handler(handler_name))

            elif ctype == "webbrowser":
                name = ctrl.get("name")
                w = QWebEngineView()
                layout.addWidget(w)

                initial_url = ctrl.get("url")
                if initial_url:
                    w.setUrl(QUrl(initial_url))

                if name:
                    vb_obj = VBWebBrowser(w)
                    self.ctx.set_var(name, vb_obj)

                events = ctrl.get("events", {})
                for event_name, handler_name in events.items():
                    if event_name.lower() in ("loadfinished", "load_done"):
                        w.loadFinished.connect(self._make_event_handler(handler_name))

    def _make_event_handler(self, handler_name: str):
        def handler(*args, **kwargs):
            self.interpreter.call_sub(handler_name)
        return handler

    def _msgbox(self, text):
        QMessageBox.information(self, "MsgBox", str(text))

    def _load_initial_data(self):
        app_data = self.ctx.get_var("AppData")
        if app_data is None:
            return

        for name, (widget, rel_path, ctype) in self.control_bindings.items():
            full_path = rel_path
            if self.data_context:
                full_path = f"{self.data_context}.{rel_path}"
            try:
                val = VBJsonRuntime.json_get(app_data, full_path)
            except Exception:
                val = None

            if ctype == "textbox":
                widget.setText("" if val is None else str(val))
            elif ctype in ("listbox", "combobox"):
                if hasattr(widget, "clear"):
                    widget.clear()
                if isinstance(val, list):
                    for item in val:
                        widget.addItem(str(item))
