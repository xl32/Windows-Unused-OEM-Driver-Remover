import os
import sys
import locale
import ctypes
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from threading import Thread
import xml.etree.ElementTree as ET

# -------------------------
# Admin elevation (UAC)
# -------------------------
def is_admin() -> bool:
	try:
		return bool(ctypes.windll.shell32.IsUserAnAdmin())
	except Exception:
		return False

def relaunch_as_admin_or_exit() -> None:
	"""
	If not elevated, re-launch the same script with UAC prompt and exit current process.
	Works for:
	  - python script: python.exe script.py ...
	  - frozen exe (pyinstaller): app.exe ...
	"""
	if is_admin():
		return

	shell32 = ctypes.windll.shell32

	if getattr(sys, "frozen", False):
		# Frozen executable (pyinstaller, cx_Freeze, etc.)
		executable = sys.executable
		params = " ".join([f'"{a}"' for a in sys.argv[1:]])
	else:
		# Running as a .py via python.exe
		executable = sys.executable
		script = os.path.abspath(sys.argv[0])
		params = " ".join([f'"{script}"'] + [f'"{a}"' for a in sys.argv[1:]])

	# SW_SHOWNORMAL = 1
	rc = shell32.ShellExecuteW(None, "runas", executable, params, None, 1)
	if int(rc) <= 32:
		# User refused or error
		# Use a basic MessageBox (no Tk needed yet)
		ctypes.windll.user32.MessageBoxW(
			None,
			"Administrator privileges are required to manage drivers.\nElevation was cancelled.",
			"Admin rights required",
			0x10,  # MB_ICONERROR
		)
	sys.exit(0)

# -------------------------
# i18n (UI strings)
# -------------------------
STRINGS = {
	"en": {
		"title": "Unused OEM Drivers",
		"loading_title": "Loading",
		"loading_msg": "Loading drivers, please wait...",
		"note": "Hold CTRL and click to select multiple items.",
		"select_all": "Select All",
		"deselect_all": "Deselect All",
		"remove": "Remove Selected Drivers",
		"total": "Total Entries: {n}",
		"no_selection": "No drivers selected for removal.",
		"confirm_title": "Confirm Removal",
		"confirm_msg": "Are you sure you want to remove the following drivers?\n\n{items}",
		"success_title": "Success",
		"success_msg": "Driver {driver} removed successfully.",
		"error_title": "Error",
		"pnputil_failed": "Failed to run pnputil.\n\n{details}",
		"remove_failed": "Failed to remove driver {driver}:\n\n{details}",
		"parse_failed": "Failed to parse pnputil XML output.\n\n{details}",
		"columns": {
			"Type": "Type",
			"File Name": "File Name",
			"Original INF Name": "Original INF Name",
			"Provider Name": "Provider Name",
			"Class Name": "Class Name",
			"Device Count": "Device Count",
		},
		"type_oem": "OEM Driver",
		"lang_menu": "Language",
	},
	"ru": {
		"title": "Неиспользуемые OEM-драйверы",
		"loading_title": "Загрузка",
		"loading_msg": "Загрузка драйверов, подождите...",
		"note": "Удерживайте CTRL и кликайте для множественного выбора.",
		"select_all": "Выбрать все",
		"deselect_all": "Снять выбор",
		"remove": "Удалить выбранные драйверы",
		"total": "Всего: {n}",
		"no_selection": "Драйверы для удаления не выбраны.",
		"confirm_title": "Подтверждение удаления",
		"confirm_msg": "Удалить следующие драйверы?\n\n{items}",
		"success_title": "Готово",
		"success_msg": "Драйвер {driver} успешно удалён.",
		"error_title": "Ошибка",
		"pnputil_failed": "Не удалось запустить pnputil.\n\n{details}",
		"remove_failed": "Не удалось удалить {driver}:\n\n{details}",
		"parse_failed": "Не удалось разобрать XML-вывод pnputil.\n\n{details}",
		"columns": {
			"Type": "Тип",
			"File Name": "Файл",
			"Original INF Name": "Исходный INF",
			"Provider Name": "Поставщик",
			"Class Name": "Класс",
			"Device Count": "Кол-во устройств",
		},
		"type_oem": "OEM-драйвер",
		"lang_menu": "Язык",
	},
	"uk": {
		"title": "Невикористані OEM-драйвери",
		"loading_title": "Завантаження",
		"loading_msg": "Завантаження драйверів, зачекайте...",
		"note": "Утримуйте CTRL і клацайте для множинного вибору.",
		"select_all": "Вибрати все",
		"deselect_all": "Зняти вибір",
		"remove": "Видалити вибрані драйвери",
		"total": "Всього: {n}",
		"no_selection": "Немає вибраних драйверів для видалення.",
		"confirm_title": "Підтвердження видалення",
		"confirm_msg": "Видалити такі драйвери?\n\n{items}",
		"success_title": "Успішно",
		"success_msg": "Драйвер {driver} успішно видалено.",
		"error_title": "Помилка",
		"pnputil_failed": "Не вдалося запустити pnputil.\n\n{details}",
		"remove_failed": "Не вдалося видалити {driver}:\n\n{details}",
		"parse_failed": "Не вдалося розібрати XML-вивід pnputil.\n\n{details}",
		"columns": {
			"Type": "Тип",
			"File Name": "Файл",
			"Original INF Name": "Початковий INF",
			"Provider Name": "Постачальник",
			"Class Name": "Клас",
			"Device Count": "К-сть пристроїв",
		},
		"type_oem": "OEM-драйвер",
		"lang_menu": "Мова",
	},
}

def detect_language() -> str:
	# Prefer Windows UI language if available
	try:
		lang_id = ctypes.windll.kernel32.GetUserDefaultUILanguage()
		# Common mappings
		if lang_id == 0x0419:  # ru-RU
			return "ru"
		if lang_id == 0x0422:  # uk-UA
			return "uk"
		if lang_id == 0x0409:  # en-US
			return "en"
	except Exception:
		pass

	# Fallback to locale
	loc = (locale.getdefaultlocale() or ["en"])[0] or "en"
	loc = loc.lower()
	if loc.startswith("ru"):
		return "ru"
	if loc.startswith("uk"):
		return "uk"
	return "en"

# -------------------------
# pnputil (language-independent parsing via XML)
# -------------------------
def run_pnputil_xml() -> str:
	# Text output is localized; use XML format instead. :contentReference[oaicite:1]{index=1}
	# Also include /devices so we can find unused packages (devices count == 0). :contentReference[oaicite:2]{index=2}
	args = ["pnputil", "/enum-drivers", "/devices", "/format", "xml"]
	# CREATE_NO_WINDOW = 0x08000000
	creationflags = 0x08000000 if os.name == "nt" else 0
	p = subprocess.run(
		args,
		capture_output=True,
		text=True,
		shell=False,
		creationflags=creationflags,
	)
	if p.returncode != 0:
		raise RuntimeError(p.stderr.strip() or p.stdout.strip() or f"pnputil exited {p.returncode}")
	return p.stdout

def parse_unused_drivers_from_xml(xml_text: str):
	print(xml_text)
	"""
	Returns list of dicts:
	  File Name (oemXX.inf), Original INF Name, Provider Name, Class Name, Device Count
	"""
	try:
		root = ET.fromstring(xml_text)
	except Exception as e:
		raise RuntimeError(f"{e}")

	print("ROOT TAG:", root.tag)
	print("ROOT ATTR:", root.attrib)
	
	# show first 40 tags in the document (unique)
	tags = []
	for el in root.iter():
		tags.append(el.tag)
	uniq = []
	for t in tags:
		if t not in uniq:
			uniq.append(t)
	print("SAMPLE TAGS:", uniq[:40])

	# XML structure is documented by example in Microsoft guidance (PowerShell reads it as [xml]). :contentReference[oaicite:3]{index=3}
	drivers = []
	for drv in root.findall(".//Driver"):
		print(drv)
		# Tag names in XML are not localized.
		published = (drv.get("DriverName") or "").strip()
		original = (drv.findtext("OriginalName") or "").strip()
		provider = (drv.findtext("ProviderName") or "").strip()
		clsname = (drv.findtext("ClassName") or "").strip()
		print(published, original, provider, clsname)

		devices_node = drv.find("devices")
		# In Microsoft’s example they use: $_.devices.count :contentReference[oaicite:4]{index=4}
		# We'll derive count robustly:
		device_count = 0
		if devices_node is not None:
			# either attribute 'count' or nested list
			c_attr = devices_node.get("count")
			if c_attr is not None and c_attr.isdigit():
				device_count = int(c_attr)
			else:
				# count child nodes (best effort)
				device_count = len(list(devices_node))

		# "Unused" = not installed on any devices (count == 0). :contentReference[oaicite:5]{index=5}
		if published.lower().startswith("oem") and published.lower().endswith(".inf"):
			drivers.append(
				{
					"File Name": published or "N/A",
					"Original INF Name": original or "N/A",
					"Provider Name": provider or "N/A",
					"Class Name": clsname or "N/A",
					"Device Count": str(device_count),
				}
			)

	return drivers

# -------------------------
# Tk App
# -------------------------
class App:
	def __init__(self, window: tk.Tk, lang: str):
		self.window = window
		self.lang = lang

		self.total_label_var = tk.StringVar(value=self.t("total").format(n=0))

		self.window.title(self.t("title"))
		self.window.geometry("1100x650")

		self._build_menu()
		self._build_ui()

		self.refresh_table()

	def t(self, key: str) -> str:
		return STRINGS.get(self.lang, STRINGS["en"]).get(key, STRINGS["en"].get(key, key))

	def tc(self, col_key: str) -> str:
		cols = STRINGS.get(self.lang, STRINGS["en"])["columns"]
		return cols.get(col_key, col_key)

	def _build_menu(self):
		menubar = tk.Menu(self.window)
		lang_menu = tk.Menu(menubar, tearoff=0)
		for code, label in [("en", "English"), ("ru", "Русский"), ("uk", "Українська")]:
			lang_menu.add_command(label=label, command=lambda c=code: self.set_language(c))
		menubar.add_cascade(label=self.t("lang_menu"), menu=lang_menu)
		self.window.config(menu=menubar)

	def _build_ui(self):
		self.frame = tk.Frame(self.window)
		self.frame.pack(fill="both", expand=True, padx=10, pady=10)

		self.columns = ("Type", "File Name", "Original INF Name", "Provider Name", "Class Name", "Device Count")

		self.tree = ttk.Treeview(
			self.frame,
			columns=self.columns,
			show="headings",
			height=18,
			selectmode="extended",
		)

		scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.tree.yview)
		self.tree.configure(yscrollcommand=scrollbar.set)
		scrollbar.pack(side="right", fill="y")
		self.tree.pack(side="left", fill="both", expand=True)

		for col in self.columns:
			self.tree.heading(col, text=self.tc(col), command=lambda _col=col: self.sort_table(_col, False))
			self.tree.column(col, anchor="w", width=190)

		button_frame = tk.Frame(self.window)
		button_frame.pack(fill="x", pady=10)

		self.note_label = tk.Label(button_frame, text=self.t("note"), fg="blue")
		self.note_label.pack(side="left", padx=10)

		self.select_all_button = tk.Button(button_frame, text=self.t("select_all"), command=self.select_all)
		self.select_all_button.pack(side="left", padx=10)

		self.deselect_all_button = tk.Button(button_frame, text=self.t("deselect_all"), command=self.deselect_all)
		self.deselect_all_button.pack(side="left", padx=10)

		self.remove_button = tk.Button(button_frame, text=self.t("remove"), command=self.on_remove_selected)
		self.remove_button.pack(side="right", padx=10)

		self.total_label = tk.Label(button_frame, textvariable=self.total_label_var)
		self.total_label.pack(side="right", padx=10)

	def set_language(self, lang: str):
		self.lang = lang
		# Update static UI text
		self.window.title(self.t("title"))
		self.note_label.config(text=self.t("note"))
		self.select_all_button.config(text=self.t("select_all"))
		self.deselect_all_button.config(text=self.t("deselect_all"))
		self.remove_button.config(text=self.t("remove"))
		# Update headings
		for col in self.columns:
			self.tree.heading(col, text=self.tc(col))
		# Update total label format
		self.total_label_var.set(self.t("total").format(n=len(self.tree.get_children())))

		# Also update menu label (simple approach: rebuild menu)
		self._build_menu()

	def show_loading(self):
		splash = tk.Toplevel(self.window)
		splash.title(self.t("loading_title"))
		splash.geometry("360x120")
		splash.transient(self.window)
		splash.grab_set()
		label = tk.Label(splash, text=self.t("loading_msg"))
		label.pack(pady=15)
		bar = ttk.Progressbar(splash, mode="indeterminate")
		bar.pack(fill="x", padx=20, pady=10)
		bar.start(10)
		splash.update()
		return splash, bar

	def refresh_table(self):
		self.tree.delete(*self.tree.get_children())
		self.total_label_var.set(self.t("total").format(n=0))
		Thread(target=self.populate_table_async, daemon=True).start()

	def populate_table_async(self):
		splash, bar = None, None

		def _open_splash():
			nonlocal splash, bar
			splash, bar = self.show_loading()

		self.window.after(0, _open_splash)

		try:
			xml_text = run_pnputil_xml()
			drivers = parse_unused_drivers_from_xml(xml_text)
		except Exception as e:
			def _err():
				if splash:
					splash.destroy()
				messagebox.showerror(self.t("error_title"), self.t("pnputil_failed").format(details=str(e)))
			self.window.after(0, _err)
			return

		def _fill():
			if splash:
				splash.destroy()

			for d in drivers:
				self.tree.insert(
					"",
					"end",
					values=(
						STRINGS[self.lang].get("type_oem", "OEM Driver"),
						d["File Name"],
						d["Original INF Name"],
						d["Provider Name"],
						d["Class Name"],
						d["Device Count"],
					),
				)
			self.total_label_var.set(self.t("total").format(n=len(drivers)))

		self.window.after(0, _fill)

	def on_remove_selected(self):
		selected_items = self.tree.selection()
		if not selected_items:
			messagebox.showinfo(self.t("title"), self.t("no_selection"))
			return

		selected_drivers = [self.tree.item(item)["values"][1] for item in selected_items]
		items_str = ", ".join(selected_drivers)
		confirmation = messagebox.askyesno(
			self.t("confirm_title"),
			self.t("confirm_msg").format(items=items_str),
		)
		if not confirmation:
			return

		# Deleting drivers requires admin in most cases. :contentReference[oaicite:6]{index=6}
		# (We already elevated at startup; this is just a safety check.)
		if not is_admin():
			messagebox.showerror(self.t("error_title"), "Not elevated.")
			return

		creationflags = 0x08000000 if os.name == "nt" else 0

		for driver in selected_drivers:
			try:
				p = subprocess.run(
					["pnputil", "/delete-driver", driver, "/uninstall", "/force"],
					capture_output=True,
					text=True,
					shell=False,
					creationflags=creationflags,
				)
				if p.returncode == 0:
					messagebox.showinfo(self.t("success_title"), self.t("success_msg").format(driver=driver))
				else:
					details = (p.stderr or p.stdout or "").strip()
					messagebox.showerror(
						self.t("error_title"),
						self.t("remove_failed").format(driver=driver, details=details),
					)
			except Exception as e:
				messagebox.showerror(
					self.t("error_title"),
					self.t("remove_failed").format(driver=driver, details=str(e)),
				)

		self.refresh_table()

	def select_all(self):
		for item in self.tree.get_children():
			self.tree.selection_add(item)

	def deselect_all(self):
		for item in self.tree.get_children():
			self.tree.selection_remove(item)

	def sort_table(self, col, reverse: bool):
		data = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
		data.sort(reverse=reverse)
		for index, (_, k) in enumerate(data):
			self.tree.move(k, "", index)
		self.tree.heading(col, command=lambda: self.sort_table(col, not reverse))


def main():
	# Auto-elevate BEFORE any Tk UI is created
	relaunch_as_admin_or_exit()

	lang = detect_language()
	window = tk.Tk()
	App(window, lang)
	window.mainloop()

if __name__ == "__main__":
	main()
