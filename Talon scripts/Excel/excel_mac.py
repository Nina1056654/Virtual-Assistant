from talon import Context, Module, actions, app, ui, registry, imgui, cron
from typing import Optional
from talon.mac import applescript

ctx = Context()
mod = Module()

ctx.matches = r"""
app: Microsoft Excel
"""


@ctx.action_class("edit")
class EditActions:
    def zoom_in():
        applescript.run(
            r"""
			tell application id "com.microsoft.Excel" to set front window's zoom to (front window's zoom) * 1.25
		"""
        )

    def zoom_out():
        applescript.run(
            r"""
			tell application id "com.microsoft.Excel" to set front window's zoom to (front window's zoom) / 1.25
		"""
        )

    def zoom_reset():
        applescript.run(
            r"""
			tell application id "com.microsoft.Excel" to set front window's zoom to 100
		"""
        )

    def line_insert_down():
        actions.key("enter")


def excel_app():
    return ui.apps(bundle="com.microsoft.Excel")[0]


def excel_window():
    return next(
        window for window in excel_app().windows() if window.doc or window.title
    )


@ctx.action_class("user")
class UserActions:
    def excel_save_as_format(format: str):
        actions.key("cmd-shift-s")
        window = excel_window()

        for attempt in range(5):
            try:
                sheet = window.children.find_one(AXRole="AXSheet", max_depth=0)
                break
            except ui.UIErr:
                actions.sleep("100ms")
        else:
            app.notify(body="Did not find save sheet as expected", title="Excel")
            return

        file_format_popup = sheet.children.find_one(
            AXRole="AXPopUpButton", AXDescription="File Format:"
        )
        file_format_popup.perform("AXPress")

        for attempt in range(5):
            try:
                file_format_menu = file_format_popup.children.find_one(
                    AXRole="AXMenu", max_depth=0
                )
                break
            except ui.UIErr:
                actions.sleep("100ms")
        else:
            app.notify(body="Did not find file format menu as expected", title="Excel")
            return

        file_format_item = file_format_menu.children.find_one(
            AXRole="AXMenuItem", AXTitle=format, max_depth=0
        )
        file_format_item.perform("AXPress")

    def find(text: str):
        actions.key("cmd-f")
        if text:
            actions.insert(text)

    def find_next():
        actions.key("cmd-g")

    def find_previous():
        actions.key("cmd-shift-g")

    def find_everywhere(text: str):
        actions.key("ctrl-f")
        if text:
            actions.insert(text)

    def find_toggle_match_by_case():
        pass  # could implement

    def find_toggle_match_by_word():
        pass

    def find_toggle_match_by_regex():
        pass

    def replace(text: str):
        actions.key("ctrl-h")
        if text:
            actions.insert(text)

    replace_everywhere = replace

    def replace_confirm():
        actions.key("cmd-r")

    def replace_confirm_all():
        actions.key("cmd-a")

    def select_previous_occurrence(text: str):
        actions.edit.find(text)
        actions.edit.find_previous()
        actions.key("esc")

    def select_next_occurrence(text: str):
        actions.edit.find(text)
        actions.edit.find_next()
        actions.key("esc")


@mod.capture(rule="<user.letters> <user.line_numbers>")
def xl_cell(m) -> str:
    "An excel cell"
    return f"{m.letters}{m.line_numbers}".upper()

@mod.action_class
class Actions:
    def excel_save_as_format(format: str):
        """Save Excel document with format"""

    def xl_select_cells(first: str, second: Optional[str] = None):
        """Select a cell, or range of cells"""
        target = f"{first}:{second}" if second else first
        actions.key("ctrl-g tab")
        actions.insert(target)
        actions.key("enter")
