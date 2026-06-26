from __future__ import annotations

from ..runtime import *  # noqa: F401,F403

class EditorCommandsWindow(Adw.ApplicationWindow):
    def __init__(self, parent: ProseWindow) -> None:
        super().__init__(application=parent.get_application(), title="Editor Commands")
        self._parent = parent
        self.set_default_size(900, 720)
        self.set_resizable(True)
        self._build_ui()

    def _build_ui(self) -> None:
        view = Adw.ToolbarView()
        header = Adw.HeaderBar()
        header.add_css_class("flat")
        header.set_title_widget(Adw.WindowTitle(title="Editor Commands", subtitle="Run or copy actions"))
        view.add_top_bar(header)

        page = Adw.PreferencesPage()
        intro = Adw.PreferencesGroup(
            title="How to use",
            description=(
                "Use Run to trigger actions inside the open Prose window. "
                "Use Copy Command to place the GApplication call on your clipboard."
            ),
        )
        page.add(intro)

        actions_group = Adw.PreferencesGroup(title="Editor Actions")
        page.add(actions_group)
        for title, action_name, param, desc, supports_profiles in _editor_command_items():
            row = self._build_action_row(title, action_name, param, desc, supports_profiles=supports_profiles)
            actions_group.add(row)

        scroller = Gtk.ScrolledWindow()
        scroller.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        scroller.set_child(page)
        view.set_content(scroller)
        self.set_content(view)

    def _build_action_row(
        self,
        title: str,
        action_name: str,
        param: str | None,
        desc: str,
        with_index: bool = False,
        supports_profiles: bool = False,
    ) -> Adw.ActionRow:
        row = Adw.ActionRow(title=title, subtitle=desc)
        row.set_activatable(False)

        suffix = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        spin = None
        profile_dropdown = None
        if with_index:
            spin = Gtk.SpinButton.new_with_range(0, 9999, 1)
            spin.set_value(0)
            spin.set_width_chars(4)
            suffix.append(spin)
        elif supports_profiles:
            profile_dropdown = Gtk.DropDown(
                model=Gtk.StringList.new(
                    [profile.display_name() for profile in self._parent._model_profiles]
                )
            )
            lookup_key = _profile_default_lookup_key_for_action_name(action_name)
            selected_profile_key = self._parent._editor_action_profile_defaults.get(lookup_key) if lookup_key else None
            selected_index = MODEL_PROFILE_IDS.index(selected_profile_key) if selected_profile_key in MODEL_PROFILE_IDS else 0
            profile_dropdown.set_selected(selected_index)
            suffix.append(profile_dropdown)

        run_btn = Gtk.Button(label="Run")
        run_btn.add_css_class("suggested-action")
        run_btn.add_css_class("flat")
        run_btn.connect("clicked", self._on_run_clicked, action_name, spin, profile_dropdown)
        suffix.append(run_btn)

        copy_btn = Gtk.Button(label="Copy Command")
        copy_btn.add_css_class("flat")
        copy_btn.add_css_class("link")
        copy_btn.connect("clicked", self._on_copy_clicked, action_name, spin, param, profile_dropdown)
        suffix.append(copy_btn)

        row.add_suffix(suffix)
        return row

    def _on_run_clicked(
        self,
        _button: Gtk.Button,
        action_name: str,
        spin: Gtk.SpinButton | None,
        profile_dropdown: Gtk.DropDown | None,
    ) -> None:
        app = self.get_application()
        if spin is None:
            if profile_dropdown is not None:
                selected = int(profile_dropdown.get_selected())
                nicknames = [profile.display_name() for profile in self._parent._model_profiles]
                if 0 <= selected < len(nicknames):
                    app.activate_action(action_name, GLib.Variant("s", nicknames[selected]))
                return
            app.activate_action(action_name, None)
            return
        value = int(spin.get_value())
        app.activate_action(action_name, GLib.Variant("i", value))

    def _on_copy_clicked(
        self,
        _button: Gtk.Button,
        action_name: str,
        spin: Gtk.SpinButton | None,
        param: str | None,
        profile_dropdown: Gtk.DropDown | None,
    ) -> None:
        if spin is not None:
            param = _format_action_param(GLib.Variant("i", int(spin.get_value())))
        elif profile_dropdown is not None:
            selected = int(profile_dropdown.get_selected())
            nicknames = [profile.display_name() for profile in self._parent._model_profiles]
            if 0 <= selected < len(nicknames):
                param = _format_action_param(GLib.Variant("s", nicknames[selected]))
        app = self.get_application()
        object_path = ACTION_OBJECT_PATH
        if isinstance(app, Gio.Application):
            app_path = app.get_dbus_object_path()
            if app_path:
                object_path = app_path
        command = _action_command(action_name, param, object_path)
        display = Gdk.Display.get_default()
        if display:
            display.get_clipboard().set(command)


