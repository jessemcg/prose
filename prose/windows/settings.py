from __future__ import annotations

from ..runtime import *  # noqa: F401,F403

class SettingsWindow(Adw.ApplicationWindow):
    def __init__(
        self,
        parent: ProseWindow,
        model_profiles: list[ModelProfile],
        proof_settings: ProofreadSettings,
        spelling_settings: SpellingStyleSettings,
        text_draft_spelling_prompt: str,
        citation_validator_settings: CitationValidatorSettings,
        improve1_settings: Improve1Settings,
        improve2_settings: Improve2Settings,
        combine_cites_settings: CombineCitesSettings,
        thesaurus_settings: ThesaurusSettings,
        reference_settings: ReferenceSettings,
        ask_settings: AskSettings,
        shorten_settings: ShortenSettings,
        introduction_settings: IntroductionSettings,
        introduction_reply_settings: IntroductionSettings,
        conclusion_settings: ConclusionSettings,
        concl_no_issues_settings: ConclNoIssuesSettings,
        topic_sentence_settings: TopicSentenceSettings,
        concl_section_settings: ConclSectionSettings,
        translate_settings: TranslateSettings,
        shared_style_rules: str,
        editor_action_profile_defaults: dict[str, str | None],
        editor_pinned_action_ids: list[str],
        text_draft_pinned_action_ids: list[str],
        text_draft_template_dir: Path | None,
        text_draft_external_actions: list[TextDraftExternalAction],
        libreoffice_python_path: Path | None,
        concordance_file_path: Path | None,
        editor_source_file: Path | None,
        on_source_change: Callable[[Path | None], None],
        on_copy_normal_profile: Callable[[], tuple[bool, str]],
        on_save: Callable[
            [
                list[ModelProfile],
                ProofreadSettings,
                SpellingStyleSettings,
                str,
                CitationValidatorSettings,
                Improve1Settings,
                Improve2Settings,
                CombineCitesSettings,
                ThesaurusSettings,
                ReferenceSettings,
                AskSettings,
                ShortenSettings,
                IntroductionSettings,
                IntroductionSettings,
                ConclusionSettings,
                ConclNoIssuesSettings,
                TopicSentenceSettings,
                ConclSectionSettings,
                TranslateSettings,
                str,
                dict[str, str | None],
                list[str],
                list[str],
                Path | None,
                list[TextDraftExternalAction],
                Path | None,
                Path | None,
            ],
            None,
        ],
    ) -> None:
        super().__init__(application=parent.get_application(), title="Settings")
        self._parent_window = parent
        self._on_save = on_save
        self._on_source_change = on_source_change
        self._on_copy_normal_profile = on_copy_normal_profile
        self._model_profiles = model_profiles
        self._proof_settings = proof_settings
        self._spelling_settings = spelling_settings
        self._text_draft_spelling_prompt = text_draft_spelling_prompt.strip() or DEFAULT_TEXT_DRAFT_SPELLINGSTYLE_PROMPT
        self._citation_validator_settings = citation_validator_settings
        self._improve1_settings = improve1_settings
        self._improve2_settings = improve2_settings
        self._combine_cites_settings = combine_cites_settings
        self._thesaurus_settings = thesaurus_settings
        self._reference_settings = reference_settings
        self._ask_settings = ask_settings
        self._shorten_settings = shorten_settings
        self._introduction_settings = introduction_settings
        self._introduction_reply_settings = introduction_reply_settings
        self._conclusion_settings = conclusion_settings
        self._concl_no_issues_settings = concl_no_issues_settings
        self._topic_sentence_settings = topic_sentence_settings
        self._concl_section_settings = concl_section_settings
        self._translate_settings = translate_settings
        self._shared_style_rules_text = str(shared_style_rules or "").strip()
        self._editor_action_profile_defaults = _sanitize_editor_action_profile_defaults(editor_action_profile_defaults)
        self._editor_action_order = _ordered_editor_quick_action_keys(editor_pinned_action_ids)
        pinned_action_set = set(_sanitize_editor_pinned_actions(editor_pinned_action_ids))
        self._quick_action_enabled = {
            key: key in pinned_action_set
            for key in self._editor_action_order
        }
        self._text_draft_action_order = _ordered_text_draft_quick_action_keys(text_draft_pinned_action_ids)
        text_draft_pinned_action_set = set(_sanitize_text_draft_pinned_actions(text_draft_pinned_action_ids))
        self._text_draft_quick_action_enabled = {
            key: key in text_draft_pinned_action_set
            for key in self._text_draft_action_order
        }
        self._text_draft_template_dir = (
            text_draft_template_dir.expanduser().resolve(strict=False) if text_draft_template_dir else None
        )
        self._text_draft_external_actions = list(text_draft_external_actions)
        self._libreoffice_python_path = libreoffice_python_path
        self._concordance_file_path = concordance_file_path
        self._editor_source_file = editor_source_file
        self._proof_suggestions_json_path = proof_settings.suggestions_json_path
        self._text_draft_spelling_prompt_buffer: Gtk.TextBuffer | None = None
        self._shared_style_rules_buffer: Gtk.TextBuffer | None = None
        self._model_profile_editors: dict[str, ModelProfileEditorWidgets] = {}
        self._prompt_editors: dict[str, PromptEditorWidgets] = {}
        self._choices_profile_dropdowns: dict[str, Gtk.DropDown] = {}
        self._prompt_row_keys: dict[Gtk.ListBoxRow, str] = {}
        self._source_row_guard = False
        self._text_draft_template_dir_row_guard = False
        self._text_draft_external_actions_box: Gtk.Box | None = None
        self._text_draft_external_action_widgets: list[TextDraftExternalActionEditorWidgets] = []
        self._libreoffice_path_row_guard = False
        self._concordance_path_row_guard = False
        self._proof_suggestions_json_path_row_guard = False
        self._quick_action_toggle_guard = False
        self._text_draft_quick_action_toggle_guard = False
        self.set_default_size(900, 720)
        self.set_resizable(True)
        self._build_ui()

    def _build_ui(self) -> None:
        view = Adw.ToolbarView()
        header = Adw.HeaderBar()
        header.add_css_class("flat")
        header.set_title_widget(Adw.WindowTitle(title="Settings"))
        view.add_top_bar(header)

        box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        box.set_margin_top(18)
        box.set_margin_bottom(12)
        box.set_margin_start(18)
        box.set_margin_end(18)

        split = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=12)
        split.set_hexpand(True)
        split.set_vexpand(False)

        prompt_list = Gtk.ListBox()
        prompt_list.set_selection_mode(Gtk.SelectionMode.SINGLE)
        prompt_list.add_css_class("navigation-sidebar")
        prompt_list.connect("row-selected", self._on_prompt_row_selected)

        prompt_list_scroller = Gtk.ScrolledWindow()
        prompt_list_scroller.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        prompt_list_scroller.set_min_content_width(220)
        prompt_list_scroller.set_size_request(220, 720)
        prompt_list_scroller.set_child(prompt_list)

        prompt_stack = Gtk.Stack()
        prompt_stack.set_hexpand(True)
        prompt_stack.set_vexpand(True)
        prompt_stack.set_transition_type(Gtk.StackTransitionType.SLIDE_LEFT_RIGHT)
        self._prompt_stack = prompt_stack
        self._prompt_list = prompt_list

        first_row: Gtk.ListBoxRow | None = None

        source_row = Gtk.ListBoxRow()
        source_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        source_row_box.set_margin_top(8)
        source_row_box.set_margin_bottom(8)
        source_row_box.set_margin_start(12)
        source_row_box.set_margin_end(12)
        source_row_box.append(Gtk.Label(label="Source Text", xalign=0))
        source_row.set_child(source_row_box)
        prompt_list.append(source_row)
        self._prompt_row_keys[source_row] = "source-text"
        prompt_stack.add_named(self._build_source_text_page(), "source-text")
        first_row = source_row

        external_action_row = Gtk.ListBoxRow()
        external_action_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        external_action_row_box.set_margin_top(8)
        external_action_row_box.set_margin_bottom(8)
        external_action_row_box.set_margin_start(12)
        external_action_row_box.set_margin_end(12)
        external_action_row_box.append(Gtk.Label(label="Text Draft External Action", xalign=0))
        external_action_row.set_child(external_action_row_box)
        prompt_list.append(external_action_row)
        self._prompt_row_keys[external_action_row] = "text-draft-external-action"
        prompt_stack.add_named(self._build_text_draft_external_action_page(), "text-draft-external-action")

        profiles_row = Gtk.ListBoxRow()
        profiles_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        profiles_row_box.set_margin_top(8)
        profiles_row_box.set_margin_bottom(8)
        profiles_row_box.set_margin_start(12)
        profiles_row_box.set_margin_end(12)
        profiles_row_box.append(Gtk.Label(label="Model Profiles", xalign=0))
        profiles_row.set_child(profiles_row_box)
        prompt_list.append(profiles_row)
        self._prompt_row_keys[profiles_row] = "model-profiles"
        prompt_stack.add_named(self._build_model_profiles_page(), "model-profiles")

        choices_models_row = Gtk.ListBoxRow()
        choices_models_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        choices_models_row_box.set_margin_top(8)
        choices_models_row_box.set_margin_bottom(8)
        choices_models_row_box.set_margin_start(12)
        choices_models_row_box.set_margin_end(12)
        choices_models_row_box.append(Gtk.Label(label="Choices Models", xalign=0))
        choices_models_row.set_child(choices_models_row_box)
        prompt_list.append(choices_models_row)
        self._prompt_row_keys[choices_models_row] = "choices-models"
        prompt_stack.add_named(self._build_choices_models_page(), "choices-models")

        style_rules_row = Gtk.ListBoxRow()
        style_rules_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        style_rules_row_box.set_margin_top(8)
        style_rules_row_box.set_margin_bottom(8)
        style_rules_row_box.set_margin_start(12)
        style_rules_row_box.set_margin_end(12)
        style_rules_row_box.append(Gtk.Label(label="Style Rules", xalign=0))
        style_rules_row.set_child(style_rules_row_box)
        prompt_list.append(style_rules_row)
        self._prompt_row_keys[style_rules_row] = "style-rules"
        prompt_stack.add_named(self._build_style_rules_page(), "style-rules")

        prompt_definitions = [
            ("proof", "Proof Reading", self._proof_settings, DEFAULT_PROMPT),
            ("spelling", "SpellingStyle", self._spelling_settings, DEFAULT_SPELLINGSTYLE_PROMPT),
            (
                "citation-validator",
                "Citation Number Validator",
                self._citation_validator_settings,
                DEFAULT_CITATION_NUMBER_VALIDATOR_PROMPT,
            ),
            ("thesaurus", "Thesaurus", self._thesaurus_settings, DEFAULT_THESAURUS_PROMPT),
            ("reference", "Reference", self._reference_settings, DEFAULT_REFERENCE_PROMPT),
            ("ask", "Ask Field", self._ask_settings, DEFAULT_ASK_PROMPT),
            ("improve-generated", "Improve Generated", self._improve1_settings, DEFAULT_IMPROVE_PROMPT),
            ("rephrase-generated", "Rephrase Generated", self._improve2_settings, DEFAULT_IMPROVE2_PROMPT),
            ("combine", "Combine Cites", self._combine_cites_settings, DEFAULT_COMBINE_CITES_PROMPT),
            ("shorten", "Shorten Selected", self._shorten_settings, DEFAULT_SHORTEN_PROMPT),
            ("intro", "Introduction", self._introduction_settings, DEFAULT_INTRO_PROMPT),
            ("intro-reply", "Introduction Reply", self._introduction_reply_settings, DEFAULT_INTRO_REPLY_PROMPT),
            ("conclusion", "Conclusion", self._conclusion_settings, DEFAULT_CONCLUSION_PROMPT),
            ("concl-no-issues", "Conclusion No Issues", self._concl_no_issues_settings, DEFAULT_CONCL_NO_ISSUES_PROMPT),
            ("topic-sentence", "Topic Sentence", self._topic_sentence_settings, DEFAULT_TOPIC_SENTENCE_PROMPT),
            ("concl-section", "Section Conclusion", self._concl_section_settings, DEFAULT_CONCL_SECTION_PROMPT),
            ("translate", "Translate", self._translate_settings, DEFAULT_TRANSLATE_PROMPT),
        ]
        for key, title, settings, default_prompt in prompt_definitions:
            row = Gtk.ListBoxRow()
            row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
            row_box.set_margin_top(8)
            row_box.set_margin_bottom(8)
            row_box.set_margin_start(12)
            row_box.set_margin_end(12)
            label = Gtk.Label(label=title, xalign=0)
            row_box.append(label)
            row.set_child(row_box)
            prompt_list.append(row)
            self._prompt_row_keys[row] = key
            if first_row is None:
                first_row = row

            page = self._build_prompt_page(key, title, settings, default_prompt)
            prompt_stack.add_named(page, key)

        text_draft_spelling_row = Gtk.ListBoxRow()
        text_draft_spelling_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        text_draft_spelling_row_box.set_margin_top(8)
        text_draft_spelling_row_box.set_margin_bottom(8)
        text_draft_spelling_row_box.set_margin_start(12)
        text_draft_spelling_row_box.set_margin_end(12)
        text_draft_spelling_row_box.append(Gtk.Label(label="Text Draft SpellingStyle", xalign=0))
        text_draft_spelling_row.set_child(text_draft_spelling_row_box)
        prompt_list.append(text_draft_spelling_row)
        self._prompt_row_keys[text_draft_spelling_row] = "text-draft-spelling"
        prompt_stack.add_named(self._build_text_draft_spellingstyle_page(), "text-draft-spelling")

        quick_actions_row = Gtk.ListBoxRow()
        quick_actions_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        quick_actions_row_box.set_margin_top(8)
        quick_actions_row_box.set_margin_bottom(8)
        quick_actions_row_box.set_margin_start(12)
        quick_actions_row_box.set_margin_end(12)
        quick_actions_row_box.append(Gtk.Label(label="Quick Actions", xalign=0))
        quick_actions_row.set_child(quick_actions_row_box)
        prompt_list.append(quick_actions_row)
        self._prompt_row_keys[quick_actions_row] = "quick-actions"
        prompt_stack.add_named(self._build_quick_actions_page(), "quick-actions")

        text_draft_quick_actions_row = Gtk.ListBoxRow()
        text_draft_quick_actions_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        text_draft_quick_actions_row_box.set_margin_top(8)
        text_draft_quick_actions_row_box.set_margin_bottom(8)
        text_draft_quick_actions_row_box.set_margin_start(12)
        text_draft_quick_actions_row_box.set_margin_end(12)
        text_draft_quick_actions_row_box.append(Gtk.Label(label="Text Draft Quick Actions", xalign=0))
        text_draft_quick_actions_row.set_child(text_draft_quick_actions_row_box)
        prompt_list.append(text_draft_quick_actions_row)
        self._prompt_row_keys[text_draft_quick_actions_row] = "text-draft-quick-actions"
        prompt_stack.add_named(self._build_text_draft_quick_actions_page(), "text-draft-quick-actions")

        editor_commands_row = Gtk.ListBoxRow()
        editor_commands_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        editor_commands_row_box.set_margin_top(8)
        editor_commands_row_box.set_margin_bottom(8)
        editor_commands_row_box.set_margin_start(12)
        editor_commands_row_box.set_margin_end(12)
        editor_commands_row_box.append(Gtk.Label(label="Editor Commands", xalign=0))
        editor_commands_row.set_child(editor_commands_row_box)
        prompt_list.append(editor_commands_row)
        self._prompt_row_keys[editor_commands_row] = "editor-commands"
        prompt_stack.add_named(self._build_editor_commands_page(), "editor-commands")

        text_draft_commands_row = Gtk.ListBoxRow()
        text_draft_commands_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        text_draft_commands_row_box.set_margin_top(8)
        text_draft_commands_row_box.set_margin_bottom(8)
        text_draft_commands_row_box.set_margin_start(12)
        text_draft_commands_row_box.set_margin_end(12)
        text_draft_commands_row_box.append(Gtk.Label(label="Text Draft Commands", xalign=0))
        text_draft_commands_row.set_child(text_draft_commands_row_box)
        prompt_list.append(text_draft_commands_row)
        self._prompt_row_keys[text_draft_commands_row] = "text-draft-commands"
        prompt_stack.add_named(self._build_text_draft_commands_page(), "text-draft-commands")

        if first_row is not None:
            prompt_list.select_row(first_row)
            prompt_stack.set_visible_child_name(self._prompt_row_keys[first_row])

        split.append(prompt_list_scroller)
        split.append(prompt_stack)
        box.append(split)

        content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        scrolled = Gtk.ScrolledWindow()
        scrolled.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        scrolled.set_hexpand(True)
        scrolled.set_vexpand(True)
        scrolled.set_child(box)
        content.append(scrolled)

        buttons = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        buttons.set_margin_top(6)
        buttons.set_margin_bottom(12)
        buttons.set_margin_start(12)
        buttons.set_margin_end(12)
        buttons.set_halign(Gtk.Align.END)
        close_btn = Gtk.Button(label="Close")
        close_btn.add_css_class("flat")
        close_btn.connect("clicked", self._on_close_clicked)
        buttons.append(close_btn)
        save_btn = Gtk.Button(label="Save Settings")
        save_btn.add_css_class("suggested-action")
        save_btn.add_css_class("flat")
        save_btn.set_action_name("app.save-settings")
        buttons.append(save_btn)
        content.append(buttons)

        view.set_content(content)
        self.set_content(view)

    def set_source_file(self, path: Path | None) -> None:
        self._set_editor_source_file(path, notify=False)

    def _build_source_text_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Source Text", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        source_group = Adw.PreferencesGroup(title="Source Text")
        source_group.add_css_class("list-stack")
        source_group.set_hexpand(True)
        page_box.append(source_group)

        source_row = Adw.EntryRow(title="Source text file")
        source_row.set_hexpand(True)
        if self._editor_source_file:
            source_row.set_text(str(self._editor_source_file))
        source_row.connect("changed", self._on_source_row_changed)
        choose_btn = Gtk.Button(label="Choose")
        choose_btn.add_css_class("flat")
        choose_btn.connect("clicked", self._on_choose_source_file)
        source_row.add_suffix(choose_btn)
        source_group.add(source_row)
        self._source_row = source_row

        text_draft_template_row, text_draft_template_entry = self._build_path_setting_row(
            title="Text Draft template directory",
            value=str(self._text_draft_template_dir or ""),
            info_text=(
                "Optional. Prose treats immediate subfolders here as template categories, and root-level .txt files "
                "appear under Uncategorized. "
                "Add leading lines like [[email: Label <address@example.com>]] to create copy buttons."
            ),
            on_changed=self._on_text_draft_template_dir_row_changed,
            on_choose=self._on_choose_text_draft_template_dir,
            on_clear=self._on_clear_text_draft_template_dir,
        )
        source_group.add(text_draft_template_row)
        self._text_draft_template_dir_row = text_draft_template_entry

        proof_suggestions_row, proof_suggestions_entry = self._build_path_setting_row(
            title="Proofreading suggestions JSON",
            value=str(self._proof_suggestions_json_path or ""),
            info_text=(
                "Optional. When this .json file exists, Prose auto-loads it when the Proof Reader tab is shown. "
                "Clear this path to disable auto-load."
            ),
            on_changed=self._on_proof_suggestions_json_path_row_changed,
            on_choose=self._on_choose_proof_suggestions_json_path,
            on_clear=self._on_clear_proof_suggestions_json_path,
        )
        source_group.add(proof_suggestions_row)
        self._proof_suggestions_json_path_row = proof_suggestions_entry

        libreoffice_row, libreoffice_entry = self._build_path_setting_row(
            title="LibreOffice Python path",
            value=str(self._libreoffice_python_path or ""),
            info_text="Required. Point this at LibreOffice's Python bridge directory, usually .../libreoffice/program.",
            on_changed=self._on_libreoffice_path_row_changed,
            on_choose=self._on_choose_libreoffice_path,
            on_clear=self._on_clear_libreoffice_path,
        )
        source_group.add(libreoffice_row)
        self._libreoffice_path_row = libreoffice_entry

        concordance_row, concordance_entry = self._build_path_setting_row(
            title="Concordance file",
            value=str(self._concordance_file_path or ""),
            info_text=(
                "Optional. Used by Add Case when appending citations to the concordance file "
                "and by the Text Draft case picker."
            ),
            on_changed=self._on_concordance_path_row_changed,
            on_choose=self._on_choose_concordance_file,
            on_clear=self._on_clear_concordance_file,
        )
        source_group.add(concordance_row)
        self._concordance_path_row = concordance_entry

        profile_row = Adw.ActionRow(
            title="LibreOffice Profile Import",
            subtitle=(
                f"Copy {_normal_libreoffice_profile_path()} into {LIBREOFFICE_PROFILE}.\n"
                "Overwrites the Prose LibreOffice profile after confirmation. Close any Prose Writer window first."
            ),
        )
        profile_row.set_subtitle_lines(4)
        profile_copy_btn = Gtk.Button(label="Copy Normal Profile to Prose")
        profile_copy_btn.add_css_class("flat")
        profile_copy_btn.connect("clicked", self._on_copy_normal_profile_clicked)
        profile_row.add_suffix(profile_copy_btn)
        source_group.add(profile_row)

        return page_box

    def _build_text_draft_external_action_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Text Draft External Action", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        actions_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        actions_box.set_hexpand(True)
        page_box.append(actions_box)
        self._text_draft_external_actions_box = actions_box
        self._rebuild_text_draft_external_action_editor_rows()

        add_button = Gtk.Button(label="Add Action", icon_name="list-add-symbolic")
        add_button.add_css_class("flat")
        add_button.add_css_class("suggested-action")
        add_button.set_halign(Gtk.Align.START)
        add_button.connect("clicked", self._on_add_text_draft_external_action_clicked)
        page_box.append(add_button)

        return page_box

    def _clear_settings_box(self, box: Gtk.Box) -> None:
        child = box.get_first_child()
        while child is not None:
            next_child = child.get_next_sibling()
            box.remove(child)
            child = next_child

    def _rebuild_text_draft_external_action_editor_rows(self) -> None:
        box = self._text_draft_external_actions_box
        if box is None:
            return
        self._clear_settings_box(box)
        self._text_draft_external_action_widgets = []
        if not self._text_draft_external_actions:
            empty_label = Gtk.Label(label="No external Text Draft actions are configured.", xalign=0)
            empty_label.add_css_class("dim-label")
            box.append(empty_label)
            return

        for index, action in enumerate(self._text_draft_external_actions):
            group = Adw.PreferencesGroup(title=f"Action {index + 1}")
            group.add_css_class("list-stack")
            group.set_hexpand(True)
            box.append(group)

            enabled_row = Adw.SwitchRow(
                title="Enable action",
                subtitle="Show this action as a Text Draft button.",
            )
            enabled_row.set_active(action.enabled)
            group.add(enabled_row)

            label_row = Adw.EntryRow(title="Button label")
            label_row.set_text(action.label or "External")
            group.add(label_row)

            icon_row = Adw.EntryRow(title="Icon name")
            icon_row.set_text(action.icon_name or DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON)
            group.add(icon_row)

            tooltip_row = Adw.EntryRow(title="Tooltip")
            tooltip_row.set_text(action.tooltip)
            group.add(tooltip_row)

            success_message_row = Adw.EntryRow(title="Success message")
            success_message_row.set_text(action.success_message)
            group.add(success_message_row)

            command_row = Adw.EntryRow(title="Command")
            command_row.set_text(shlex.join(action.command))
            command_row.set_show_apply_button(False)
            command_row.set_tooltip_text(
                f"Use {TEXT_DRAFT_EXTERNAL_ACTION_DRAFT_FILE_TOKEN} where the Draft temp-file path should go."
            )
            group.add(command_row)

            cwd_row = Adw.EntryRow(title="Working directory")
            cwd_row.set_text(str(action.cwd or ""))
            cwd_row.set_show_apply_button(False)
            group.add(cwd_row)

            codex_reasoning_dropdown = None
            env_values = dict(action.env)
            if self._parent_window._is_text_draft_codex_action(action):
                env_values.pop("CODEX_REASONING_EFFORT", None)

                codex_reasoning_row = Adw.ActionRow(
                    title="Codex reasoning effort",
                    subtitle="Reasoning level used when this Codex action starts.",
                )
                codex_reasoning_row.set_activatable(False)
                codex_reasoning_dropdown = Gtk.DropDown(
                    model=Gtk.StringList.new(list(TEXT_DRAFT_CODEX_REASONING_EFFORTS))
                )
                codex_reasoning_dropdown.set_selected(
                    TEXT_DRAFT_CODEX_REASONING_EFFORTS.index(
                        _sanitize_text_draft_codex_reasoning_effort(action.codex_reasoning_effort)
                    )
                )
                codex_reasoning_row.add_suffix(codex_reasoning_dropdown)
                group.add(codex_reasoning_row)

            env_buffer = Gtk.TextBuffer()
            env_buffer.set_text(_format_text_draft_external_action_env(env_values))
            env_row = Adw.PreferencesRow()
            env_row.set_selectable(False)
            env_row.set_activatable(False)
            env_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
            env_box.set_margin_top(10)
            env_box.set_margin_bottom(10)
            env_box.set_margin_start(12)
            env_box.set_margin_end(12)

            env_title = Gtk.Label(label="Environment", xalign=0)
            env_title.add_css_class("heading")
            env_box.append(env_title)

            env_hint = Gtk.Label(
                label="Optional. Enter one KEY=value pair per line.",
                xalign=0,
            )
            env_hint.add_css_class("dim-label")
            env_hint.set_wrap(True)
            env_box.append(env_hint)

            env_view = Gtk.TextView.new_with_buffer(env_buffer)
            env_view.set_monospace(True)
            env_view.set_wrap_mode(Gtk.WrapMode.NONE)
            env_view.set_hexpand(True)
            env_view.set_top_margin(8)
            env_view.set_bottom_margin(8)
            env_view.set_left_margin(8)
            env_view.set_right_margin(8)

            env_scroller = Gtk.ScrolledWindow()
            env_scroller.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
            env_scroller.set_hexpand(True)
            env_scroller.set_min_content_height(82)
            env_scroller.set_max_content_height(150)
            env_scroller.set_child(env_view)
            env_box.append(env_scroller)

            env_row.set_child(env_box)
            group.add(env_row)

            remove_row = Adw.ActionRow(title="Remove action")
            remove_button = Gtk.Button(icon_name="user-trash-symbolic")
            remove_button.add_css_class("flat")
            remove_button.add_css_class("destructive-action")
            remove_button.set_tooltip_text("Remove this external Text Draft action.")
            remove_button.connect(
                "clicked",
                lambda _button, action_index=index: self._on_remove_text_draft_external_action_clicked(
                    action_index,
                ),
            )
            remove_row.add_suffix(remove_button)
            group.add(remove_row)

            self._text_draft_external_action_widgets.append(
                TextDraftExternalActionEditorWidgets(
                    enabled_row=enabled_row,
                    label_row=label_row,
                    icon_row=icon_row,
                    tooltip_row=tooltip_row,
                    success_message_row=success_message_row,
                    command_row=command_row,
                    cwd_row=cwd_row,
                    codex_reasoning_dropdown=codex_reasoning_dropdown,
                    env_buffer=env_buffer,
                )
            )

    def _collect_text_draft_external_actions_from_settings(self) -> list[TextDraftExternalAction] | None:
        actions: list[TextDraftExternalAction] = []
        for index, widgets in enumerate(self._text_draft_external_action_widgets):
            command: list[str] = []
            command_text = widgets.command_row.get_text().strip()
            if command_text:
                try:
                    command = shlex.split(command_text)
                except ValueError as exc:
                    self._parent_window._show_toast(f"External action {index + 1} command is invalid: {exc}")
                    return None
            env, env_error = _parse_text_draft_external_action_env(self._prompt_text(widgets.env_buffer))
            if env_error is not None or env is None:
                self._parent_window._show_toast(f"External action {index + 1} environment {env_error}")
                return None
            codex_reasoning_effort = DEFAULT_TEXT_DRAFT_CODEX_REASONING_EFFORT
            if widgets.codex_reasoning_dropdown is not None:
                env.pop("CODEX_REASONING_EFFORT", None)
                selected_reasoning = int(widgets.codex_reasoning_dropdown.get_selected())
                if 0 <= selected_reasoning < len(TEXT_DRAFT_CODEX_REASONING_EFFORTS):
                    codex_reasoning_effort = TEXT_DRAFT_CODEX_REASONING_EFFORTS[selected_reasoning]
            cwd_text = widgets.cwd_row.get_text().strip()
            actions.append(
                TextDraftExternalAction(
                    enabled=widgets.enabled_row.get_active(),
                    label=widgets.label_row.get_text().strip() or "External",
                    command=command,
                    cwd=Path(cwd_text).expanduser().resolve(strict=False) if cwd_text else None,
                    env=env,
                    icon_name=widgets.icon_row.get_text().strip() or DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON,
                    tooltip=widgets.tooltip_row.get_text().strip(),
                    success_message=widgets.success_message_row.get_text().strip(),
                    codex_reasoning_effort=codex_reasoning_effort,
                )
            )
        return actions

    def _on_add_text_draft_external_action_clicked(self, _button: Gtk.Button) -> None:
        actions = self._collect_text_draft_external_actions_from_settings()
        if actions is None:
            return
        actions.append(
            TextDraftExternalAction(
                enabled=True,
                label="External",
                command=[],
                cwd=None,
                env={},
                icon_name=DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON,
                tooltip="",
                success_message="",
            )
        )
        self._text_draft_external_actions = actions
        self._rebuild_text_draft_external_action_editor_rows()

    def _on_remove_text_draft_external_action_clicked(self, index: int) -> None:
        actions = self._collect_text_draft_external_actions_from_settings()
        if actions is None:
            return
        if 0 <= index < len(actions):
            actions.pop(index)
        self._text_draft_external_actions = actions
        self._rebuild_text_draft_external_action_editor_rows()

    def _profile_dropdown_model(self, include_unset: bool = False) -> Gtk.StringList:
        labels = [profile.display_name() for profile in self._model_profiles]
        if include_unset:
            labels = [UNSET_PROFILE_LABEL, *labels]
        return Gtk.StringList.new(labels)

    def _build_model_profiles_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_vexpand(True)

        title_label = Gtk.Label(label="Model Profiles", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        info_label = Gtk.Label(
            label=(
                "Set up up to four shared model profiles here. Proof Reading now keeps its own credentials, "
                "while eligible editor actions reuse these shared profiles and keep only their own prompts."
            ),
            xalign=0,
        )
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(info_label)

        for profile in self._model_profiles:
            group = Adw.PreferencesGroup(title=profile.display_name())
            group.add_css_class("list-stack")

            nickname_row = Adw.EntryRow(title="Nickname")
            nickname_row.set_text(profile.display_name())
            group.add(nickname_row)

            abbreviation_row = Adw.EntryRow(title="Abbreviation (optional)")
            abbreviation_row.set_text(profile.abbreviation)
            group.add(abbreviation_row)

            api_url_row = Adw.EntryRow(title="API URL")
            api_url_row.set_text(profile.api_url)
            group.add(api_url_row)

            model_row = Adw.EntryRow(title="Model ID (optional)")
            model_row.set_text(profile.model_id)
            group.add(model_row)

            api_key_row = Adw.PasswordEntryRow(title="API Key")
            api_key_row.set_text(profile.api_key)
            group.add(api_key_row)

            disable_reasoning_row = Adw.SwitchRow(title="Disable reasoning")
            disable_reasoning_row.set_active(bool(profile.disable_reasoning))
            group.add(disable_reasoning_row)

            priority_service_tier_row = Adw.SwitchRow(title="Priority")
            priority_service_tier_row.set_active(bool(profile.priority_service_tier))
            group.add(priority_service_tier_row)

            self._model_profile_editors[profile.key] = ModelProfileEditorWidgets(
                nickname_row=nickname_row,
                abbreviation_row=abbreviation_row,
                api_url_row=api_url_row,
                model_row=model_row,
                api_key_row=api_key_row,
                disable_reasoning_row=disable_reasoning_row,
                priority_service_tier_row=priority_service_tier_row,
            )
            page_box.append(group)

        page = Gtk.ScrolledWindow()
        page.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        page.set_hexpand(True)
        page.set_vexpand(True)
        page.set_child(page_box)
        return page

    def _build_choices_models_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Choices Models", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        info_label = Gtk.Label(
            label=(
                "Generated Choices and Selected Choices use these same three model assignments. "
                "Improve uses the Improve Generated prompt; Rephrase 1 and Rephrase 2 use the Rephrase Generated prompt. "
                "Draft introduction and conclusion Choices, including Section Conclusion Choices, use the action prompt "
                "for all three assignments."
            ),
            xalign=0,
        )
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(info_label)

        group = Adw.PreferencesGroup(title="Choices Model Profiles")
        group.add_css_class("list-stack")
        group.set_hexpand(True)
        page_box.append(group)

        self._choices_profile_dropdowns = {}
        rows = (
            ("choices-improve", "Improve", "Default Choices model assignment."),
            ("choices-rephrase-1", "Rephrase 1", "First Choices variation model assignment."),
            ("choices-rephrase-2", "Rephrase 2", "Second Choices variation model assignment."),
        )
        for profile_key, title, subtitle in rows:
            row = Adw.ActionRow(title=title, subtitle=subtitle)
            row.set_activatable(False)
            dropdown = Gtk.DropDown(model=self._profile_dropdown_model(include_unset=True))
            selected_profile_key = self._editor_action_profile_defaults.get(profile_key)
            selected_index = (
                MODEL_PROFILE_IDS.index(selected_profile_key) + 1
                if selected_profile_key in MODEL_PROFILE_IDS
                else 0
            )
            dropdown.set_selected(selected_index)
            row.add_suffix(dropdown)
            group.add(row)
            self._choices_profile_dropdowns[profile_key] = dropdown

        return page_box

    def _build_quick_actions_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Quick Actions", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        info_label = Gtk.Label(
            label=(
                f"Pin up to {MAX_PINNED_EDITOR_ACTIONS} actions for the Editor toolbar. "
                "Use the arrows to change their order."
            ),
            xalign=0,
        )
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(info_label)

        quick_actions_group = Adw.PreferencesGroup(title="Pinned Editor Actions")
        quick_actions_group.add_css_class("list-stack")
        quick_actions_group.set_hexpand(True)
        page_box.append(quick_actions_group)

        quick_actions_row = Adw.PreferencesRow()
        quick_actions_row.set_selectable(False)
        quick_actions_row.set_activatable(False)

        quick_actions_content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
        quick_actions_content.set_margin_top(10)
        quick_actions_content.set_margin_bottom(10)
        quick_actions_content.set_margin_start(12)
        quick_actions_content.set_margin_end(12)

        quick_actions_hint = Gtk.Label(xalign=0)
        quick_actions_hint.add_css_class("caption")
        quick_actions_hint.add_css_class("dim-label")
        quick_actions_hint.set_wrap(True)
        quick_actions_hint.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        quick_actions_content.append(quick_actions_hint)
        self._quick_actions_hint = quick_actions_hint

        quick_actions_list = Gtk.ListBox()
        quick_actions_list.set_selection_mode(Gtk.SelectionMode.NONE)
        quick_actions_list.add_css_class("boxed-list")
        quick_actions_content.append(quick_actions_list)
        self._quick_actions_list = quick_actions_list

        quick_actions_row.set_child(quick_actions_content)
        quick_actions_group.add(quick_actions_row)
        self._rebuild_quick_action_rows()
        return page_box

    def _build_text_draft_quick_actions_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Text Draft Quick Actions", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        info_label = Gtk.Label(
            label=(
                f"Pin up to {MAX_PINNED_EDITOR_ACTIONS} actions for the Text Draft toolbar. "
                "Unpinned actions stay under More."
            ),
            xalign=0,
        )
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(info_label)

        quick_actions_group = Adw.PreferencesGroup(title="Pinned Text Draft Actions")
        quick_actions_group.add_css_class("list-stack")
        quick_actions_group.set_hexpand(True)
        page_box.append(quick_actions_group)

        quick_actions_row = Adw.PreferencesRow()
        quick_actions_row.set_selectable(False)
        quick_actions_row.set_activatable(False)

        quick_actions_content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
        quick_actions_content.set_margin_top(10)
        quick_actions_content.set_margin_bottom(10)
        quick_actions_content.set_margin_start(12)
        quick_actions_content.set_margin_end(12)

        quick_actions_hint = Gtk.Label(xalign=0)
        quick_actions_hint.add_css_class("caption")
        quick_actions_hint.add_css_class("dim-label")
        quick_actions_hint.set_wrap(True)
        quick_actions_hint.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        quick_actions_content.append(quick_actions_hint)
        self._text_draft_quick_actions_hint = quick_actions_hint

        quick_actions_list = Gtk.ListBox()
        quick_actions_list.set_selection_mode(Gtk.SelectionMode.NONE)
        quick_actions_list.add_css_class("boxed-list")
        quick_actions_content.append(quick_actions_list)
        self._text_draft_quick_actions_list = quick_actions_list

        quick_actions_row.set_child(quick_actions_content)
        quick_actions_group.add(quick_actions_row)
        self._rebuild_text_draft_quick_action_rows()
        return page_box

    def _build_editor_commands_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Editor Commands", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        info_label = Gtk.Label(
            label=(
                "Use Run to trigger actions inside the open Prose window. "
                "Use Copy Command to place the GApplication call on your clipboard."
            ),
            xalign=0,
        )
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(info_label)

        actions_group = Adw.PreferencesGroup(title="Editor Actions")
        actions_group.add_css_class("list-stack")
        page_box.append(actions_group)

        for title, action_name, param, desc, supports_profiles in _editor_command_items():
            row = self._build_settings_editor_command_row(
                title,
                action_name,
                param,
                desc,
                supports_profiles=supports_profiles,
            )
            actions_group.add(row)

        return page_box

    def _build_text_draft_commands_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Text Draft Commands", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        info_label = Gtk.Label(
            label=(
                "Use Run to trigger Text Draft actions inside the open Prose window. "
                "Use Copy Command to place the GApplication call on your clipboard for use in other programs."
            ),
            xalign=0,
        )
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(info_label)

        actions_group = Adw.PreferencesGroup(title="Text Draft Actions")
        actions_group.add_css_class("list-stack")
        page_box.append(actions_group)

        for title, action_name, param, desc, supports_profiles in _text_draft_command_items():
            row = self._build_settings_editor_command_row(
                title,
                action_name,
                param,
                desc,
                supports_profiles=supports_profiles,
            )
            actions_group.add(row)

        return page_box

    def _build_settings_editor_command_row(
        self,
        title: str,
        action_name: str,
        param: str | None,
        desc: str,
        supports_profiles: bool = False,
    ) -> Adw.ActionRow:
        row = Adw.ActionRow(title=title, subtitle=desc)
        row.set_activatable(False)

        suffix = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        profile_dropdown = None
        if supports_profiles:
            profile_dropdown = Gtk.DropDown(model=self._profile_dropdown_model())
            lookup_key = _profile_default_lookup_key_for_action_name(action_name)
            selected_profile_key = self._editor_action_profile_defaults.get(lookup_key) if lookup_key else None
            selected_index = MODEL_PROFILE_IDS.index(selected_profile_key) if selected_profile_key in MODEL_PROFILE_IDS else 0
            profile_dropdown.set_selected(selected_index)
            suffix.append(profile_dropdown)

        run_btn = Gtk.Button(label="Run")
        run_btn.add_css_class("suggested-action")
        run_btn.add_css_class("flat")
        run_btn.connect("clicked", self._on_settings_editor_command_run_clicked, action_name, profile_dropdown)
        suffix.append(run_btn)

        copy_btn = Gtk.Button(label="Copy Command")
        copy_btn.add_css_class("flat")
        copy_btn.add_css_class("link")
        copy_btn.connect("clicked", self._on_settings_editor_command_copy_clicked, action_name, param, profile_dropdown)
        suffix.append(copy_btn)

        row.add_suffix(suffix)
        return row

    def _on_settings_editor_command_run_clicked(
        self,
        _button: Gtk.Button,
        action_name: str,
        profile_dropdown: Gtk.DropDown | None,
    ) -> None:
        app = self.get_application()
        if profile_dropdown is not None:
            selected = int(profile_dropdown.get_selected())
            nicknames = [profile.display_name() for profile in self._model_profiles]
            if 0 <= selected < len(nicknames):
                app.activate_action(action_name, GLib.Variant("s", nicknames[selected]))
            return
        app.activate_action(action_name, None)

    def _on_settings_editor_command_copy_clicked(
        self,
        _button: Gtk.Button,
        action_name: str,
        param: str | None,
        profile_dropdown: Gtk.DropDown | None,
    ) -> None:
        if profile_dropdown is not None:
            selected = int(profile_dropdown.get_selected())
            nicknames = [profile.display_name() for profile in self._model_profiles]
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

    def _on_choose_source_file(self, _button: Gtk.Button) -> None:
        dialog = Gtk.FileDialog(title="Choose a source text file")
        dialog.open(self, None, self._on_source_file_chosen)

    def _on_source_file_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.open_finish(result)
            path = Path(file.get_path() or "")
        except Exception:
            return
        if not path:
            return
        self._set_editor_source_file(path, notify=True)

    def _on_source_row_changed(self, row: Adw.EntryRow) -> None:
        if self._source_row_guard:
            return
        raw = row.get_text().strip()
        if not raw:
            self._set_editor_source_file(None, notify=True)
            return
        self._set_editor_source_file(Path(raw), notify=True)

    def _build_path_setting_row(
        self,
        title: str,
        value: str,
        info_text: str,
        on_changed: Callable[[Gtk.Editable], None],
        on_choose: Callable[[Gtk.Button], None],
        on_clear: Callable[[Gtk.Button], None],
    ) -> tuple[Adw.PreferencesRow, Gtk.Entry]:
        row = Adw.PreferencesRow()
        row.set_selectable(False)
        row.set_activatable(False)

        content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
        content.set_margin_top(10)
        content.set_margin_bottom(10)
        content.set_margin_start(12)
        content.set_margin_end(12)

        title_label = Gtk.Label(label=title, xalign=0)
        title_label.add_css_class("caption")
        title_label.add_css_class("dim-label")
        content.append(title_label)

        entry_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        entry = Gtk.Entry()
        entry.set_hexpand(True)
        entry.set_text(value)
        entry.connect("changed", on_changed)
        entry_box.append(entry)

        choose_btn = Gtk.Button(label="Choose")
        choose_btn.add_css_class("flat")
        choose_btn.connect("clicked", on_choose)
        entry_box.append(choose_btn)

        clear_btn = Gtk.Button(label="Clear")
        clear_btn.add_css_class("flat")
        clear_btn.connect("clicked", on_clear)
        entry_box.append(clear_btn)

        content.append(entry_box)

        info_label = Gtk.Label(label=info_text, xalign=0)
        info_label.add_css_class("caption")
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        content.append(info_label)

        row.set_child(content)
        return row, entry

    def _set_editor_source_file(self, path: Path | None, notify: bool) -> None:
        self._editor_source_file = path.expanduser().resolve(strict=False) if path else None
        self._source_row_guard = True
        self._source_row.set_text(str(self._editor_source_file or ""))
        self._source_row_guard = False
        if notify:
            self._on_source_change(self._editor_source_file)

    def _on_choose_text_draft_template_dir(self, _button: Gtk.Button) -> None:
        dialog = Gtk.FileDialog(title="Choose Text Draft template directory")
        dialog.select_folder(self, None, self._on_text_draft_template_dir_chosen)

    def _on_text_draft_template_dir_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.select_folder_finish(result)
            path = Path(file.get_path() or "")
        except Exception:
            return
        if not path:
            return
        self._set_text_draft_template_dir(path)

    def _on_clear_text_draft_template_dir(self, _button: Gtk.Button) -> None:
        self._set_text_draft_template_dir(None)

    def _on_text_draft_template_dir_row_changed(self, row: Gtk.Editable) -> None:
        if self._text_draft_template_dir_row_guard:
            return
        raw = row.get_text().strip()
        if not raw:
            self._set_text_draft_template_dir(None)
            return
        self._set_text_draft_template_dir(Path(raw))

    def _set_text_draft_template_dir(self, path: Path | None) -> None:
        self._text_draft_template_dir = path.expanduser().resolve(strict=False) if path else None
        self._text_draft_template_dir_row_guard = True
        self._text_draft_template_dir_row.set_text(str(self._text_draft_template_dir or ""))
        self._text_draft_template_dir_row_guard = False

    def _on_choose_proof_suggestions_json_path(self, _button: Gtk.Button) -> None:
        dialog = Gtk.FileDialog(title="Choose proofreading suggestions JSON")
        dialog.open(self, None, self._on_proof_suggestions_json_path_chosen)

    def _on_proof_suggestions_json_path_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.open_finish(result)
            path = Path(file.get_path() or "")
        except Exception:
            return
        if not path:
            return
        self._set_proof_suggestions_json_path(path)

    def _on_clear_proof_suggestions_json_path(self, _button: Gtk.Button) -> None:
        self._set_proof_suggestions_json_path(None)

    def _on_proof_suggestions_json_path_row_changed(self, row: Gtk.Editable) -> None:
        if self._proof_suggestions_json_path_row_guard:
            return
        raw = row.get_text().strip()
        if not raw:
            self._set_proof_suggestions_json_path(None)
            return
        self._set_proof_suggestions_json_path(Path(raw))

    def _set_proof_suggestions_json_path(self, path: Path | None) -> None:
        self._proof_suggestions_json_path = path.expanduser().resolve(strict=False) if path else None
        self._proof_suggestions_json_path_row_guard = True
        self._proof_suggestions_json_path_row.set_text(str(self._proof_suggestions_json_path or ""))
        self._proof_suggestions_json_path_row_guard = False

    def _on_choose_libreoffice_path(self, _button: Gtk.Button) -> None:
        dialog = Gtk.FileDialog(title="Choose LibreOffice Python directory")
        dialog.select_folder(self, None, self._on_libreoffice_path_chosen)

    def _on_libreoffice_path_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.select_folder_finish(result)
            path = Path(file.get_path() or "")
        except Exception:
            return
        if not path:
            return
        self._set_libreoffice_python_path(path)

    def _on_clear_libreoffice_path(self, _button: Gtk.Button) -> None:
        self._set_libreoffice_python_path(None)

    def _on_libreoffice_path_row_changed(self, row: Gtk.Editable) -> None:
        if self._libreoffice_path_row_guard:
            return
        raw = row.get_text().strip()
        if not raw:
            self._set_libreoffice_python_path(None)
            return
        self._set_libreoffice_python_path(Path(raw))

    def _set_libreoffice_python_path(self, path: Path | None) -> None:
        self._libreoffice_python_path = path.expanduser().resolve(strict=False) if path else None
        self._libreoffice_path_row_guard = True
        self._libreoffice_path_row.set_text(str(self._libreoffice_python_path or ""))
        self._libreoffice_path_row_guard = False

    def _on_choose_concordance_file(self, _button: Gtk.Button) -> None:
        dialog = Gtk.FileDialog(title="Choose concordance file")
        dialog.open(self, None, self._on_concordance_file_chosen)

    def _on_concordance_file_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.open_finish(result)
            path = Path(file.get_path() or "")
        except Exception:
            return
        if not path:
            return
        self._set_concordance_file_path(path)

    def _on_clear_concordance_file(self, _button: Gtk.Button) -> None:
        self._set_concordance_file_path(None)

    def _on_concordance_path_row_changed(self, row: Gtk.Editable) -> None:
        if self._concordance_path_row_guard:
            return
        raw = row.get_text().strip()
        if not raw:
            self._set_concordance_file_path(None)
            return
        self._set_concordance_file_path(Path(raw))

    def _set_concordance_file_path(self, path: Path | None) -> None:
        self._concordance_file_path = path.expanduser().resolve(strict=False) if path else None
        self._concordance_path_row_guard = True
        self._concordance_path_row.set_text(str(self._concordance_file_path or ""))
        self._concordance_path_row_guard = False

    def _on_copy_normal_profile_clicked(self, _button: Gtk.Button) -> None:
        source = _normal_libreoffice_profile_path()
        target = LIBREOFFICE_PROFILE
        dialog = Adw.MessageDialog.new(
            self,
            "Copy normal LibreOffice profile?",
            (
                f"This replaces the Prose LibreOffice profile.\n\n"
                f"Source: {source}\n"
                f"Target: {target}\n\n"
                "Any Prose-specific LibreOffice settings in the target profile will be lost."
            ),
        )
        dialog.add_response("cancel", "Cancel")
        dialog.add_response("copy", "Copy Profile")
        dialog.set_response_appearance("copy", Adw.ResponseAppearance.DESTRUCTIVE)
        dialog.set_default_response("cancel")
        dialog.set_close_response("cancel")
        dialog.connect("response", self._on_copy_normal_profile_response)
        dialog.present()

    def _on_copy_normal_profile_response(self, _dialog: Adw.MessageDialog, response: str) -> None:
        if response != "copy":
            return
        ok, message = self._on_copy_normal_profile()
        self._parent_window._show_toast(message)
        if ok:
            self._parent_window._status_label.set_label("LibreOffice profile imported. Launch Writer to use it.")

    def _clear_list_box(self, list_box: Gtk.ListBox) -> None:
        child = list_box.get_first_child()
        while child is not None:
            next_child = child.get_next_sibling()
            list_box.remove(child)
            child = next_child

    def _current_editor_pinned_actions(self) -> list[str]:
        return [key for key in self._editor_action_order if self._quick_action_enabled.get(key, False)]

    def _update_quick_actions_hint(self) -> None:
        count = len(self._current_editor_pinned_actions())
        self._quick_actions_hint.set_label(
            f"{count} of {MAX_PINNED_EDITOR_ACTIONS} pinned. Unpinned actions stay under More."
        )

    def _rebuild_quick_action_rows(self) -> None:
        self._clear_list_box(self._quick_actions_list)
        self._update_quick_actions_hint()

        last_index = len(self._editor_action_order) - 1
        for index, key in enumerate(self._editor_action_order):
            definition = EDITOR_QUICK_ACTION_BY_KEY[key]
            row = Adw.ActionRow(title=definition.title, subtitle=definition.description)
            row.set_activatable(False)

            toggle = Gtk.CheckButton()
            toggle.set_valign(Gtk.Align.CENTER)
            toggle.set_active(self._quick_action_enabled.get(key, False))
            toggle.set_tooltip_text("Show this action in the Editor quick-actions row")
            toggle.connect("toggled", self._on_quick_action_toggled, key)
            row.add_prefix(toggle)

            up_btn = Gtk.Button(icon_name="go-up-symbolic")
            up_btn.add_css_class("flat")
            up_btn.set_tooltip_text("Move up")
            up_btn.set_sensitive(index > 0)
            up_btn.connect("clicked", self._on_move_quick_action_clicked, key, -1)
            row.add_suffix(up_btn)

            down_btn = Gtk.Button(icon_name="go-down-symbolic")
            down_btn.add_css_class("flat")
            down_btn.set_tooltip_text("Move down")
            down_btn.set_sensitive(index < last_index)
            down_btn.connect("clicked", self._on_move_quick_action_clicked, key, 1)
            row.add_suffix(down_btn)

            self._quick_actions_list.append(row)

    def _current_text_draft_pinned_actions(self) -> list[str]:
        return [key for key in self._text_draft_action_order if self._text_draft_quick_action_enabled.get(key, False)]

    def _update_text_draft_quick_actions_hint(self) -> None:
        count = len(self._current_text_draft_pinned_actions())
        self._text_draft_quick_actions_hint.set_label(
            f"{count} of {MAX_PINNED_EDITOR_ACTIONS} pinned. Unpinned actions stay under More."
        )

    def _rebuild_text_draft_quick_action_rows(self) -> None:
        self._clear_list_box(self._text_draft_quick_actions_list)
        self._update_text_draft_quick_actions_hint()

        last_index = len(self._text_draft_action_order) - 1
        for index, key in enumerate(self._text_draft_action_order):
            definition = TEXT_DRAFT_QUICK_ACTION_BY_KEY[key]
            row = Adw.ActionRow(title=definition.title, subtitle=definition.description)
            row.set_activatable(False)

            toggle = Gtk.CheckButton()
            toggle.set_valign(Gtk.Align.CENTER)
            toggle.set_active(self._text_draft_quick_action_enabled.get(key, False))
            toggle.set_tooltip_text("Show this action in the Text Draft quick-actions row")
            toggle.connect("toggled", self._on_text_draft_quick_action_toggled, key)
            row.add_prefix(toggle)

            up_btn = Gtk.Button(icon_name="go-up-symbolic")
            up_btn.add_css_class("flat")
            up_btn.set_tooltip_text("Move up")
            up_btn.set_sensitive(index > 0)
            up_btn.connect("clicked", self._on_move_text_draft_quick_action_clicked, key, -1)
            row.add_suffix(up_btn)

            down_btn = Gtk.Button(icon_name="go-down-symbolic")
            down_btn.add_css_class("flat")
            down_btn.set_tooltip_text("Move down")
            down_btn.set_sensitive(index < last_index)
            down_btn.connect("clicked", self._on_move_text_draft_quick_action_clicked, key, 1)
            row.add_suffix(down_btn)

            self._text_draft_quick_actions_list.append(row)

    def _on_quick_action_toggled(self, button: Gtk.CheckButton, key: str) -> None:
        if self._quick_action_toggle_guard:
            return

        active = button.get_active()
        was_active = self._quick_action_enabled.get(key, False)
        if active == was_active:
            return

        if active and len(self._current_editor_pinned_actions()) >= MAX_PINNED_EDITOR_ACTIONS:
            self._quick_action_toggle_guard = True
            button.set_active(False)
            self._quick_action_toggle_guard = False
            self._parent_window._show_toast(f"Pin up to {MAX_PINNED_EDITOR_ACTIONS} quick actions.")
            return

        self._quick_action_enabled[key] = active
        self._update_quick_actions_hint()

    def _on_move_quick_action_clicked(self, _button: Gtk.Button, key: str, delta: int) -> None:
        index = self._editor_action_order.index(key)
        target_index = index + delta
        if target_index < 0 or target_index >= len(self._editor_action_order):
            return
        self._editor_action_order[index], self._editor_action_order[target_index] = (
            self._editor_action_order[target_index],
            self._editor_action_order[index],
        )
        self._rebuild_quick_action_rows()

    def _on_text_draft_quick_action_toggled(self, button: Gtk.CheckButton, key: str) -> None:
        if self._text_draft_quick_action_toggle_guard:
            return

        active = button.get_active()
        was_active = self._text_draft_quick_action_enabled.get(key, False)
        if active == was_active:
            return

        if active and len(self._current_text_draft_pinned_actions()) >= MAX_PINNED_EDITOR_ACTIONS:
            self._text_draft_quick_action_toggle_guard = True
            button.set_active(False)
            self._text_draft_quick_action_toggle_guard = False
            self._parent_window._show_toast(f"Pin up to {MAX_PINNED_EDITOR_ACTIONS} Text Draft quick actions.")
            return

        self._text_draft_quick_action_enabled[key] = active
        self._update_text_draft_quick_actions_hint()

    def _on_move_text_draft_quick_action_clicked(self, _button: Gtk.Button, key: str, delta: int) -> None:
        index = self._text_draft_action_order.index(key)
        target_index = index + delta
        if target_index < 0 or target_index >= len(self._text_draft_action_order):
            return
        self._text_draft_action_order[index], self._text_draft_action_order[target_index] = (
            self._text_draft_action_order[target_index],
            self._text_draft_action_order[index],
        )
        self._rebuild_text_draft_quick_action_rows()

    def trigger_save(self) -> None:
        self._on_save_clicked(None)

    def _on_close_clicked(self, _button: Gtk.Button) -> None:
        self.close()

    def _on_save_clicked(self, _button: Gtk.Button) -> None:
        model_profiles: list[ModelProfile] = []
        for profile_key in MODEL_PROFILE_IDS:
            widgets = self._model_profile_editors.get(profile_key)
            if widgets is None:
                continue
            model_profiles.append(
                ModelProfile(
                    key=profile_key,
                    nickname=widgets.nickname_row.get_text().strip() or DEFAULT_MODEL_PROFILE_NICKNAMES[profile_key],
                    abbreviation=widgets.abbreviation_row.get_text().strip(),
                    api_url=widgets.api_url_row.get_text().strip(),
                    model_id=widgets.model_row.get_text().strip(),
                    api_key=widgets.api_key_row.get_text().strip(),
                    disable_reasoning=widgets.disable_reasoning_row.get_active(),
                    priority_service_tier=widgets.priority_service_tier_row.get_active(),
                )
            )

        proof_widgets = self._prompt_editors.get("proof")
        spelling_widgets = self._prompt_editors.get("spelling")
        citation_validator_widgets = self._prompt_editors.get("citation-validator")
        thesaurus_widgets = self._prompt_editors.get("thesaurus")
        reference_widgets = self._prompt_editors.get("reference")
        ask_widgets = self._prompt_editors.get("ask")
        improve_widgets = self._prompt_editors.get("improve-generated")
        rephrase_widgets = self._prompt_editors.get("rephrase-generated")
        combine_widgets = self._prompt_editors.get("combine")
        shorten_widgets = self._prompt_editors.get("shorten")
        intro_widgets = self._prompt_editors.get("intro")
        intro_reply_widgets = self._prompt_editors.get("intro-reply")
        conclusion_widgets = self._prompt_editors.get("conclusion")
        concl_no_issues_widgets = self._prompt_editors.get("concl-no-issues")
        topic_sentence_widgets = self._prompt_editors.get("topic-sentence")
        concl_section_widgets = self._prompt_editors.get("concl-section")
        translate_widgets = self._prompt_editors.get("translate")
        text_draft_spelling_prompt_buffer = self._text_draft_spelling_prompt_buffer
        if not all(
            (
                proof_widgets,
                spelling_widgets,
                citation_validator_widgets,
                thesaurus_widgets,
                reference_widgets,
                ask_widgets,
                improve_widgets,
                rephrase_widgets,
                combine_widgets,
                shorten_widgets,
                intro_widgets,
                intro_reply_widgets,
                conclusion_widgets,
                concl_no_issues_widgets,
                topic_sentence_widgets,
                concl_section_widgets,
                translate_widgets,
                text_draft_spelling_prompt_buffer,
            )
        ):
            return

        proof_prompt_text = self._prompt_text(proof_widgets.prompt_buffer)
        spelling_prompt_text = self._prompt_text(spelling_widgets.prompt_buffer)
        citation_validator_prompt_text = self._prompt_text(citation_validator_widgets.prompt_buffer)
        text_draft_spelling_prompt_text = self._prompt_text(text_draft_spelling_prompt_buffer)
        thesaurus_prompt_text = self._prompt_text(thesaurus_widgets.prompt_buffer)
        reference_prompt_text = self._prompt_text(reference_widgets.prompt_buffer)
        ask_prompt_text = self._prompt_text(ask_widgets.prompt_buffer)
        improve_prompt_text = self._prompt_text(improve_widgets.prompt_buffer)
        rephrase_prompt_text = self._prompt_text(rephrase_widgets.prompt_buffer)
        combine_prompt_text = self._prompt_text(combine_widgets.prompt_buffer)
        shorten_prompt_text = self._prompt_text(shorten_widgets.prompt_buffer)
        intro_prompt_text = self._prompt_text(intro_widgets.prompt_buffer)
        intro_reply_prompt_text = self._prompt_text(intro_reply_widgets.prompt_buffer)
        conclusion_prompt_text = self._prompt_text(conclusion_widgets.prompt_buffer)
        concl_no_issues_prompt_text = self._prompt_text(concl_no_issues_widgets.prompt_buffer)
        topic_sentence_prompt_text = self._prompt_text(topic_sentence_widgets.prompt_buffer)
        concl_section_prompt_text = self._prompt_text(concl_section_widgets.prompt_buffer)
        translate_prompt_text = self._prompt_text(translate_widgets.prompt_buffer)
        shared_style_rules_text = (
            self._prompt_text(self._shared_style_rules_buffer)
            if self._shared_style_rules_buffer is not None
            else self._shared_style_rules_text
        )
        proof_settings = ProofreadSettings(
            api_url=proof_widgets.api_url_row.get_text().strip(),
            model_id=proof_widgets.model_row.get_text().strip(),
            api_key=proof_widgets.api_key_row.get_text().strip(),
            prompt=proof_prompt_text.strip() or DEFAULT_PROMPT,
            disable_reasoning=proof_widgets.disable_reasoning_row.get_active(),
            suggestions_json_path=self._proof_suggestions_json_path,
        )
        spelling_settings = SpellingStyleSettings(
            api_url=spelling_widgets.api_url_row.get_text().strip(),
            model_id=spelling_widgets.model_row.get_text().strip(),
            api_key=spelling_widgets.api_key_row.get_text().strip(),
            prompt=spelling_prompt_text.strip() or DEFAULT_SPELLINGSTYLE_PROMPT,
            disable_reasoning=spelling_widgets.disable_reasoning_row.get_active(),
        )
        citation_validator_settings = CitationValidatorSettings(
            api_url=citation_validator_widgets.api_url_row.get_text().strip(),
            model_id=citation_validator_widgets.model_row.get_text().strip(),
            api_key=citation_validator_widgets.api_key_row.get_text().strip(),
            prompt=citation_validator_prompt_text.strip() or DEFAULT_CITATION_NUMBER_VALIDATOR_PROMPT,
            disable_reasoning=citation_validator_widgets.disable_reasoning_row.get_active(),
        )
        thesaurus_settings = ThesaurusSettings(
            api_url=thesaurus_widgets.api_url_row.get_text().strip(),
            model_id=thesaurus_widgets.model_row.get_text().strip(),
            api_key=thesaurus_widgets.api_key_row.get_text().strip(),
            prompt=thesaurus_prompt_text.strip() or DEFAULT_THESAURUS_PROMPT,
            disable_reasoning=thesaurus_widgets.disable_reasoning_row.get_active(),
        )
        reference_settings = ReferenceSettings(
            api_url=reference_widgets.api_url_row.get_text().strip(),
            model_id=reference_widgets.model_row.get_text().strip(),
            api_key=reference_widgets.api_key_row.get_text().strip(),
            tavily_api_key=(reference_widgets.tavily_api_key_row.get_text().strip() if reference_widgets.tavily_api_key_row else ""),
            prompt=reference_prompt_text.strip() or DEFAULT_REFERENCE_PROMPT,
            disable_reasoning=reference_widgets.disable_reasoning_row.get_active(),
        )
        ask_settings = AskSettings(
            api_url=ask_widgets.api_url_row.get_text().strip(),
            model_id=ask_widgets.model_row.get_text().strip(),
            api_key=ask_widgets.api_key_row.get_text().strip(),
            tavily_api_key=(ask_widgets.tavily_api_key_row.get_text().strip() if ask_widgets.tavily_api_key_row else ""),
            prompt=ask_prompt_text.strip() or DEFAULT_ASK_PROMPT,
            disable_reasoning=ask_widgets.disable_reasoning_row.get_active(),
        )
        improve1_settings = Improve1Settings(
            api_url=improve_widgets.api_url_row.get_text().strip(),
            model_id=improve_widgets.model_row.get_text().strip(),
            api_key=improve_widgets.api_key_row.get_text().strip(),
            prompt=improve_prompt_text.strip() or DEFAULT_IMPROVE_PROMPT,
            disable_reasoning=improve_widgets.disable_reasoning_row.get_active(),
        )
        improve2_settings = Improve2Settings(
            api_url=rephrase_widgets.api_url_row.get_text().strip(),
            model_id=rephrase_widgets.model_row.get_text().strip(),
            api_key=rephrase_widgets.api_key_row.get_text().strip(),
            prompt=rephrase_prompt_text.strip() or DEFAULT_IMPROVE2_PROMPT,
            disable_reasoning=rephrase_widgets.disable_reasoning_row.get_active(),
        )
        combine_cites_settings = CombineCitesSettings(
            api_url=combine_widgets.api_url_row.get_text().strip(),
            model_id=combine_widgets.model_row.get_text().strip(),
            api_key=combine_widgets.api_key_row.get_text().strip(),
            prompt=combine_prompt_text.strip() or DEFAULT_COMBINE_CITES_PROMPT,
            disable_reasoning=combine_widgets.disable_reasoning_row.get_active(),
        )
        shorten_settings = ShortenSettings(
            api_url=shorten_widgets.api_url_row.get_text().strip(),
            model_id=shorten_widgets.model_row.get_text().strip(),
            api_key=shorten_widgets.api_key_row.get_text().strip(),
            prompt=shorten_prompt_text.strip() or DEFAULT_SHORTEN_PROMPT,
            disable_reasoning=shorten_widgets.disable_reasoning_row.get_active(),
        )
        introduction_settings = IntroductionSettings(
            api_url=intro_widgets.api_url_row.get_text().strip(),
            model_id=intro_widgets.model_row.get_text().strip(),
            api_key=intro_widgets.api_key_row.get_text().strip(),
            prompt=intro_prompt_text.strip() or DEFAULT_INTRO_PROMPT,
            disable_reasoning=intro_widgets.disable_reasoning_row.get_active(),
        )
        introduction_reply_settings = IntroductionSettings(
            api_url=intro_reply_widgets.api_url_row.get_text().strip(),
            model_id=intro_reply_widgets.model_row.get_text().strip(),
            api_key=intro_reply_widgets.api_key_row.get_text().strip(),
            prompt=intro_reply_prompt_text.strip() or DEFAULT_INTRO_REPLY_PROMPT,
            disable_reasoning=intro_reply_widgets.disable_reasoning_row.get_active(),
        )
        conclusion_settings = ConclusionSettings(
            api_url=conclusion_widgets.api_url_row.get_text().strip(),
            model_id=conclusion_widgets.model_row.get_text().strip(),
            api_key=conclusion_widgets.api_key_row.get_text().strip(),
            prompt=conclusion_prompt_text.strip() or DEFAULT_CONCLUSION_PROMPT,
            disable_reasoning=conclusion_widgets.disable_reasoning_row.get_active(),
        )
        concl_no_issues_settings = ConclNoIssuesSettings(
            api_url=concl_no_issues_widgets.api_url_row.get_text().strip(),
            model_id=concl_no_issues_widgets.model_row.get_text().strip(),
            api_key=concl_no_issues_widgets.api_key_row.get_text().strip(),
            prompt=concl_no_issues_prompt_text.strip() or DEFAULT_CONCL_NO_ISSUES_PROMPT,
            disable_reasoning=concl_no_issues_widgets.disable_reasoning_row.get_active(),
        )
        topic_sentence_settings = TopicSentenceSettings(
            api_url=topic_sentence_widgets.api_url_row.get_text().strip(),
            model_id=topic_sentence_widgets.model_row.get_text().strip(),
            api_key=topic_sentence_widgets.api_key_row.get_text().strip(),
            prompt=topic_sentence_prompt_text.strip() or DEFAULT_TOPIC_SENTENCE_PROMPT,
            disable_reasoning=topic_sentence_widgets.disable_reasoning_row.get_active(),
        )
        concl_section_settings = ConclSectionSettings(
            api_url=concl_section_widgets.api_url_row.get_text().strip(),
            model_id=concl_section_widgets.model_row.get_text().strip(),
            api_key=concl_section_widgets.api_key_row.get_text().strip(),
            prompt=concl_section_prompt_text.strip() or DEFAULT_CONCL_SECTION_PROMPT,
            disable_reasoning=concl_section_widgets.disable_reasoning_row.get_active(),
        )
        translate_settings = TranslateSettings(
            api_url=translate_widgets.api_url_row.get_text().strip(),
            model_id=translate_widgets.model_row.get_text().strip(),
            api_key=translate_widgets.api_key_row.get_text().strip(),
            prompt=translate_prompt_text.strip() or DEFAULT_TRANSLATE_PROMPT,
            disable_reasoning=translate_widgets.disable_reasoning_row.get_active(),
        )
        editor_action_profile_defaults = dict(self._editor_action_profile_defaults)
        dropdowns_by_action_key: dict[str, Gtk.DropDown] = {}
        for widgets in self._prompt_editors.values():
            if not widgets.default_profile_dropdowns:
                continue
            dropdowns_by_action_key.update(widgets.default_profile_dropdowns)
        dropdowns_by_action_key.update(self._choices_profile_dropdowns)
        for key in PROFILE_BACKED_COMMAND_KEYS:
            dropdown = dropdowns_by_action_key.get(key)
            if dropdown is None:
                continue
            selected = int(dropdown.get_selected())
            if selected <= 0:
                editor_action_profile_defaults[key] = None
                continue
            selected_index = selected - 1
            if selected_index < 0 or selected_index >= len(MODEL_PROFILE_IDS):
                editor_action_profile_defaults[key] = None
                continue
            editor_action_profile_defaults[key] = MODEL_PROFILE_IDS[selected_index]
        editor_pinned_action_ids = self._current_editor_pinned_actions()
        text_draft_pinned_action_ids = self._current_text_draft_pinned_actions()
        text_draft_external_actions = self._collect_text_draft_external_actions_from_settings()
        if text_draft_external_actions is None:
            return
        self._on_save(
            model_profiles,
            proof_settings,
            spelling_settings,
            text_draft_spelling_prompt_text.strip() or DEFAULT_TEXT_DRAFT_SPELLINGSTYLE_PROMPT,
            citation_validator_settings,
            improve1_settings,
            improve2_settings,
            combine_cites_settings,
            thesaurus_settings,
            reference_settings,
            ask_settings,
            shorten_settings,
            introduction_settings,
            introduction_reply_settings,
            conclusion_settings,
            concl_no_issues_settings,
            topic_sentence_settings,
            concl_section_settings,
            translate_settings,
            shared_style_rules_text,
            editor_action_profile_defaults,
            editor_pinned_action_ids,
            text_draft_pinned_action_ids,
            self._text_draft_template_dir,
            text_draft_external_actions,
            self._libreoffice_python_path,
            self._concordance_file_path,
        )
        self._parent_window._show_toast("Settings saved.")

    def _on_prompt_row_selected(self, _listbox: Gtk.ListBox, row: Gtk.ListBoxRow | None) -> None:
        if not row:
            return
        key = self._prompt_row_keys.get(row)
        if key:
            self._prompt_stack.set_visible_child_name(key)

    def _build_style_rules_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_vexpand(True)

        title_label = Gtk.Label(label="Style Rules", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        subtitle = Gtk.Label(
            label=(
                f"These rules are inserted anywhere a prompt contains {STYLE_RULES_TOKEN}. "
                "Edit them here to change all token-based prompts at once."
            ),
            xalign=0,
        )
        subtitle.add_css_class("caption")
        subtitle.add_css_class("dim-label")
        subtitle.set_wrap(True)
        subtitle.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(subtitle)

        prompt_section = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        prompt_section.set_hexpand(True)
        prompt_section.set_vexpand(True)
        prompt_label = Gtk.Label(label="Shared Rules", xalign=0)
        prompt_label.add_css_class("dim-label")
        prompt_section.append(prompt_label)
        prompt_scroller, buffer = self._build_prompt_editor(self._shared_style_rules_text or DEFAULT_SHARED_STYLE_RULES)
        self._shared_style_rules_buffer = buffer
        prompt_section.append(prompt_scroller)
        page_box.append(prompt_section)

        page = Gtk.ScrolledWindow()
        page.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        page.set_hexpand(True)
        page.set_vexpand(True)
        page.set_child(page_box)
        return page

    def _build_prompt_page(
        self,
        key: str,
        title: str,
        settings: (
            ProofreadSettings
            | SpellingStyleSettings
            | CitationValidatorSettings
            | Improve1Settings
            | Improve2Settings
            | CombineCitesSettings
            | ThesaurusSettings
            | ReferenceSettings
            | AskSettings
            | ShortenSettings
            | IntroductionSettings
            | ConclusionSettings
            | ConclNoIssuesSettings
            | TopicSentenceSettings
            | ConclSectionSettings
            | TranslateSettings
        ),
        default_prompt: str,
    ) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_vexpand(True)

        title_label = Gtk.Label(label=title, xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        uses_shared_profiles = key in PROFILE_BACKED_COMMAND_KEYS

        api_url_row = Adw.EntryRow(title="API URL")
        api_url_row.set_text(settings.api_url)

        model_row = Adw.EntryRow(title="Model ID (optional)")
        model_row.set_text(settings.model_id)

        api_key_row = Adw.PasswordEntryRow(title="API Key")
        api_key_row.set_text(settings.api_key)

        tavily_api_key_row = None
        disable_reasoning_row = Adw.SwitchRow(title="Disable reasoning")
        disable_reasoning_row.set_active(bool(settings.disable_reasoning))
        default_profile_dropdowns: dict[str, Gtk.DropDown] = {}

        if uses_shared_profiles:
            profile_group = Adw.PreferencesGroup(title="Default Model Profile")
            profile_group.add_css_class("list-stack")
            profile_group.set_hexpand(True)

            profile_row = Adw.PreferencesRow()
            profile_row.set_selectable(False)
            profile_row.set_activatable(False)
            profile_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
            profile_box.set_margin_top(10)
            profile_box.set_margin_bottom(10)
            profile_box.set_margin_start(12)
            profile_box.set_margin_end(12)

            caption_text = "Choose the default profile used by this command. Its reasoning setting will also be used."
            profile_keys = (key,)
            if key == "improve-generated":
                caption_text = (
                    "Choose the default profiles used by Improve Generated and Improve Selected. "
                    "Generated Choices and Selected Choices use the separate Choices Models page."
                )
                profile_keys = (
                    "improve-generated",
                    "improve-selected",
                )

            caption = Gtk.Label(label=caption_text, xalign=0)
            caption.add_css_class("caption")
            caption.add_css_class("dim-label")
            caption.set_wrap(True)
            caption.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
            profile_box.append(caption)

            for profile_key in profile_keys:
                if len(profile_keys) > 1:
                    profile_label = Gtk.Label(
                        label=PROFILE_BACKED_COMMAND_TITLES.get(profile_key, profile_key.replace("-", " ").title()),
                        xalign=0,
                    )
                    profile_label.add_css_class("caption")
                    profile_label.add_css_class("dim-label")
                    profile_box.append(profile_label)

                dropdown = Gtk.DropDown(model=self._profile_dropdown_model(include_unset=True))
                selected_profile_key = self._editor_action_profile_defaults.get(profile_key)
                selected_index = (
                    MODEL_PROFILE_IDS.index(selected_profile_key) + 1
                    if selected_profile_key in MODEL_PROFILE_IDS
                    else 0
                )
                dropdown.set_selected(selected_index)
                profile_box.append(dropdown)
                default_profile_dropdowns[profile_key] = dropdown

            profile_row.set_child(profile_box)
            profile_group.add(profile_row)
            page_box.append(profile_group)
        else:
            credentials_group = Adw.PreferencesGroup(title="Credentials")
            credentials_group.add_css_class("list-stack")
            credentials_group.set_hexpand(True)
            page_box.append(credentials_group)

            credentials_group.add(api_url_row)
            credentials_group.add(model_row)
            credentials_group.add(api_key_row)

            if isinstance(settings, (ReferenceSettings, AskSettings)):
                tavily_api_key_row = Adw.PasswordEntryRow(title="Tavily API Key")
                tavily_api_key_row.set_text(settings.tavily_api_key)
                credentials_group.add(tavily_api_key_row)

            credentials_group.add(disable_reasoning_row)

        prompt_section = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        prompt_section.set_hexpand(True)
        prompt_section.set_vexpand(True)
        prompt_label = Gtk.Label(label="Prompt", xalign=0)
        prompt_label.add_css_class("dim-label")
        prompt_section.append(prompt_label)
        prompt_hint = Gtk.Label(label=STYLE_RULES_HINT_TEXT, xalign=0)
        prompt_hint.add_css_class("caption")
        prompt_hint.add_css_class("dim-label")
        prompt_hint.set_wrap(True)
        prompt_hint.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        prompt_section.append(prompt_hint)
        prompt_scroller, buffer = self._build_prompt_editor(settings.prompt or default_prompt)
        prompt_section.append(prompt_scroller)
        page_box.append(prompt_section)

        ask_prompt_buffer = None

        page = Gtk.ScrolledWindow()
        page.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        page.set_hexpand(True)
        page.set_vexpand(True)
        page.set_child(page_box)

        self._prompt_editors[key] = PromptEditorWidgets(
            api_url_row=api_url_row,
            model_row=model_row,
            api_key_row=api_key_row,
            tavily_api_key_row=tavily_api_key_row,
            disable_reasoning_row=disable_reasoning_row,
            prompt_buffer=buffer,
            ask_prompt_buffer=ask_prompt_buffer,
            default_profile_dropdowns=default_profile_dropdowns or None,
        )
        return page

    def _build_text_draft_spellingstyle_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_vexpand(True)

        title_label = Gtk.Label(label="Text Draft SpellingStyle", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        details_group = Adw.PreferencesGroup(title="Shared Credentials")
        details_group.add_css_class("list-stack")
        details_group.set_hexpand(True)
        page_box.append(details_group)

        details_row = Adw.ActionRow(
            title="Uses the main SpellingStyle model profile",
            subtitle=(
                "Text Draft SpellingStyle shares the default model profile selected on the SpellingStyle page. "
                "Only the prompt below is separate."
            ),
        )
        details_row.set_activatable(False)
        details_group.add(details_row)

        prompt_section = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        prompt_section.set_hexpand(True)
        prompt_section.set_vexpand(True)
        prompt_label = Gtk.Label(label="Prompt", xalign=0)
        prompt_label.add_css_class("dim-label")
        prompt_section.append(prompt_label)
        prompt_hint = Gtk.Label(label=STYLE_RULES_HINT_TEXT, xalign=0)
        prompt_hint.add_css_class("caption")
        prompt_hint.add_css_class("dim-label")
        prompt_hint.set_wrap(True)
        prompt_hint.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        prompt_section.append(prompt_hint)
        prompt_scroller, buffer = self._build_prompt_editor(
            self._text_draft_spelling_prompt or DEFAULT_TEXT_DRAFT_SPELLINGSTYLE_PROMPT
        )
        self._text_draft_spelling_prompt_buffer = buffer
        prompt_section.append(prompt_scroller)
        page_box.append(prompt_section)

        page = Gtk.ScrolledWindow()
        page.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        page.set_hexpand(True)
        page.set_vexpand(True)
        page.set_child(page_box)
        return page

    def _build_prompt_editor(self, text: str) -> tuple[Gtk.ScrolledWindow, Gtk.TextBuffer]:
        scroller = Gtk.ScrolledWindow()
        scroller.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        scroller.set_hexpand(True)
        scroller.set_vexpand(True)
        scroller.set_has_frame(False)

        buffer = Gtk.TextBuffer()
        buffer.set_text(text)
        prompt_view = Gtk.TextView.new_with_buffer(buffer)
        prompt_view.set_wrap_mode(Gtk.WrapMode.WORD_CHAR)
        prompt_view.set_monospace(True)
        prompt_view.set_vexpand(True)
        prompt_view.set_hexpand(True)
        prompt_view.set_top_margin(12)
        prompt_view.set_bottom_margin(12)
        prompt_view.set_left_margin(12)
        prompt_view.set_right_margin(12)
        scroller.set_child(prompt_view)
        return scroller, buffer

    def _prompt_text(self, buffer: Gtk.TextBuffer) -> str:
        start, end = buffer.get_bounds()
        return buffer.get_text(start, end, True)


