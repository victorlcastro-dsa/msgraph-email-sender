import asyncio

from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.uix.label import Label
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput

from app.controller.home_controller import HomeController


class MainScreen(BoxLayout):
    """
    MainScreen class handles the user interface for the email sender application.

    Attributes:
        controller (HomeController): The controller to handle the business logic.
        title_label (Label): The label for the title.
        file_box (BoxLayout): The layout for file selection.
        select_button (Button): The button to select the Excel file.
        file_label (Label): The label to display the selected file.
        sender_input (TextInput): The input field for the sender email.
        send_button (Button): The button to send emails.
        status_label (Label): The label to display the status message.
        body_spinner (Spinner): The spinner to select the email body.
        formats (Dict[str, Dict[str, str]]): The dictionary to store the formats for each body part.
        hyperlink_checkboxes (Dict[str, bool]): The dictionary to store the hyperlink status for each body part.
        line_breaks (Dict[str, int]): The dictionary to store the line breaks for each body part.
        preview_label (Label): The label to preview the selected format.
    """

    def __init__(self, **kwargs) -> None:
        """
        Initializes the MainScreen instance with default values.
        """
        super().__init__(**kwargs)
        self.orientation = "vertical"
        self.spacing = 15
        self.padding = [30, 30, 30, 30]

        self.controller = HomeController()
        self.formats = {}
        self.hyperlink_checkboxes = {}
        self.line_breaks = {}

        self.title_label = Label(
            text="DISPARADOR OUTLOOK",
            font_size="22sp",
            bold=True,
            size_hint_y=None,
            height=40,
        )
        self.add_widget(self.title_label)

        self.file_box = BoxLayout(
            orientation="horizontal", size_hint_y=None, height=50, spacing=10
        )
        self.select_button = Button(
            text="SELECIONAR PLANILHA", size_hint_x=0.7, on_press=self.open_file_dialog
        )
        self.file_label = Label(
            text="No file selected",
            size_hint_x=1.3,
            color=(0.5, 0.5, 0.5, 1),
        )
        self.file_box.add_widget(self.select_button)
        self.file_box.add_widget(self.file_label)
        self.add_widget(self.file_box)

        self.sender_input = TextInput(
            hint_text="Remetente",
            size_hint_y=None,
            height=50,
            multiline=False,
            padding=[10, 10],
            font_size="16sp",
            background_color=(0.2, 0.2, 0.2, 1),
            foreground_color=(1, 1, 1, 1),
        )
        self.add_widget(self.sender_input)

        self.body_spinner = Spinner(
            text="Selecione",
            values=[],
            size_hint_y=None,
            height=50,
            background_color=(0.2, 0.2, 0.2, 1),
            color=(1, 1, 1, 1),
        )
        self.body_spinner.bind(text=self.on_body_spinner_select)
        self.add_widget(self.body_spinner)

        self.format_buttons = BoxLayout(
            orientation="horizontal", size_hint_y=None, height=50, spacing=10
        )
        self.bold_button = Button(
            text="B",
            font_size="20sp",
            bold=True,
            on_press=lambda _: self.apply_format("Negrito"),
            background_color=(0.2, 0.6, 1, 1),
            color=(1, 1, 1, 1),
        )
        self.italic_button = Button(
            text="I",
            font_size="20sp",
            italic=True,
            on_press=lambda _: self.apply_format("Itálico"),
            background_color=(0.2, 0.6, 1, 1),
            color=(1, 1, 1, 1),
        )
        self.underline_button = Button(
            text="U",
            font_size="20sp",
            underline=True,
            on_press=lambda _: self.apply_format("Sublinhado"),
            background_color=(0.2, 0.6, 1, 1),
            color=(1, 1, 1, 1),
        )
        self.increase_font_button = Button(
            text="A+",
            font_size="20sp",
            on_press=lambda _: self.apply_format("Aumentar Fonte"),
            background_color=(0.2, 0.6, 1, 1),
            color=(1, 1, 1, 1),
        )
        self.format_buttons.add_widget(self.bold_button)
        self.format_buttons.add_widget(self.italic_button)
        self.format_buttons.add_widget(self.underline_button)
        self.format_buttons.add_widget(self.increase_font_button)
        self.add_widget(self.format_buttons)

        self.link_break_box = BoxLayout(
            orientation="horizontal", size_hint_y=None, height=50, spacing=10
        )
        self.hyperlink_label = Label(
            text="Hiperlink",
            size_hint_x=0.5,
            color=(1, 1, 1, 1),
        )
        self.hyperlink_checkbox = CheckBox()
        self.hyperlink_checkbox.bind(active=self.toggle_hyperlink)
        self.link_break_box.add_widget(self.hyperlink_label)
        self.link_break_box.add_widget(self.hyperlink_checkbox)

        self.line_break_label = Label(
            text="Quebras de Linha:",
            size_hint_x=0.5,
            color=(1, 1, 1, 1),
        )
        self.line_break_spinner = Spinner(
            text="0",
            values=["0", "1", "2", "3", "4", "5"],
            size_hint_y=None,
            height=50,
            background_color=(0.2, 0.2, 0.2, 1),
            color=(1, 1, 1, 1),
        )
        self.line_break_spinner.bind(text=self.update_line_breaks)
        self.link_break_box.add_widget(self.line_break_label)
        self.link_break_box.add_widget(self.line_break_spinner)
        self.add_widget(self.link_break_box)

        self.preview_label = Label(
            text="Pré-visualização",
            size_hint_y=None,
            height=50,
            font_size="16sp",
            color=(1, 1, 1, 1),
        )
        self.add_widget(self.preview_label)

        self.send_button = Button(
            text="Enviar",
            size_hint_y=None,
            height=50,
            bold=True,
            background_color=(0.2, 0.6, 1, 1),
            color=(1, 1, 1, 1),
            on_press=self.schedule_send_emails,
        )
        self.add_widget(self.send_button)

        self.status_label = Label(text="", size_hint_y=None, height=50, color=(1, 1, 1, 1))
        self.add_widget(self.status_label)

    def open_file_dialog(self, instance: Button) -> None:
        """
        Opens the file dialog to select an Excel file.

        Args:
            instance (Button): The button instance that triggered this method.
        """
        self.file_label.text = self.controller.open_file_dialog()
        self.update_body_spinner()

    def update_body_spinner(self) -> None:
        """
        Updates the body spinner values based on the selected Excel file.
        """
        email_data = self.controller._process_excel()
        body_columns = [
            f"CORPO E-MAIL {i + 1}" for i in range(len(email_data["bodies"][0]))
        ]
        self.body_spinner.values = body_columns

    def on_body_spinner_select(self, spinner, value) -> None:
        """
        Handles the selection of a new email body.

        Args:
            spinner (Spinner): The spinner instance.
            value (str): The selected value.
        """
        self.update_preview()
        self.hyperlink_checkbox.active = self.hyperlink_checkboxes.get(value, False)
        self.line_break_spinner.text = str(self.line_breaks.get(value, 0))
        self.update_format_buttons()

    def apply_format(self, format_type: str) -> None:
        """
        Applies the selected format to the selected email body.

        Args:
            format_type (str): The format type to apply.
        """
        selected_body = self.body_spinner.text
        if selected_body != "Selecione":
            if selected_body not in self.formats:
                self.formats[selected_body] = {}
            self.formats[selected_body][format_type] = not self.formats[
                selected_body
            ].get(format_type, False)
            self.update_preview()
            self.update_format_buttons()

    def toggle_hyperlink(self, checkbox, value) -> None:
        """
        Toggles the hyperlink status for the selected email body.

        Args:
            checkbox (CheckBox): The checkbox instance.
            value (bool): The value of the checkbox.
        """
        selected_body = self.body_spinner.text
        if selected_body != "Selecione":
            self.hyperlink_checkboxes[selected_body] = value
            self.update_preview()

    def update_line_breaks(self, spinner, value) -> None:
        """
        Updates the line breaks for the selected email body.

        Args:
            spinner (Spinner): The spinner instance.
            value (str): The selected value.
        """
        selected_body = self.body_spinner.text
        if selected_body != "Selecione":
            self.line_breaks[selected_body] = int(value)
            self.update_preview()

    def update_preview(self, *args) -> None:
        """
        Updates the preview label to reflect the selected format.
        """
        selected_body = self.body_spinner.text
        format_info = self.formats.get(selected_body, {})
        is_hyperlink = self.hyperlink_checkboxes.get(selected_body, False)
        line_breaks = self.line_breaks.get(selected_body, 0)

        preview_text = "Pré-visualização"
        if "Negrito" in format_info:
            preview_text = f"<b>{preview_text}</b>"
        if "Itálico" in format_info:
            preview_text = f"<i>{preview_text}</i>"
        if "Sublinhado" in format_info:
            preview_text = f"<u>{preview_text}</u>"
        if "Aumentar Fonte" in format_info:
            preview_text = f"<font size='5'>{preview_text}</font>"
        if is_hyperlink:
            preview_text = f"<a href='#'>{preview_text}</a>"
        preview_text += "<br>" * line_breaks

        self.preview_label.text = preview_text

    def update_format_buttons(self) -> None:
        """
        Updates the visual feedback of the format buttons based on the selected email body.
        """
        selected_body = self.body_spinner.text
        format_info = self.formats.get(selected_body, {})

        self.bold_button.background_color = (
            (0.2, 0.6, 1, 1) if format_info.get("Negrito") else (1, 1, 1, 1)
        )
        self.italic_button.background_color = (
            (0.2, 0.6, 1, 1) if format_info.get("Itálico") else (1, 1, 1, 1)
        )
        self.underline_button.background_color = (
            (0.2, 0.6, 1, 1) if format_info.get("Sublinhado") else (1, 1, 1, 1)
        )
        self.increase_font_button.background_color = (
            (0.2, 0.6, 1, 1) if format_info.get("Aumentar Fonte") else (1, 1, 1, 1)
        )

    def schedule_send_emails(self, instance: Button) -> None:
        """
        Schedules the sending of emails.

        Args:
            instance (Button): The button instance that triggered this method.
        """
        sender_email = self.sender_input.text
        formats = {
            body: {
                "formats": fmt,
                "hyperlink": self.hyperlink_checkboxes.get(body, False),
                "line_breaks": self.line_breaks.get(body, 0),
            }
            for body, fmt in self.formats.items()
        }
        asyncio.run(self.controller.send_emails(sender_email, formats))
        self.status_label.text = self.controller.status_message
