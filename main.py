import logging

from kivy.app import App

from app.views.main_screen import MainScreen
from app.config.settings import Settings


class HomeApp(App):
    """
    HomeApp class initializes and runs the Kivy application.
    """
    settings = Settings()
    title = settings.APP_TITLE

    def build(self) -> MainScreen:
        """
        Builds the main screen of the application.

        Returns:
            MainScreen: The main screen of the application.
        """
        self.icon = self.settings.APP_ICON_PATH
        return MainScreen()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    HomeApp().run()
