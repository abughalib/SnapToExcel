import platform
import logging


if platform.system() == "Windows":
    from winotify import notify

if platform.system() == "Linux":
    import notify2


class Notification:

    @staticmethod
    def send_notification_success(
        heading: str, message: str, location: str, duration=1
    ):
        if platform.system() == "Windows":
            logging.log(
                logging.INFO,
                f"Sending Toas Message, Heading: {heading}, message: \
                    {message}, location: {location}, duration: {duration}",
            )
            notify(heading, message, on_click=f"{location}")
        elif platform.system() == "Linux":
            logging.log(
                logging.INFO,
                f"Sending Notification, Heading: {heading}, message: \
                    {message}, location: {location}, duration: {duration}",
            )
            notify2.init("SnapToExcel")
            notify2.Notification(heading, message, icon=f"{location}").set_timeout(
                duration
            )

    @staticmethod
    def send_notification_error(heading: str, message: str, duration=1):
        if platform.system() == "Windows":
            logging.log(
                logging.INFO,
                f"Sending Toas Message, Heading: {heading}, message: \
                    {message}, duration: {duration}",
            )
            notify(heading, message)

        elif platform.system() == "Linux":
            logging.log(
                logging.INFO,
                f"Sending Notification, Heading: {heading}, message: \
                    {message}, duration: {duration}",
            )
            notify2.init("SnapToExcel")
            notify2.Notification(heading, message).show()
