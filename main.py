from gui import ui
from features import utils
import logging

LOGGING_FORMAT = (
    "[(%(asctime)s %(filename)s->%(funcName)s():%(lineno)s)] %(levelname)s: %(message)s"
)


def main():
    for handler in logging.root.handlers:
        logging.root.removeHandler(handler)

    logging.basicConfig(
        filename=utils.get_log_path(),
        filemode="a",
        format=LOGGING_FORMAT,
        level=logging.INFO,
        force=True,
    )

    logging.log(logging.INFO, "---Initializing GUI---")
    ui.startApp()


main()
